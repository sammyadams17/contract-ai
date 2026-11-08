"""The planner: an agent divides the form into the pieces the drafters take.

Where a piece begins and ends, what it is called and which playbook section governs it are
read from the document by the planner. Which of the term sheet's sections each piece needs is
named in a second pass, one small agent per piece, with only that piece in front of it. A piece's
rules are not chosen: rules_for gives it its own section, each section whose "Read also by:" line
names it, and every General bullet. Code here knows only the render's syntax (heading hashes, the
footnote and structure tails) and checks that the plan is whole and honest: every line in exactly
one piece, every section a real one, every term sheet section one the term sheet has.
"""

from __future__ import annotations

import pathlib
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import ClassVar

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field, model_validator

from drafting_engine.draft.paragraphs import paragraph_text
from drafting_engine.llm.agent import AgentDidNotCompleteError, run_agent

BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "planner.md"
INSTRUCTIONS_BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "piece-instructions.md"
HOMES_BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "termsheet-homes.md"
TERMS_BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "termsheet-terms.md"


def brief() -> str:
    return BRIEF_FILE.read_text()


def instructions_brief() -> str:
    return INSTRUCTIONS_BRIEF_FILE.read_text()


@dataclass
class Line:
    n: int
    level: int  # heading level from the render's hashes; 0 for body or a row
    text: str  # as the document holds it
    raw: str  # the render's line, for the drafter


@dataclass
class Unit:
    title: str
    text: str  # the render's lines of this piece
    section: str  # the playbook section that governs it
    rules: list[str] = field(default_factory=list)  # its playbook section, then General bullets
    termsheet_sections: list[int] = field(default_factory=list)  # the term sheet's sections it needs
    client_particulars: bool = False  # whether the form asks it for the client's own details
    # for each term sheet section several pieces read: [section, term, the piece that writes it]
    term_homes: list[list] = field(default_factory=list)


def lines_of(markdown: str) -> list[Line]:
    """The document's paragraphs and rows as numbered lines, before the render's footnote list
    and read-only structure tail. The render's syntax, not a reading of content."""
    out: list[Line] = []
    pending_level = 0  # a heading whose text the render puts on the next line
    for raw in markdown.split("\n"):
        if raw.strip() == "<!-- READONLY_BOUNDARY_START -->" or re.fullmatch(
            r"#*\s*(Footnotes|Endnotes)\s*", raw
        ):
            break
        if not raw.strip() or re.fullmatch(r"[\s|:-]+", raw):
            continue
        level = len(raw) - len(raw.lstrip("#")) if raw.startswith("#") else pending_level
        text = paragraph_text(raw)
        if not text:
            pending_level = level
            continue
        pending_level = 0
        out.append(Line(len(out), level, text, raw))
    return out


def document_being_edited(precedent: bool) -> str:
    """One sentence naming the document in front of the drafter."""
    if precedent:
        return (
            "THE DOCUMENT YOU ARE EDITING is a precedent: a signed contract from an earlier deal,"
            " whose own terms give way to this term sheet's."
        )
    return "THE DOCUMENT YOU ARE EDITING is the client's template."


def playbook_sections(playbook: pathlib.Path) -> dict[str, str]:
    """The hand-written playbook, section name to text.

    A line "@include general.md" above the first section brings in another playbook file from
    the same folder, in its place: the rules every playbook shares live in one file, so a rule
    learned on one deal reaches every contract type.
    """
    lines: list[str] = []
    for line in playbook.read_text().split("\n"):
        if line.startswith("@include "):
            lines.extend((playbook.parent / line[9:].strip()).read_text().split("\n"))
        else:
            lines.append(line)
    sections: dict[str, str] = {}
    current = ""
    for line in lines:
        if line.startswith("## "):
            current = line[3:].strip()
            # "General (NDA)": a playbook names its contract type so its General is not read
            # as a copy of the shared one, but its bullets are General rules all the same
            if current.startswith("General ("):
                current = "General"
            sections.setdefault(current, "")
        elif current:
            sections[current] += line + "\n"
    return sections


class Piece(BaseModel):
    title: str = Field(description="the form's heading, or the block title the brief describes")
    start: int = Field(description="first line number of the piece")
    end: int = Field(description="last line number of the piece, inclusive")
    section: str = Field(default="", description="the playbook section governing it, by exact name")


class Instructions(BaseModel):
    termsheet_sections: list[int] = Field(
        default_factory=list, description="the term sheet's sections this piece needs, by number"
    )
    client_particulars: bool = Field(
        default=False,
        description="true where the form asks this piece for the client's own details",
    )


def _squash(text: str) -> str:
    return " ".join(text.split())


def _bullets(section_text: str) -> list[str]:
    """A section's bullets as the playbook writes them: each `- ` item with its continuation
    lines. The render's syntax, not a reading of content."""
    out: list[str] = []
    for line in section_text.split("\n"):
        if line.startswith("- "):
            out.append(line[2:].strip())
        elif out and line.startswith("  ") and line.strip():
            out[-1] += " " + line.strip()
    return out


_READ_ALSO_BY = "Read also by:"


def rules_for(book: dict[str, str], section: str) -> list[str]:
    """Every rule a piece governed by `section` is given, the same on every run: its own
    section, each other section whose `Read also by:` line names it, then every General
    bullet. No model chooses: a rule the playbook writes for a piece always reaches it."""
    own = book.get(section, "").strip()
    named = [
        text.strip()
        for name, text in book.items()
        if name not in (section, "General")
        and any(
            section in [n.strip() for n in line[len(_READ_ALSO_BY) :].split(";")]
            for line in text.split("\n")
            if line.startswith(_READ_ALSO_BY)
        )
    ]
    return [t for t in [own, *named] if t] + _bullets(book.get("General", ""))


def divide(
    markdown: str, playbook: pathlib.Path, llm, max_turns: int = 40, *, precedent: bool = False
) -> list[Piece]:
    """Pass one: the planner agent divides the form into pieces, each with its playbook
    section. Raises if the agent runs out of turns."""
    lines = lines_of(markdown)
    names = list(playbook_sections(playbook))

    def outline() -> dict:
        """The form's headings (levels 1 and 2): line number, level, opening words; and how
        many lines the form has in all."""
        return {
            "lines": len(lines),
            "headings": [
                {"n": ln.n, "level": ln.level, "text": ln.text[:90]}
                for ln in lines
                if 0 < ln.level <= 2
            ],
        }

    def read(start: int, end: int) -> list[dict]:
        """The lines from start to end inclusive: number, level, opening words. At most 120
        lines at a time."""
        return [
            {"n": ln.n, "level": ln.level, "text": ln.text[:160]}
            for ln in lines[max(start, 0) : min(end, start + 119) + 1]
        ]

    def sections() -> list[str]:
        """The playbook's section names, one of which each piece takes."""
        return names

    total = len(lines)

    class Plan(BaseModel):
        pieces: list[Piece] = Field(description="every piece of the form, in document order")
        n_lines: ClassVar[int] = total
        section_names: ClassVar[list[str]] = names

        @model_validator(mode="after")
        def whole(self) -> Plan:
            problems = []
            expected = 0
            for p in self.pieces:
                if p.start != expected:
                    problems.append(
                        f"'{p.title}' starts at {p.start}; lines {expected}..{p.start - 1} are "
                        + ("in two pieces" if p.start < expected else "in no piece")
                    )
                if p.end < p.start:
                    problems.append(f"'{p.title}' ends before it starts")
                if self.section_names and p.section and p.section not in self.section_names:
                    problems.append(f"'{p.title}': no playbook section named {p.section!r}")
                expected = p.end + 1
            if expected != self.n_lines:
                problems.append(
                    f"the form has lines 0..{self.n_lines - 1}; the plan ends at {expected - 1}"
                )
            seen: dict[str, int] = {}
            for p in self.pieces:
                seen[p.title] = seen.get(p.title, 0) + 1
            for title, n in seen.items():
                if n > 1:
                    problems.append(
                        f"{title!r} is the title of {n} pieces; every piece needs its own title"
                        " (say which clause or schedule it is, e.g. 'Payment (Schedule 2)')"
                    )
            if problems:
                raise ValueError("; ".join(problems))
            return self

    tools = [
        StructuredTool.from_function(outline),
        StructuredTool.from_function(read),
        StructuredTool.from_function(sections),
    ]
    task = (
        f"{document_being_edited(precedent)}\n\n"
        + f"THE FORM has {total} lines (0..{total - 1}). Divide it into pieces as the brief says. "
        "Start with outline; call sections for the playbook's section names; read where the "
        "outline does not settle a boundary. Finish with submit_outcome."
    )
    plan = run_agent(
        system_prompt=brief(),
        task_message=task,
        tools=tools,
        final_output_schema=Plan,
        max_iterations=max_turns,
        llm=llm,
    )
    assert isinstance(plan, Plan)
    return plan.pieces


def instruct(
    title: str,
    piece_text: str,
    section_text: str,
    termsheet: list[tuple[int, str, str]],
    llm,
    max_turns: int = 30,
) -> Instructions:
    """Pass two, for one piece: an agent with the piece and its playbook section in front of
    it names the term sheet's sections the drafter needs. The rules a piece gets are not its to
    choose (see rules_for). Raises if the agent runs out of turns."""
    known = {n: (h, t) for n, h, t in termsheet}

    def termsheet_section(n: int) -> dict:
        """One section of the term sheet in full, by number: to see whether it carries a fact for
        this piece."""
        if n not in known:
            return {"error": f"the term sheet has no section {n}"}
        return {"n": n, "heading": known[n][0], "text": known[n][1]}

    class PieceInstructions(Instructions):
        @model_validator(mode="after")
        def honest(self) -> PieceInstructions:
            problems = []
            for n in self.termsheet_sections:
                if n not in known:
                    problems.append(f"the term sheet has no section {n}")
            if problems:
                raise ValueError("; ".join(problems))
            return self

    task = (
        f"THE PIECE: {title}.\n\n{piece_text}\n\n"
        + (f"ITS PLAYBOOK SECTION:\n{section_text}\n\n" if section_text else "")
        + "THE TERM SHEET'S SECTIONS:\n"
        + "\n".join(f"  [{n}] {h}" for n, h, _ in termsheet)
        + "\n\nName the term sheet's sections this piece's drafter needs. Open a section with "
        "termsheet_section where its heading leaves you unsure. Finish with submit_outcome."
    )
    out = run_agent(
        system_prompt=instructions_brief(),
        task_message=task,
        tools=[StructuredTool.from_function(termsheet_section)],
        final_output_schema=PieceInstructions,
        max_iterations=max_turns,
        llm=llm,
    )
    assert isinstance(out, Instructions)
    return out


class Home(BaseModel):
    section: int = Field(description="the term sheet section's number, as listed")
    pieces: list[str] = Field(
        default_factory=list, description="titles of the pieces whose drafters must read it"
    )
    reason: str = Field(default="", description="only where no piece needs it: why, in a few words")


class Homes(BaseModel):
    homes: list[Home] = Field(default_factory=list)


def homes_for(
    unclaimed: list[tuple[int, str, str]], titles: list[str], llm, max_turns: int = 20
) -> dict[int, list[str]]:
    """The term sheet's sections no piece named, each given the pieces that must read it or a
    reason none does. Returns section -> piece titles, for the sections given a home."""
    wanted = {n for n, _, _ in unclaimed}
    known = set(titles)

    class Checked(Homes):
        @model_validator(mode="after")
        def whole(self) -> Checked:
            problems = []
            given = {h.section for h in self.homes}
            for n in sorted(wanted - given):
                problems.append(f"section {n} is given no home and no reason")
            for h in self.homes:
                if h.section not in wanted:
                    problems.append(f"section {h.section} is not one of the sections listed")
                bad = [p for p in h.pieces if p not in known]
                if bad:
                    problems.append(f"no piece titled {bad[0]!r}; copy a title exactly as listed")
                if not h.pieces and not h.reason.strip():
                    problems.append(f"section {h.section}: name pieces, or give the reason none needs it")
            if problems:
                raise ValueError("; ".join(problems))
            return self

    task = (
        "THE PIECES OF THE CONTRACT, by title:\n"
        + "\n".join(f"  - {t}" for t in titles)
        + "\n\nTHE TERM SHEET'S SECTIONS NO PIECE NAMED:\n"
        + "\n\n".join(f"[{n}] {h}\n{t}" for n, h, t in unclaimed)
        + "\n\nGive each a home or a reason. Finish with submit_outcome."
    )
    out = run_agent(
        system_prompt=HOMES_BRIEF_FILE.read_text(),
        task_message=task,
        tools=[],
        final_output_schema=Checked,
        max_iterations=max_turns,
        llm=llm,
    )
    assert isinstance(out, Homes)
    return {h.section: h.pieces for h in out.homes if h.pieces}


EVERY_PIECE = "every piece"  # a rule the term sheet sets for the whole contract, applied wherever it reaches

_DEFINED = re.compile(r'^\**[“"]\**([^”"*]+)\**[”"]\**\s+means\b')
_DEFINED_INLINE = re.compile(r'\([^()“"]{0,12}[“"]\**([^”"*()]{2,60}?)\**[”"]\**\)')


def what_it_holds(text: str) -> str:
    """A piece's clause and sub-clause headings and the terms it defines, in its definitions
    or inline in a clause, in order: what an agent choosing where a term sheet term is written
    needs to see beyond the piece's title."""
    held: list[str] = []
    for line in text.split("\n"):
        line = line.strip()
        m = _DEFINED.match(line)
        if m:
            held.append(m.group(1).strip())
        elif line.startswith(("# ", "## ")):
            held.append(re.sub(r"\{#[^}]*\}|[*_~]", "", line.lstrip("#")).strip())
        held += [t.strip() for t in _DEFINED_INLINE.findall(line)]
    return "; ".join(dict.fromkeys(h for h in held if h))


def _holds_nothing(text: str) -> bool:
    """A piece whose only words beyond a short title line ("SCHEDULE 4") are a bracketed note,
    "[Intentionally Blank]": a clause or schedule the form leaves empty."""

    def words(line: str) -> list[str]:
        return re.sub(r"\[[^\]]*\]|\{#[^}]*\}|[*_~#]", " ", line).split()

    lines = [line for line in text.split("\n") if line.strip()]
    return (
        "[" in text
        and bool(lines)
        and len(words(lines[0])) <= 3
        and not any(words(line) for line in lines[1:])
    )


def headings_above(termsheet: list[tuple[int, str, str]], n: int, depth: int = 5) -> str:
    """The headings of the term sheet sections just before section `n`: where it sits (the
    appendix and Part it belongs to), which its own heading may not say."""
    at = next((i for i, s in enumerate(termsheet) if s[0] == n), 0)
    return " > ".join(" ".join(s[1].split()) for s in termsheet[max(0, at - depth) : at])


class TermHome(BaseModel):
    term: str = Field(description="the term the section sets, in a few words")
    piece: str = Field(description="title of the one piece that writes it, exactly as listed")


class TermHomes(BaseModel):
    terms: list[TermHome] = Field(default_factory=list)


def terms_for(
    section: tuple[int, str, str],
    readers: list[tuple[str, str, str]],
    llm,
    book: dict[str, str] | None = None,
    max_turns: int = 12,
    above: str = "",
) -> list[tuple[str, str]]:
    """The terms one term sheet section sets, each with the one piece, of the several that read
    it (`readers`, each a title, the playbook section governing it and what the piece holds),
    that writes it: (term, piece title). `above` gives the term sheet's headings before it. The
    playbook sections in `book` that govern them are shown, so a home the playbook names is the
    one given."""
    titles = [t for t, _, _ in readers]
    known = set(titles) | {EVERY_PIECE}

    class Checked(TermHomes):
        @model_validator(mode="after")
        def whole(self) -> Checked:
            bad = [t.piece for t in self.terms if t.piece not in known]
            if bad:
                raise ValueError(f"no piece titled {bad[0]!r}; copy a title exactly as listed")
            return self

    n, heading, text = section
    governing = sorted({g for _, g, _ in readers if g and (book or {}).get(g, "").strip()})
    task = (
        (f"THE TERM SHEET'S HEADINGS BEFORE IT: {above}\n\n" if above else "")
        + f"THE TERM SHEET'S SECTION [{n}] {heading}\n{text}\n\nTHE PIECES THAT READ IT, by title,"
        " with the playbook section governing each and the headings and definitions it holds:\n"
        + "\n".join(
            f"  - {t}" + (f"  [{g}]" if g else "") + (f"\n      holds: {h}" if h else "")
            for t, g, h in readers
        )
        + (
            "\n\nTHOSE PLAYBOOK SECTIONS, where they say where a term is written:\n\n"
            + "\n\n".join(f"## {g}\n{(book or {})[g].strip()}" for g in governing)
            if governing
            else ""
        )
        + "\n\nGive each term it sets one piece that writes it. Finish with submit_outcome."
    )
    out = run_agent(
        system_prompt=TERMS_BRIEF_FILE.read_text(),
        task_message=task,
        tools=[],
        final_output_schema=Checked,
        max_iterations=max_turns,
        llm=llm,
    )
    assert isinstance(out, TermHomes)
    return [(t.term, t.piece) for t in out.terms]


def give_terms_homes(
    units: list[Unit],
    termsheet: list[tuple[int, str, str]],
    llm,
    workers: int = 8,
    book: dict[str, str] | None = None,
) -> None:
    """Each term sheet section several pieces read is given, term by term, the one piece that
    writes it; every piece reading the section is told which terms are its own and which
    another piece writes. A section whose agent cannot settle it is left as read. A piece
    holding no provision (a schedule left blank) reads no section: one only it read is given
    a home among the others."""
    blank = [u for u in units if _holds_nothing(u.text)]
    others = [u for u in units if u not in blank]
    orphans = {n for u in blank for n in u.termsheet_sections} - {
        n for u in others for n in u.termsheet_sections
    }
    for u in blank:
        u.termsheet_sections = []
    if orphans and others:
        try:
            homes = homes_for([s for s in termsheet if s[0] in orphans], [u.title for u in others], llm)
        except AgentDidNotCompleteError:
            homes = {}
        for n, titles in sorted(homes.items()):
            for u in others:
                if u.title in titles and n not in u.termsheet_sections:
                    u.termsheet_sections.append(n)
    readers: dict[int, list[tuple[str, str, str]]] = {}
    for u in units:
        for n in u.termsheet_sections:
            readers.setdefault(n, []).append((u.title, u.section, what_it_holds(u.text)))
    shared = [s for s in termsheet if len(readers.get(s[0], [])) > 1]

    def one(s: tuple[int, str, str]) -> tuple[int, list[tuple[str, str]]]:
        try:
            return s[0], terms_for(s, readers[s[0]], llm, book, above=headings_above(termsheet, s[0]))
        except AgentDidNotCompleteError:
            return s[0], []

    with ThreadPoolExecutor(max_workers=workers) as pool:
        settled = dict(pool.map(one, shared))
    for u in units:
        u.term_homes = [
            [n, term, piece]
            for n in u.termsheet_sections
            for term, piece in settled.get(n, [])
        ]


def _opening(bullet: str, words: int = 5) -> str:
    return " ".join(_squash(bullet).split()[:words]).lower()


def refresh_rules(unit: Unit, playbook: pathlib.Path) -> Unit:
    """A saved plan carries the playbook text as it was when the plan was made. Reused, its
    pieces take the playbook as it is now (rules_for): their own section, each section whose
    "Read also by:" line names it, and every General bullet, in their current text. A rule
    reworded since the plan was made reaches the piece as reworded; one removed from the
    playbook does not."""
    unit.rules = rules_for(playbook_sections(playbook), unit.section)
    return unit


def refresh_plan(units: list[Unit], playbook: pathlib.Path) -> list[Unit]:
    """Every unit of a saved plan refreshed, and a General bullet written since the plan was
    made (its opening words in no saved unit) given to every piece: a rule nobody is given is
    a rule the draft breaks."""
    book = playbook_sections(playbook)
    general = _bullets(book.get("General", ""))
    known = {_opening(r) for u in units for r in u.rules}
    fresh = [b for b in general if _opening(b) not in known]
    out = [refresh_rules(u, playbook) for u in units]
    for u in out:
        u.rules = u.rules + [b for b in fresh if b not in u.rules]
    return out


def plan_units(
    markdown: str,
    playbook: pathlib.Path,
    llm,
    termsheet: list[tuple[int, str, str]] | None = None,
    workers: int = 8,
    *,
    precedent: bool = False,
) -> list[Unit]:
    """Both passes: divide the form, then write each piece's instructions in parallel. Returns
    the pieces in document order, each carrying its lines, its rules (rules_for), its term sheet
    sections and whether the client's own details reach it."""
    lines = lines_of(markdown)
    book = playbook_sections(playbook)
    pieces = divide(markdown, playbook, llm, precedent=precedent)
    sections = termsheet or []

    def one(p: Piece) -> Unit:
        text = "\n".join(ln.raw for ln in lines[p.start : p.end + 1])
        own = book.get(p.section, "").strip()
        given = None
        for _ in range(2):
            try:
                given = instruct(p.title, text, own, sections, llm)
                break
            except AgentDidNotCompleteError:
                continue
        if given is None:
            # an agent that cannot settle the piece's sections never costs the draft: the
            # drafter is given the whole term sheet, which holds every fact it could need
            given = Instructions(termsheet_sections=[n for n, _, _ in sections], client_particulars=True)
        return Unit(
            p.title,
            text,
            p.section,
            rules_for(book, p.section),
            given.termsheet_sections,
            given.client_particulars,
        )

    with ThreadPoolExecutor(max_workers=workers) as pool:
        units = list(pool.map(one, pieces))
    # each piece names its own sections, so a section can be named by none and read by no
    # drafter (a right of first refusal standing on its own); those are given a home in one more pass
    claimed = {n for u in units for n in u.termsheet_sections}
    unclaimed = [s for s in sections if s[0] not in claimed]
    if unclaimed:
        try:
            homes = homes_for(unclaimed, [u.title for u in units], llm)
        except AgentDidNotCompleteError:
            homes = {}
        for n, titles in sorted(homes.items()):
            for u in units:
                if u.title in titles and n not in u.termsheet_sections:
                    u.termsheet_sections.append(n)
    give_terms_homes(units, sections, llm, workers, book)
    return units
