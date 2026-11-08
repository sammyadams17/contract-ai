"""The clause agent: one lawyer, one clause, a loop of reading, acting and reading back.

The drafting of a clause is an agent's work, not a pipeline's. The agent is given the clause as
numbered paragraphs, the term sheet's sections routed to it, the deal's facts,
the firm's playbook for this clause on this kind of deal, and tools that address paragraphs by
their number: read them, change them, add after them, remove them, search the rest of the
contract and the term sheet. Every change hands back the paragraph as it now reads, so the agent sees
what it did before it goes on. A check tool reports facts about the clause as it stands: which
paragraphs still hold a bracket, which footnote notes still hang on a paragraph, which term is
defined twice, which paragraph was struck without a reason. It reports; the agent decides. The
agent stops when it has accounted for every rule it was given, or when its budget runs out;
either way its account of the rules and the check go to the lawyer with the draft.

Nothing here decides content by pattern. The paragraphs come from the document's own structure,
and the only refusals are structural: a paragraph that does not exist, a row given the wrong
number of cells, a paragraph already removed.
"""

from __future__ import annotations

import difflib
import json
import pathlib
import re
from dataclasses import asdict
from typing import Literal

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from drafting_engine.draft.paragraphs import _FOOTNOTE, Para, Rewriting, document_text, paragraphs_of
from drafting_engine.draft.planner import EVERY_PIECE, document_being_edited
from drafting_engine.draft.writer import _inserted_text, _key, _letters, _number, shown_comment
from drafting_engine.llm.agent import AgentDidNotCompleteError, run_agent
from drafting_engine.termsheet.record import Record, plain

BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "drafter.md"
PRECEDENT_BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "drafter-precedent.md"


def brief(precedent: bool = False) -> str:
    """The drafter's brief: a file a lawyer can read and edit, not a string in code. A
    precedent has its own: on a template every word the term sheet does not reach stands, and on
    a precedent the old deal's terms are exactly the words the term sheet does not reach."""
    return (PRECEDENT_BRIEF_FILE if precedent else BRIEF_FILE).read_text()


def deal_lines(record: Record) -> str:
    """The deal's shape as given for the run: the facts that decide the form's alternatives."""
    lines = [f"- {k}: {v.get('value', '')}" for k, v in record.structure.items() if v.get("value")]
    return "\n".join(lines) or "(none given)"


def footnotes_of(clean_markdown: str) -> dict[str, str]:
    """The form's footnotes as the render lists them, marker to text: `[^fn-12]: ...`. The
    render's own structure, not a reading of content."""
    notes = {}
    for line in clean_markdown.split("\n"):
        if line.startswith("[^fn-") and "]:" in line:
            marker, _, text = line.partition("]:")
            notes[marker + "]"] = text.strip()
    return notes


def _clean(text: str) -> str:
    """New text as the contract will hold it: the render's footnote markers and heading marks
    are syntax of the view the agent reads, never words of the document."""
    text = _FOOTNOTE.sub("", text)
    text = text.replace("\x00a0", "\u00a0")  # a non-breaking space the agent wrote as NUL and "a0"
    text = re.sub(r"[\u00b9\u00b2\u00b3\u2070-\u2079]+", "", text)  # a footnote number typed as text
    text = re.sub(r"^\s*#+\s*", "", text)
    return " ".join(text.split())


# the render's bold and italic: "**words**", "_words_"
_EMPHASIS = re.compile(r"\*\*([^*\n]+?)\*\*|(?<![A-Za-z0-9_])_([^_\n]+?)_(?![A-Za-z0-9_])")


def _unmarked(text: str, render: str) -> str:
    """New text without the bold and italic marks the piece's render showed: an agent that
    copies a span as the render drew it ("**Permitted Assignee**") means its words."""
    for m in _EMPHASIS.finditer(render):
        text = text.replace(m.group(0), m.group(1) or m.group(2))
    return text


_ITEM = re.compile(r"\(?(?:[0-9]{1,3}|[a-z]{1,2}|[ivxlc]{1,5})[.)]\s", re.I)
_CLOSES = (".", ";", ":", "!", "?")
_COLUMNS = re.compile(r"\S {3,}\S.* {3,}\S")  # a scanned table's row, its columns spaced apart


def _row(line: str) -> bool:
    return " | " in line or bool(_COLUMNS.search(line))


def _furniture(record: Record) -> set[str]:
    """The lines of the term sheet its pages print rather than its text says: a line, not a table
    row, a section's heading or the end of a sentence, that recurs, near word for word, in
    three other sections (a scan's footer, its page number and its misreadings changing from
    page to page). Mirrored conditions, at the start and at the end, repeat each other's lines, once."""
    seen: dict[str, list[tuple[int, str]]] = {}
    for s in record.sections:
        for line in plain(s.text).split("\n")[1:]:
            line = line.strip()
            if len(line.split()) > 1 and not _row(line) and not line.endswith(_CLOSES):
                seen.setdefault(line.split()[0].lower(), []).append((s.index, line))
    out: set[str] = set()
    for group in seen.values():
        for n, line in group:
            others = {
                m
                for m, other in group
                if m != n and difflib.SequenceMatcher(a=line.lower(), b=other.lower()).ratio() >= 0.8
            }
            if len(others) >= 3:
                out.add(line)
    return out


def termsheet_paragraphs(record: Record, sections: list[int]) -> tuple[list[str], list[str]]:
    """The paragraphs of the term sheet's sections, in order, and the page furniture left out. A
    term sheet written in Word gives each paragraph a line. A scan breaks a paragraph into lines
    where it reaches the page's edge, and marks where one ends with a blank line, which it then
    sets between paragraphs of text, not only beside a table: there a line runs on from one
    that reached the edge, unless it opens with an item's letter or number."""
    furniture = _furniture(record)
    known = {s.index: s for s in record.sections}
    paragraphs: list[str] = []
    left_out: list[str] = []
    for n in sections:
        lines = [x.strip() for x in plain(known[n].text).split("\n")]
        scan = any(
            not lines[k] and lines[k - 1] and lines[k + 1]
            and not _row(lines[k - 1]) and not _row(lines[k + 1])
            for k in range(1, len(lines) - 1)
        )
        # a scan breaks a line only where it reaches the page's edge, as its longest lines do
        width = max((len(x) for x in lines if not _row(x)), default=0)
        open_ = False  # whether the last paragraph runs on to the next line
        for line in lines:
            if not line:
                open_ = False
                continue
            if scan and line in furniture:
                left_out.append(line)
                open_ = open_ and not paragraphs[-1].endswith(_CLOSES)
                continue
            if scan and open_ and not _row(line) and not _ITEM.match(line):
                paragraphs[-1] = f"{paragraphs[-1]} {line}"
                open_ = len(line) >= 0.75 * width
                continue
            paragraphs.append(line)
            open_ = scan and not _row(line) and len(line) >= 0.75 * width
    return paragraphs, left_out


def _bare(text: str) -> str:
    """Words alone: case, spacing and punctuation set aside."""
    return re.sub(r"\W", "", text.lower())


def _styled_as(p: Para) -> str:
    """The form paragraph whose style a paragraph written beside `p` takes: `p` itself, or, for
    a paragraph this piece added, the form paragraph it took its own style from, or was
    restyled as since. The form has no paragraph reading as an added one, so the writer could
    not find it by its words."""
    if p.inserted is not None:
        return p.like or p.heading_like or p.text
    return p.original or p.text


def _take_level(p: Para, model: Para) -> None:
    """`p`, a paragraph this piece added, set at the level of `model`, another it added with no
    form style of its own: a heading level, body text, or a limb of the same parent."""
    p.heading, p.heading_like, p.like, p.like_struck = model.heading, "", "", False
    p.beneath, p.beneath_of = model.beneath, model.beneath_of


class ClauseAgent:
    """The working state of one clause and the tools the agent acts with."""

    def __init__(
        self,
        heading: str,
        clause_text: str,
        anchored_text: str,
        contract_text: str,
        record: Record,
        footnotes: dict[str, str],
        contract_units: list[tuple[str, str]],
        numbers: dict[tuple[str, int], str] | None = None,
    ) -> None:
        self.heading = heading
        self.work = Rewriting(
            paras=paragraphs_of(clause_text, anchored_text), clause_text=clause_text
        )
        self.contract_text = contract_text
        self.record = record
        self.footnotes = footnotes
        self.contract_units = contract_units
        self.open_points: list[dict] = []
        self._struck_after: dict[str, str] = {}
        self.trace: list[dict] = []
        self.task = ""  # what the drafter was told; the fixer works from it too
        self.precedent = False  # drafting from a precedent: check lists the old deal's numbers
        self.system_brief = ""  # the brief it drafted on; its fixer works on the same one
        self.rule_check: list[dict] = []  # a fresh read's verdict on each rule, once drafted
        self._next = 0
        self.numbers: dict[str, str] = {}  # paragraph id -> the number Word gives it
        self._line_numbers: list[list[str]] = []  # each unit's lines' numbers
        self.struck: set[str] = set()  # words of form paragraphs the draft has struck elsewhere
        self.current: dict[str, str] = {}  # form paragraphs' words -> as the draft now reads them
        if numbers:
            self._number(numbers)

    def _number(self, numbers: dict[tuple[str, int], str]) -> None:
        """Each paragraph's number in the document `numbers` was read from (the writer's
        clause_numbers), by the line showing it: the line's text, and which occurrence of that
        text it is, counting the lines of every unit before it, as the writer counts them."""

        def numbered(text: str, seen: dict[str, int]) -> list[tuple[str, str]]:
            out = []
            for line in text.split("\n"):
                key = _key(document_text(line))
                rank = seen.get(key, 0)
                seen[key] = rank + 1
                out.append((key, numbers.get((key, rank), "")))
            return out

        seen: dict[str, int] = {}
        before: dict[str, int] = {}
        for title, unit in self.contract_units:
            if title == self.heading:
                before = dict(seen)
            self._line_numbers.append([n for _, n in numbered(unit, seen)])
        own = numbered(self.work.clause_text, before)
        at = 0
        for p in self.work.paras:
            key = _key(p.original)
            line = next((j for j in range(at, len(own)) if own[j][0] == key), None)
            if line is None:
                continue  # a table row spans several lines and is never numbered
            at = line + 1
            if own[line][1]:
                self.numbers[p.id] = own[line][1]

    def renumber(self, numbers: dict[tuple[str, int], str]) -> None:
        """Each paragraph's number as the written draft prints it (`numbers`, the writer's
        clause_numbers of that draft), found by its words as they now read; words the draft
        prints more than once keep the number they had."""
        for p in self.work.paras:
            key = _key(document_text(p.text))
            if (key, 0) in numbers and (key, 1) not in numbers:
                if numbers[(key, 0)]:
                    self.numbers[p.id] = numbers[(key, 0)]
                else:
                    self.numbers.pop(p.id, None)

    def state(self) -> dict:
        """What a later run needs to correct this piece where it stands, never drafting it
        again: every paragraph as it now reads, those struck, the task it was drafted from."""
        return {
            "paras": [asdict(p) for p in self.work.paras],
            "deleted": [asdict(p) for p in self.work.deleted],
            "next": self._next,
            "struck_after": self._struck_after,
            "task": self.task,
        }

    def resume(self, state: dict, edits: list[dict], open_points: list[dict]) -> None:
        """The piece as a saved run left it, from its state(), its edits and open points. A
        row of empty cells the form has and the state does not (saved before such rows were
        listed) is put in where the form has it, with its own id; every saved id stands."""

        def para(d: dict) -> Para:
            return Para(**{**d, "inserted": tuple(d["inserted"]) if d["inserted"] else None})

        fresh = self.work.paras  # as the form reads: empty rows a saved state may not hold
        self.work.paras = [para(d) for d in state["paras"]]
        held = {p.id for p in self.work.paras}
        struck = {d["id"] for d in state["deleted"]}
        for k, row in enumerate(fresh):
            nxt = next((q.id for q in fresh[k + 1 :] if not q.id.startswith("e")), None)
            if row.id.startswith("e") and row.id not in held | struck and nxt in held:
                ids = [p.id for p in self.work.paras]
                self.work.paras.insert(ids.index(nxt), row)
        self.work.deleted = [para(d) for d in state["deleted"]]
        self.work.edits = list(edits)
        self.open_points = list(open_points)
        self._next = state["next"]
        self._struck_after = dict(state["struck_after"])
        self.task = state["task"]

    # -- reading -------------------------------------------------------------------------
    def _para(self, pid: str) -> Para | None:
        return next((p for p in self.work.paras if p.id == pid), None)

    def _model(self, pid: str) -> Para | None:
        """A paragraph a style may be taken from: one standing, or a form paragraph this piece
        struck, whose style the writer still finds."""
        return self._para(pid) or next(
            (p for p in self.work.deleted if p.id == pid and p.inserted is None), None
        )

    def _contract_para(self, pid: str) -> Para | None:
        """A paragraph of another piece of the contract, by the id search_contract shows for it
        ("L4.12": that piece, that line), as a style to take: a Part B whose form paragraphs
        are all body text has no heading or numbered item of its own to name. A clause or
        schedule title reads as level 1, as the piece's own does; a table row is no style."""
        m = re.fullmatch(r"L(\d+)\.(\d+)", pid)
        if m is None or int(m.group(1)) >= len(self.contract_units):
            return None
        title, unit = self.contract_units[int(m.group(1))]
        lines = unit.split("\n")
        j = int(m.group(2))
        if title == self.heading or j >= len(lines) or " | " in lines[j]:
            return None
        text = document_text(lines[j])
        if not text:
            return None
        hashes = re.match(r"\s*(#*)", lines[j])
        level = len(hashes.group(1)) if hashes else 0
        first = next((x for x in lines if document_text(x)), "")
        if lines[j] == first and " ".join(text.split()) == " ".join(document_text(title).split()):
            level = 1
        return Para(id=pid, text=text, heading=level)

    def _styled_para(self, pid: str) -> Para | None:
        """The paragraph a like= or level_of= names: one of this piece, or of another."""
        return self._para(pid) or self._contract_para(pid)

    def _own_title(self, p: Para) -> bool:
        """The piece's clause or schedule heading: its first paragraph, where that is a
        top-level heading or reads as the piece's title (a schedule title the render shows as
        body text). A piece has one title: in a piece of definitions every definition renders
        at the top level, and each is a sibling of the next, not a title. A paragraph of
        another piece (a search_contract hit) at the top level is that clause or schedule's title."""
        if re.fullmatch(r"L\d+\.\d+", p.id) and self._para(p.id) is None:
            return p.heading == 1
        first = next((q for q in self.work.paras if q.inserted is None), None)
        if p is not first:
            return False
        # read on the form's words: a title the drafter has reworded is still the title
        words = p.original or p.text
        return p.heading == 1 or " ".join(words.lower().split()) == " ".join(self.heading.lower().split())

    def _notes_of(self, text: str) -> list[str]:
        return [f"{m}: {self.footnotes[m]}" for m in self.footnotes if m in text]

    def _view(self, p: Para) -> dict:
        number = self.numbers.get(p.id)
        view: dict = {"id": p.id, **({"number": number} if number else {}), "text": p.text}
        if p.heading:
            view["heading"] = p.heading
        if p.cells is not None:
            view["row"] = True
            view["cells"] = len(p.cells)
        if "[" in p.text or "]" in p.text:
            view["has_brackets"] = True
        notes = self._notes_of(p.text)
        if notes:
            view["footnote_notes"] = notes
        return view

    form_bytes: bytes = b""  # the form itself, where the run has it: numbers are read off it
    before_text: str = ""  # the render of the form before this piece, for the writer's anchors

    def _refresh_numbers(self) -> None:
        """Each paragraph's number as Word will print it once this piece's changes are written,
        read off the form with them written in: a limb added or struck renumbers those after it,
        and an added limb has its own. Read again only after the piece changes."""
        if not self.form_bytes or self.__dict__.get("_numbered_at") == len(self.work.edits):
            return
        self.__dict__["_numbered_at"] = len(self.work.edits)
        if not self.work.edits:
            return
        from drafting_engine.draft.writer import clause_numbers, snapshot, write
        from drafting_engine.word.adeu_engine import render_markdown

        try:
            plain, _ = write(self.form_bytes, [snapshot(self.work, self.heading, before_text=self.before_text)])
            self.renumber(clause_numbers(plain, render_markdown(plain, clean_view=True)))
        except Exception:  # the form's numbers stand: a view, never a reason to stop drafting
            return

    def outline(self, within: str = "") -> list[dict]:
        """Every paragraph of the clause: id, opening words, level, whether it is a table row,
        whether it holds brackets, and any footnote note the form attached to it. On a whole
        document, with no `within`, only the headings (levels 1 and 2) are listed; give the id,
        the number or the opening words of a heading as `within` to list everything under it.
        A numbered paragraph shows `number`, the number Word prints for it with this piece's
        changes written, an added limb's own included: it changes again where other pieces add
        or strike clauses, and the form's own cross-references are live fields that follow it."""
        self._refresh_numbers()
        paras = self.work.paras
        if within:
            start = next(
                (
                    i
                    for i, p in enumerate(paras)
                    if p.id == within
                    or self.numbers.get(p.id) == within
                    or p.text.lower().startswith(within.lower()[:60])
                ),
                None,
            )
            if start is None:
                return [{"error": f"no paragraph or heading {within!r}"}]
            depth = paras[start].heading or 99
            end = next(
                (i for i in range(start + 1, len(paras)) if 0 < paras[i].heading <= depth),
                len(paras),
            )
            chosen = paras[start:end]
        elif len(paras) > 400:
            chosen = [p for p in paras if 0 < p.heading <= 2]
        else:
            chosen = paras
        out = []
        for p in chosen:
            v = self._view(p)
            v["text"] = p.text[:90]
            out.append(v)
        return out

    def read(self, ids: list[str]) -> list[dict]:
        """The full text of the paragraphs named, as they now read. A table row shows its cells
        separated by ' | ' and a paragraph break inside a cell as ' // '; a numbered paragraph
        shows its number, as outline does."""
        self._refresh_numbers()
        return [
            self._view(p) if (p := self._para(i)) else {"id": i, "error": "no such paragraph"}
            for i in ids
        ]

    def defined_terms(self, term: str = "") -> list[dict]:
        """Whether the contract defines a term: every line of the whole contract reading `"Term"
        means ...`, with the term and how many times it is defined; `term` narrows the list to
        terms containing those words. A capitalised term with no line here is undefined."""
        counts: dict[str, int] = {}
        for line in self.contract_text.split("\n"):
            head = plain(line).lstrip("[ ").strip().replace("\u201c", '"').replace("\u201d", '"')
            m = re.match(r'"([^"]+)"\s+(?:shall\s+)?(?:means?|has the meaning|have the meaning|bears the meaning)\b', head)
            if m:
                found = m.group(1)
                counts[found] = counts.get(found, 0) + 1
        return [
            {"term": t, "defined": n}
            for t, n in counts.items()
            if not term or term.lower() in t.lower()
        ]

    template_text: str = ""  # the client's template, rendered, where a precedent run has it

    def template_definition(self, term: str) -> dict:
        """The client's template's definition of a term, with its limbs, for a term this contract
        writes but does not define: the words a new definition takes where the term sheet gives
        none of its own."""
        if not self.template_text:
            return {"error": "no template is available for this contract"}
        lines = [plain(x).strip() for x in self.template_text.split("\n")]

        def defines(line: str) -> str:
            head = line.lstrip("[ ").replace("“", '"').replace("”", '"')
            m = re.match(r'"([^"]+)"\s+(?:shall\s+)?(?:means?|has the meaning|have the meaning|includes?)\b', head)
            return m.group(1) if m else ""

        want = " ".join(term.lower().split())
        for k, line in enumerate(lines):
            if " ".join(defines(line).lower().split()) != want:
                continue
            out = [line]
            for nxt in lines[k + 1 :]:
                if defines(nxt) or nxt.startswith("#"):
                    break
                if nxt:
                    out.append(nxt)
            return {"term": defines(line), "text": "\n".join(out)[:3000]}
        return {"error": f"the template does not define {term!r}"}

    def template_clause(self, words: str) -> dict:
        """A clause, sub-clause or limb of the client's template, whole with what lies under it:
        the one whose heading reads as `words`, else the first whose text holds them ("Governing
        Law", "confidential information"). For a provision the precedent lacks, written from
        the template's words."""
        if not self.template_text:
            return {"error": "no template is available for this contract"}
        body = re.split(r"\n#*\s*(?:Footnotes|Endnotes)\s*\n|<!-- READONLY", self.template_text)[0]
        lines = [x for x in body.split("\n") if x.strip()]  # the clauses, before their notes
        level = [len(x) - len(x.lstrip("#")) for x in lines]
        text = [plain(x.lstrip("#")).strip() for x in lines]
        want = _bare(words)  # "_Governing-Law_:" reads as "Governing Law"
        named = [k for k in range(len(lines)) if level[k] and _bare(text[k]) == want]
        found = max(named, key=lambda k: level[k]) if named else next(
            (k for k in range(len(lines)) if want in _bare(text[k])), None
        )
        if found is None:
            return {"error": f"the template has no provision reading {words!r}"}
        end = next(
            (k for k in range(found + 1, len(lines)) if level[k] and level[k] <= (level[found] or 99)),
            len(lines),
        )
        return {"text": "\n".join(t.replace("_", "") if level[found] else t for t in text[found:end])[:4000]}

    def _styled_elsewhere(self, text: str, titles: bool = False) -> tuple[str, bool] | None:
        """A form paragraph of another piece reading exactly as `text`, as (its words as the
        writer finds them, whether Word numbers it); None where none does or it is a clause or
        schedule title, whose style would start a new one, unless `titles`: the piece's own
        title taking a title's style (a schedule heading the form set as a limb)."""
        want = _key(text)
        if not want:
            return None
        for u, (title, unit) in enumerate(self.contract_units):
            if title == self.heading:
                continue
            for j, line in enumerate(unit.split("\n")):
                words = document_text(line)
                if _key(words) != want:
                    continue
                if line.lstrip().startswith("# ") and not titles:
                    return None
                numbered = bool(self._line_numbers and self._line_numbers[u][j])
                return words, numbered
        return None

    def search_contract(self, text: str) -> list[dict]:
        """Paragraphs elsewhere in the contract containing the words given, or numbered as given
        ("12.1"): which clause, the paragraph's number as outline shows it, and the passage as
        the draft now reads it. For defined terms and cross-references. Each hit's `id` may be
        given as like= or level_of= to take that paragraph's style and numbering, where this
        piece has no paragraph styled as the new one should be."""
        want = " ".join(text.lower().split())
        gone = {_key(t) for t in self.struck}
        now = {_key(o): t for o, t in self.current.items()}
        hits = []
        for u, (title, unit) in enumerate(self.contract_units):
            if title == self.heading:
                continue
            for j, line in enumerate(unit.split("\n")):
                key = _key(document_text(line)) if gone or now else ""
                if gone and key in gone:
                    continue  # the draft no longer says it
                if key and key in now:
                    line = now[key]  # another piece has rewritten it: the words the draft has
                number = self._line_numbers[u][j] if self._line_numbers else ""
                if want in " ".join(line.lower().split()) or (number and want == number.lower()):
                    hits.append(
                        {
                            "id": f"L{u}.{j}",
                            "clause": title[:80],
                            **({"number": number} if number else {}),
                            "passage": " ".join(line.split())[:300],
                        }
                    )
                if len(hits) >= 12:
                    return hits
        return hits

    def search_record(self, text: str) -> list[dict]:
        """Rows of the term sheet's record whose term or value contains any of the words given,
        the rows matching most words first."""
        words = [w for w in text.lower().split() if len(w) > 2]
        scored = []
        for i, f in enumerate(self.record.facts):
            hay = f"{f.term} {f.value}".lower()
            score = sum(w in hay for w in words)
            if score:
                scored.append((-score, i, f))
        return [
            {"row": i, "section": f.section, "term": f.term, "value": f.value[:300]}
            for _, i, f in sorted(scored)[:20]
        ]

    def read_termsheet_section(self, index: int) -> dict:
        """A section of the term sheet, whole, a line for each of its paragraphs (as copy_termsheet_text
        counts them: a scan's broken lines joined, its page furniture left out)."""
        for s in self.record.sections:
            if s.index == index:
                return {
                    "index": s.index,
                    "heading": s.heading.strip(),
                    "text": "\n".join(termsheet_paragraphs(self.record, [index])[0]),
                }
        return {
            "error": f"no section {index}",
            "sections": [f"[{s.index}] {s.heading.strip()}" for s in self.record.sections],
        }

    # -- acting --------------------------------------------------------------------------
    def _words(self, text: str) -> str:
        return _unmarked(_clean(text), self.work.clause_text)

    def replace(
        self, id: str, new_text: str, reason: str, termsheet_quote: str = "", like: str = ""
    ) -> dict:
        """Replace the whole text of paragraph `id` with new_text. A table row is written
        whole, cells separated by ' | ', the same number of cells: words fused into one cell
        (a label and its values, two entries) are split by writing each into its own cell, an
        empty cell of the row taking the value. A cell's paragraphs are written ' // ' apart;
        a paragraph left out of the cell is struck, never overwritten with a dash. With `like`, the id of
        another paragraph, `id` also takes that paragraph's style and numbering: a numbered
        limb made an unnumbered closing sentence (like= a body paragraph), body text made a
        numbered item or a heading (like= one), a number typed by hand dropped from the words
        and given by the heading style instead. Where this piece has no paragraph styled as it
        should be (a Part B of body text taking another schedule's headings and items),
        like= the id of a search_contract hit in another piece. like= its own id puts a form paragraph restyled
        earlier back in the form's own style. A table row like= another row takes that row's
        cells: their shading, width and run formatting. new_text may be the words as they stand, to
        change only the style. An empty reason leaves the paragraph's comment as it was: where
        that comment no longer describes the paragraph, rewrite or remove it with note. Returns
        the paragraph as it now reads."""
        p = self._para(id)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        heading = self.__dict__.get("_heading_of", {}).get(id)
        if like and heading and like != heading:
            # copied subheadings, styled as the precedent's numbered headings, were once
            # restyled as list items and printed with list numbers
            return {
                "error": f"{id} reads as the precedent's numbered heading {heading} and keeps its"
                f" style: like={heading}, or leave the style as it is"
            }
        own_style = like == id and p.inserted is None and p.cells is None
        model = (self._model(like) or self._contract_para(like)) if like and not own_style else None
        if like and not own_style and model is None:
            # the outline no longer shows a struck heading, so the drafter named it
            # by its words, which the piece's own title also reads as: the struck one is meant
            model = next(
                (q for q in self.work.deleted if q.inserted is None and _key(q.original) == _key(like)),
                None,
            )
            title = next((q for q in self.work.paras if self._own_title(q)), None)
            if model is None and title is not None and title is not p and _key(title.text) == _key(like):
                return {
                    "error": f"{like!r} reads as this piece's own title: text styled like it starts"
                    " a new clause or schedule and renumbers every one after it. Take like= a"
                    " heading or body paragraph of the piece (a struck form paragraph may be named)."
                }
        elsewhere = (
            self._styled_elsewhere(like, titles=self._own_title(p))
            if like and not own_style and model is None and p.cells is None
            else None
        )
        if like and not own_style and elsewhere is None and (model is None or model is p or (p.cells is None and model.cells is not None)):
            return {
                "error": f"like must be the id of another paragraph (of this piece, or a search_contract hit's), or for"
                f" a table row another row,"
                f" or the exact words of a paragraph of another piece as search_contract shows them"
                f" (a clause or schedule title only for this piece's own title): {like}"
            }
        # a paragraph this piece added with no form style of its own (a limb, body text) is
        # styled by its level alone: one added beside it takes that level (limbs
        # it could not style were otherwise deleted)
        unstyled = model is not None and model.inserted is not None and not (model.heading_like or model.like)
        if unstyled and p.inserted is None:
            return {"error": f"{like} was added by this piece; name the form paragraph it is styled as"}
        if model is not None and self._own_title(model) and not self._own_title(p):
            return {
                "error": f"{like} is a clause or schedule's own title: text"
                " styled like it starts a new clause or schedule and renumbers every one after"
                " it. Take like= a body paragraph of the piece instead (a struck form paragraph"
                " may be named)."
            }
        new_text = self._words(new_text)
        if p.cells is not None:
            given = [c.strip() for c in re.split(r"\s*\|\s*", new_text)]
            if len(given) != len(p.cells):
                return {
                    "error": f"this row has {len(p.cells)} cells; {len(given)} given"
                    " (separate cells with ' | ')"
                }
            # a cell of several paragraphs shows them " // " apart: two of them rewritten
            # without the separator between, the writer would run them together as one
            was = [c.strip() for c in re.split(r"\s*\|\s*", p.original or "")]
            run_together = [
                k
                for k, (old, new) in enumerate(zip(was, given, strict=False))
                if " // " in old
                and " // " not in new
                and sum(bool(_bare(x)) and _bare(x)[:40] in _bare(new) for x in old.split(" // ")) > 1
            ]
            if run_together:
                k = run_together[0]
                return {
                    "error": f"cell {k + 1} of this row holds {was[k].count(' // ') + 1} paragraphs,"
                    " shown ' // ' apart; write the new cell with ' // ' between the paragraphs"
                    " it keeps, changing only the words that change. A paragraph left out of the"
                    " cell is struck."
                }
            new_text = " | ".join(given)  # an empty last cell keeps its separator
        if not new_text:
            return {"error": "empty text: use delete to remove a paragraph"}
        shown = shown_comment(self.work, p)
        p.text = new_text
        if model is not None and unstyled:
            _take_level(p, model)
        elif model is not None:
            p.like = _styled_as(model)
            if p.inserted is not None:
                # it now sits at its model's level: one restyled as body text is no heading
                p.heading = model.heading
            # a struck form paragraph named as the model: the writer must not take a live one
            # reading alike (a schedule title and its first heading)
            p.like_struck = (
                model.like_struck
                if model.inserted is not None
                else any(model is q for q in self.work.deleted)
            )
        elif elsewhere is not None:
            p.like, p.like_struck = elsewhere[0], False
        elif own_style:
            p.like, p.like_struck = "", False
        self.work.edits.append(
            {
                "kind": "modify",
                "para": p.id,
                "old_text": p.original,
                "new_text": new_text,
                "reason": reason,
                "termsheet_quote": termsheet_quote,
                **({"like": p.like} if model is not None or elsewhere is not None else {}),
            }
        )
        if not reason.strip() and shown_comment(self.work, p) != shown:
            # no reason given: the comment the lawyer reads stays exactly as it was, none
            # included, never an older edit's reason coming back
            self.note(p.id, shown[0], shown[1])
        view = self._view(p)
        changed = _form_words_changed(p.original, new_text) if p.inserted is None else []
        if changed:
            view["form_words_changed"] = changed
            view["check"] = (
                "These are the form's words you changed. A change the term sheet or a rule requires "
                "stays; put back exactly, with replace, only a form word changed for neither."
            )
        return view

    def insert_after(
        self,
        id: str,
        paragraphs: list[str],
        reason: str,
        termsheet_quote: str = "",
        heading: bool = False,
        level_of: str = "",
        beneath: bool = False,
    ) -> dict:
        """Insert new paragraphs after paragraph `id`, one string per paragraph, in the form's
        register. After a table row, the new paragraphs go after the table. With beneath=True
        the new paragraphs are sub-paragraphs of `id`, one numbering level below it: the (i),
        (ii), (iii) limbs of a paragraph (a). With heading=True
        the first string is the heading of a new sub-clause (a short title, "Service Credit
        Period", never a sentence) and the rest are its numbered paragraphs. The heading takes
        the level of `level_of`, the id of an existing heading the new one is a sibling of
        ("Payment Date" for a new sub-clause of Payment); without it, the level of the nearest
        shallower paragraph above the anchor. With level_of and no heading, every new paragraph
        is a sibling of `level_of`, at its level and in its style: a new item of a list added
        after the last sub-paragraph of the item before it. level_of may be the id of a
        search_contract hit in another piece, where this one has no paragraph at that level and in
        that style. Returns the new ids."""
        p = self._para(id)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        texts = [self._words(t) for t in paragraphs if t.strip()]
        if not texts:
            return {"error": "nothing to insert"}
        if heading and len(texts) < 2:
            return {
                "error": "a new sub-clause is one call: its heading as the first string and every"
                " paragraph under it after, in order. Call again with the heading and the body"
                " together; never the heading alone or the body first."
            }
        at = self.work.paras.index(p) + 1
        if p.cells is not None and not heading:
            # after a table row, text with the row's number of cells is a new row of that table
            split = [[c.strip() for c in re.split(r"\s*\|\s*", t)] for t in texts]
            if all(len(cells) == len(p.cells) for cells in split):
                rows = []
                for k, (t, cells) in enumerate(zip(texts, split, strict=True)):
                    self.work.edits.append(
                        {
                            "kind": "insert_row",
                            "old_text": p.original or p.text,
                            "new_text": t,
                            "cells": cells,
                            "reason": reason,
                            "termsheet_quote": termsheet_quote,
                        }
                    )
                    self._next += 1
                    q = Para(
                        id=f"n{self._next}",
                        text=t,
                        cells=cells,
                        inserted=(len(self.work.edits) - 1, 0),
                    )
                    self.work.paras.insert(at + k, q)
                    rows.append(self._view(q))
                return {"inserted": rows}
            following = self.work.paras[at] if at < len(self.work.paras) else None
            if following is not None and following.cells is not None:
                n = len(p.cells)
                words = [c for c in split[0] if c]
                cells = words[: n - 1] + [" ".join(words[n - 1 :])] if len(words) >= n else words + [""] * (n - len(words))
                return {
                    "error": f"{id} is a row inside a table, and text after a row goes after the"
                    " whole table. To add a line inside the table, write each string as a row of"
                    f" {n} cells, as {id} has, separated by ' | ', an empty cell left empty"
                    f" (\"{' | '.join(cells)}\"); {len(split[0])} were given."
                }
        sibling = self._styled_para(level_of) if level_of and not heading else None
        if level_of and not heading and sibling is None:
            return {"error": f"level_of must be the id of a paragraph: {level_of}"}
        if sibling is not None and self._own_title(sibling):
            return {
                "error": f"{level_of} is a clause or schedule's own title:"
                " paragraphs at its level start a new clause or schedule and renumber every one"
                " after it. Take level_of= a body paragraph of the piece instead."
            }
        self.work.edits.append(
            {
                "kind": "insert",
                "old_text": p.original or p.text,
                "new_text": "\n\n".join(texts),
                "paragraphs": texts,
                "reason": reason,
                "termsheet_quote": termsheet_quote,
            }
        )
        ei = len(self.work.edits) - 1
        level, like = 0, ""
        if sibling is not None:
            new = []
            for k, t in enumerate(texts):
                self._next += 1
                q = Para(id=f"n{self._next}", text=t, inserted=(ei, k))
                if sibling.inserted is not None and not (sibling.heading_like or sibling.like):
                    _take_level(q, sibling)
                else:
                    q.heading, q.heading_like = sibling.heading or 1, _styled_as(sibling)
                self.work.paras.insert(at + k, q)
                new.append(self._view(q))
            return {"inserted": new}
        if heading:
            # the render numbers body paragraphs with hash levels too, so the sub-clause the
            # anchor sits in is the nearest paragraph above it at a shallower level
            depth = p.heading or 99
            following = self.work.paras[at] if at < len(self.work.paras) else None
            enclosing: Para | None
            if level_of:
                enclosing = self._styled_para(level_of)
                if enclosing is None or not enclosing.heading:
                    return {"error": f"level_of must be the id of a numbered paragraph: {level_of}"}
            elif following is not None and following.heading > depth:
                # the anchor is a title with paragraphs under it: a sibling written straight
                # after it takes them over (a heading after Part B became Part C,
                # and a condition fell outside the defined Part B). Which was meant
                # is the drafter's to say
                return {
                    "error": f"{id} is a heading with paragraphs under it; a heading straight"
                    " after it would take those paragraphs as its own. For a sub-heading under"
                    " it, give level_of= one of the headings under it; for a new heading beside"
                    " it, insert after its last paragraph with level_of= this heading."
                }
            else:
                enclosing = next(
                    (x for x in reversed(self.work.paras[:at]) if 0 < x.heading < depth), None
                )
            if enclosing is None and p.heading:
                # nothing heads the anchor in this piece: it is the piece's own title (a
                # schedule's empty Part B), and a sub-heading added after it goes beneath it
                level, like = p.heading + 1, ""
            else:
                level = enclosing.heading if enclosing else 1
                like = _styled_as(enclosing) if enclosing else ""
        # each later group of a grouped list added as a
        # heading with its items, the items body text where the groups before had numbered ones
        item = None
        if heading and enclosing is not None and enclosing in self.work.paras:
            after = self.work.paras.index(enclosing) + 1
            first = self.work.paras[after] if after < len(self.work.paras) else None
            if first is not None and first.heading > enclosing.heading and (first.inserted is None or first.like or first.heading_like):
                item = first
        new = []
        for k, t in enumerate(texts):
            self._next += 1
            q = Para(
                id=f"n{self._next}",
                text=t,
                inserted=(ei, k),
                heading=level if k == 0 else 0,
                heading_like=like if k == 0 else "",
                beneath=beneath and not heading,
                beneath_of=p.id if beneath and not heading and p.inserted is not None else "",
            )
            self.work.paras.insert(at + k, q)
            new.append(self._view(q))
        if p.heading and not heading and not beneath:
            # the rest of a grouped list inserted after a
            # numbered item printed as plain text, its group headings and items unnumbered
            return {
                "inserted": new,
                "unnumbered": f"these print as unnumbered body text, while {id}, which they follow,"
                " is numbered. Where they are items or headings of its list, delete them and"
                " insert again with level_of= the paragraph each is a sibling of (heading=True"
                " for a heading with its items); a term sheet's appendix is copied with"
                " copy_termsheet_text.",
            }
        if item is not None and len(new) > 1:
            return {
                "inserted": new,
                "unnumbered": f"the paragraphs under your new heading print as unnumbered body text,"
                f" while those under {enclosing.id if enclosing else ''}, the heading beside it, are numbered items."
                f" Where they are items of the same kind, restyle each with replace like={item.id}.",
            }
        return {"inserted": new}

    def insert_table(
        self, id: str, header: list[str], rows: list[list[str]], reason: str, termsheet_quote: str = ""
    ) -> dict:
        """A new table after paragraph `id` (after the whole table, if `id` is a row): header
        is the column titles, each row the same number of cells, values as the term sheet states
        them. For a matrix or a schedule of values the form has no table for."""
        p = self._para(id)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        header = [self._words(h) for h in header]
        rows = [[self._words(c) for c in r] for r in rows]
        if not header or not rows or any(len(r) != len(header) for r in rows):
            return {"error": "every row needs as many cells as the header"}
        if p.cells is not None and len(header) == len(p.cells):
            # under a row of a table of the same width these are rows of that table, and a
            # blank header is no row: a new one-row table with a blank header lost its row
            added = [r for r in [header, *rows] if any(c for c in r)]
            if added:
                return self.insert_after(id, [" | ".join(r) for r in added], reason, termsheet_quote=termsheet_quote)
        self.work.edits.append(
            {
                "kind": "table",
                "old_text": p.original or p.text,
                "new_text": "\n".join(" | ".join(r) for r in [header, *rows]),
                "header": header,
                "rows": rows,
                "reason": reason,
                "termsheet_quote": termsheet_quote,
            }
        )
        ei = len(self.work.edits) - 1
        at = self.work.paras.index(p) + 1
        new = []
        # a blank header is no row: written, it is a blank row the drafter then strikes as stray
        grid = [header, *rows] if any(c for c in header) else rows
        for k, cells in enumerate(grid):
            self._next += 1
            q = Para(id=f"t{self._next}", text=" | ".join(cells), cells=cells, inserted=(ei, k))
            self.work.paras.insert(at + k, q)
            new.append(self._view(q))
        return {"inserted_table": new}

    def copy_termsheet_text(
        self,
        id: str,
        sections: list[int],
        styles: list[str],
        reason: str = "",
        terms: dict[str, str] | None = None,
        strike_first: str = "",
        strike_last: str = "",
    ) -> dict:
        """Copy the term sheet's text into the contract after paragraph `id`, every line as the term sheet
        writes it, in the term sheet's order: for a procedure, conditions
        precedent, adjustments or a documents list the contract carries as the term sheet writes them.
        `sections` are the term sheet's section numbers, in order (read_termsheet_section shows each, one
        line per paragraph, its heading the first line). `styles` names, for each paragraph in
        turn, the form paragraph whose style and level it takes (a heading's, an item's, a
        sub-item's): its id, one this piece has struck as readily as one standing, or where this
        piece has none in the style needed, the exact words of
        a paragraph of another piece as search_contract shows them (one schedule's
        heading and items for another's). A line set in a numbered style loses the number the term sheet typed
        before it, which the style prints. `terms` maps the term sheet's words to the contract's, written
        into every copied line in this call: a party's name or shorthand to its defined term
        ("Customer"), the term sheet's label for a defined thing to the defined term ("the Expiry
        Date"), a reference to the term sheet's own appendix to the contract's schedule, a scan's slip
        to the right word. `strike_first` and `strike_last` are the ids of the first and last
        precedent paragraphs the copy replaces, struck in this call; the paragraph `id` is never
        among them. Use one call for a whole schedule or Part, then replace only what `terms`
        cannot express. Returns the new ids."""
        known = {s.index: s for s in self.record.sections}
        missing = [n for n in sections if n not in known]
        if not sections or missing:
            return {"error": f"the term sheet has no section {missing}" if missing else "name the sections"}
        struck_ids: list[str] = []
        if strike_first or strike_last:
            ids = [q.id for q in self.work.paras]
            if strike_first not in ids or strike_last not in ids:
                return {"error": "strike_first and strike_last must both be ids of this piece's paragraphs"}
            i, j = ids.index(strike_first), ids.index(strike_last)
            if j < i:
                return {"error": f"{strike_last} comes before {strike_first}"}
            struck_ids = ids[i : j + 1]
            if id in struck_ids:
                return {"error": f"{id} is the paragraph the copy follows: it cannot be struck by it"}
        lines, furniture = termsheet_paragraphs(self.record, sections)
        if len(styles) != len(lines):
            return {
                "error": f"these sections have {len(lines)} paragraphs and {len(styles)} styles"
                " were given: name one style for each paragraph, in order",
                "lines": [f"{k + 1}. {line[:100]}" for k, line in enumerate(lines)],
            }
        numbered: list[bool] = []
        for s in styles:
            m = self._model(s)
            if m is not None and m.inserted is None:
                numbered.append(bool(self.numbers.get(m.id)))
                continue
            found = self._styled_elsewhere(s) if m is None else None
            if found is None:
                return {
                    "error": f"each style must be the id of a form paragraph, or the exact words of"
                    f" a paragraph of another piece (never a clause or schedule title): {s}"
                }
            numbered.append(found[1])
        # a line reading as one of this piece's numbered form headings is that heading: its
        # words and its numbered style, whatever style was named (a copied
        # Part B printed no numbers, its subheadings copied in an unnumbered style)
        typed = re.compile(r"^\(?[0-9ivxlcdm]{1,4}(?:\.[0-9]+)*[.)]?\s+", re.I)
        # of several form paragraphs reading alike, the heading: the one least deep in the list
        headings: dict = {}
        for q in self.work.paras:
            number = self.numbers.get(q.id)
            if q.inserted is None and number and len((q.original or q.text).split()) <= 8:
                key = _bare(q.original or q.text)
                depth = len(re.findall(r"[.(]", number.strip(".")))
                if key not in headings or depth < headings[key][0]:
                    headings[key] = (depth, q)
        headings = {key: q for key, (_, q) in headings.items()}
        styles = list(styles)
        lines = list(lines)
        heading_ids: set[str] = set()
        for k, line in enumerate(lines):
            form = headings.get(_bare(typed.sub("", line)))
            if form is not None and len(line.split()) <= 10:
                lines[k], styles[k], numbered[k] = form.original or form.text, form.id, True
                heading_ids.add(form.id)
        # a style that prints its own number takes the line without the number the term sheet typed
        texts = [
            re.sub(r"^\(?[0-9ivxlcdm]{1,4}(?:\.[0-9]+)*[.)]?\s+", "", line, flags=re.I)
            if is_numbered
            else line
            for line, is_numbered in zip(lines, numbered, strict=True)
        ]
        for wrong, right in sorted((terms or {}).items(), key=lambda kv: -len(kv[0])):
            if wrong and wrong != right:
                # whole words only, and never inside the contract's own term where it already
                # stands ("Renewal" became "ReNewal", "the Service
                # Provider's" became "the Service Service Provider's")
                word = re.compile(rf"(?<![\w-]){re.escape(wrong)}(?![\w-])")
                texts = [
                    "".join(
                        part if k % 2 else word.sub(lambda _m: right, part)
                        for k, part in enumerate(re.split(f"({re.escape(right)})", t))
                    )
                    for t in texts
                ]
        # a run of two or more of the term sheet's table rows, as the reader gives them, is copied
        # as a table (rate tables came out one numbered paragraph per row);
        # the rest as paragraphs. A scan's spaced columns stay lines: they may be a list.
        runs: list[tuple[bool, list[int]]] = []
        for k, line in enumerate(lines):
            row = " | " in line and any(
                0 <= j < len(lines) and " | " in lines[j] for j in (k - 1, k + 1)
            )
            if runs and runs[-1][0] == row:
                runs[-1][1].append(k)
            else:
                runs.append((row, [k]))
        after = id
        inserted: list[dict] = []
        table: list[dict] = []
        for is_table, ks in runs:
            if is_table:
                cells = [[_clean(c) for c in re.split(r"\s*\|\s*", texts[k].strip())] for k in ks]
                width = len(cells[0])
                rows = [
                    r + [""] * (width - len(r)) if len(r) < width else r[: width - 1] + [" ".join(r[width - 1 :])]
                    for r in cells[1:]
                ]
                made = self.insert_table(after, cells[0], rows, reason)
                if "error" in made:
                    return made
                table += made["inserted_table"]
                after = made["inserted_table"][-1]["id"]
                continue
            made = self.insert_after(after, [texts[k] for k in ks], reason)
            if "error" in made:
                return made
            # styled while the precedent paragraphs they take their style from still stand
            for q, k in zip(made["inserted"], ks, strict=True):
                styled = self.replace(q["id"], q["text"], "", like=styles[k])
                if "error" in styled:
                    return styled
            inserted += made["inserted"]
            after = made["inserted"][-1]["id"]
            for q, k in zip(made["inserted"], ks, strict=True):
                if styles[k] in heading_ids:
                    self.__dict__.setdefault("_heading_of", {})[q["id"]] = styles[k]
        ids = [q["id"] for q in inserted] + [r["id"] for r in table]
        self.__dict__.setdefault("_copied", set()).update(ids)
        self.__dict__.setdefault("_copies", []).append(ids)
        for pid in struck_ids:
            self.delete(pid, "")
            self.work.edits[-1]["by_copy"] = True  # replaced by the term sheet's text, not struck bare
        return {
            "inserted": [self._view(self._para(q["id"])) for q in inserted],  # type: ignore[arg-type]
            **({"inserted_table": table} if table else {}),
            **({"struck": len(struck_ids)} if struck_ids else {}),
            **({"page_furniture_left_out": furniture} if furniture else {}),
        }

    def copy_termsheet_table(self, id: str, sections: list[int], reason: str) -> dict:
        """Copy a list the term sheet holds as a table (a list of items) into a new
        table after paragraph `id`, every row as the term sheet writes it. `sections` are the
        term sheet's section numbers holding the list, in order (a list the term sheet splits across
        many sections, one per group, names every one). The first pipe-separated row is
        the header; a header repeated in a later section is not copied again; a short row is
        padded. Use this, not insert_table, for any list of more than a few rows."""
        p = self._para(id)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        tables = self._termsheet_tables(sections)
        if isinstance(tables, dict):
            return tables
        out: dict = {"inserted_table": []}
        after = id
        for header, rows in tables:
            made = self.insert_table(after, header, rows, reason)
            if "error" in made:
                return made
            out["inserted_table"] += made["inserted_table"]
            after = made["inserted_table"][-1]["id"]
        out["rows_copied"] = sum(len(rows) for _, rows in tables)
        return out

    def _termsheet_tables(self, sections: list[int]) -> list[tuple[list[str], list[list[str]]]] | dict:
        """The term sheet's lists in these sections, each its header and its rows, as
        copy_termsheet_table copies them; an error where the sections hold none."""
        known = {s.index: s for s in self.record.sections}
        missing = [n for n in sections if n not in known]
        if not sections or missing:
            return {
                "error": f"the term sheet has no section {missing}" if missing else "name the sections"
            }
        rows_of = [
            (n, [_clean(c) for c in re.split(r"\s*\|\s*", line.strip())])
            for n in sections
            for line in plain(known[n].text).split("\n")
            if " | " in line
        ]
        lines = [cells for _, cells in rows_of]
        # a list the term sheet splits across sections, one per group, keeps each group's
        # heading as a group row before that group's rows, as the term sheet shows it
        group = {n: _clean(known[n].heading) for n in sections} if len(sections) > 1 else {}
        pending = ""
        current = None
        tables: list[tuple[list[str], list[list[str]]]] = []
        first_item: dict[int, str] = {}  # by table, the first cell of its first row of items
        for k, cells in enumerate(lines):
            if rows_of[k][0] != current:
                current = rows_of[k][0]
                pending = group.get(current, "")
            if tables:
                header = tables[-1][0]
                # another list with its own header gets its own table, its section heading its
                # title: a run of wider rows anywhere (a four-column list under a
                # two-column one), or a run of narrower rows where a section of the
                # term sheet opens (a three-column list after a four-column one). Narrower rows
                # inside a section (a blank last cell) are the same list.
                opens = k == 0 or rows_of[k - 1][0] != rows_of[k][0]
                run = k + 1 < len(lines) and len(lines[k + 1]) == len(cells)
                wider, narrower = len(cells) > len(header), len(cells) < len(header)
                if run and (wider or (narrower and opens)):
                    tables.append((cells, []))
                    pending = ""
                    continue
                restarts = (
                    k + 1 < len(lines)
                    and bool(lines[k + 1][0])
                    and lines[k + 1][0] == first_item.get(len(tables), "")
                )
                if cells == header and not (opens and restarts):
                    # a group's header repeated, its numbering running on; a list whose
                    # numbering starts again keeps the header the term sheet repeats for it
                    continue
                if len(cells) < len(header):
                    cells = cells + [""] * (len(header) - len(cells))
                elif len(cells) > len(header):
                    cells = cells[: len(header) - 1] + [" ".join(cells[len(header) - 1 :])]
                if pending:
                    tables[-1][1].append([pending] + [""] * (len(header) - 1))
                    pending = ""
                tables[-1][1].append(cells)
                first_item.setdefault(len(tables), cells[0])
            else:
                tables.append((cells, []))
        tables = [t for t in tables if t[1]]
        if not tables:
            return {"error": "those sections hold no pipe-separated table rows"}
        return tables

    def reconcile_list(self, first: str, last: str, rows: list[str], reason: str) -> dict:
        """Make the precedent's list, the table rows from `first` to `last`, the term sheet's, in
        one call. `rows` is the term sheet's whole list, every entry in the term sheet's order, each
        written in the precedent's columns with its cells ' | ' apart (a heading row of a grouped table
        as the precedent writes one), a scan's slips mended. The code lines the two lists up:
        a row in both stands untouched; a row that keeps most of its cells word for word and
        changes the rest (a revised number) takes the new words; a precedent row not in `rows`
        is struck; a row only in `rows` is added in its place. Returns how many rows of each."""
        ids = [q.id for q in self.work.paras]
        if first not in ids or last not in ids:
            return {"error": "first and last must be ids of this piece's table rows; call outline"}
        i, j = ids.index(first), ids.index(last)
        if j < i:
            return {"error": f"{last} comes before {first}"}
        olds = [q for q in self.work.paras[i : j + 1] if q.cells is not None]
        if not olds:
            return {"error": f"{first} to {last} holds no table rows"}
        # a heading row of a grouped table may have fewer cells than an item row: each row written takes
        # a width some row of the precedent's table has
        widths = sorted({len(q.cells or []) for q in olds})
        news = [[_clean(c) for c in re.split(r"\s*\|\s*", self._words(r).strip())] for r in rows if r.strip()]
        if not news:
            return {"error": "rows is empty: write the term sheet's whole list"}
        wrong = [r for r in news if len(r) not in widths]
        if wrong:
            return {
                "error": f"the precedent's rows have {' or '.join(map(str, widths))} cells; write"
                f" every row with as many as the precedent's row of its kind, an empty cell left"
                f" empty: {' | '.join(wrong[0])}"
            }

        def key(text: str) -> str:
            return " ".join(text.lower().split())

        counts = {"kept": 0, "changed": 0, "struck": 0, "added": 0}
        # one comment for the list, on its first change: a comment on every row of a long
        # list buries the lawyer's other notes, each row with the same comment
        said = [False]

        def why() -> str:
            if said[0]:
                return ""
            said[0] = True
            return reason
        struck: list[str] = []
        anchor = ""

        def add(cells: list[str]) -> None:
            # a row of the table after `anchor`, of its own width: the writer takes a row of
            # the table with as many cells as its model
            nonlocal anchor
            after = self._para(anchor or olds[0].id)
            assert after is not None
            text = " | ".join(cells)
            self.work.edits.append(
                {
                    "kind": "insert_row",
                    "old_text": after.original or after.text,
                    "new_text": text,
                    "cells": cells,
                    "reason": why(),
                    "termsheet_quote": "",
                }
            )
            self._next += 1
            q = Para(id=f"n{self._next}", text=text, cells=cells, inserted=(len(self.work.edits) - 1, 0))
            self.work.paras.insert(self.work.paras.index(after) + 1, q)
            anchor = q.id
            counts["added"] += 1

        matcher = difflib.SequenceMatcher(
            None,
            [key(q.text) for q in olds],
            [key(" | ".join(r)) for r in news],
            autojunk=False,
        )
        for op, a1, a2, b1, b2 in matcher.get_opcodes():
            if op == "equal":
                counts["kept"] += a2 - a1
                anchor = olds[a2 - 1].id
                continue
            pending = list(news[b1:b2])
            for old in olds[a1:a2]:
                # the same item re-worded keeps most of its cells word for word (its code and
                # description, its number revised); lists whose rows only look alike (a
                # shared number format) share fewer
                same = next(
                    (
                        k
                        for k, r in enumerate(pending)
                        if 2 * sum(key(x) == key(y) for x, y in zip(old.cells or [], r, strict=False) if key(x))
                        > len(r)
                    ),
                    -1,
                )
                if same >= 0:
                    for cells in pending[:same]:
                        add(cells)
                    self.replace(old.id, " | ".join(pending[same]), why())
                    counts["changed"] += 1
                    pending = pending[same + 1 :]
                else:
                    struck.append(old.id)
                anchor = old.id
            for cells in pending:
                add(cells)
        for pid in struck:
            self.delete(pid, why())
            counts["struck"] += 1
        return counts

    def move(self, id: str, after: str, reason: str, level_of: str = "") -> dict:
        """Move paragraph `id` (and, for a heading, nothing else: move its paragraphs too, one
        by one, in order) to sit after paragraph `after`. The words are unchanged; the
        paragraph sits as insert_after places it: at `after`'s level, or, with level_of, as a
        sibling of `level_of`, at its level and in its style (a new item of a list moved to
        follow the last limb of the item before it)."""
        p = self._para(id)
        target = self._para(after)
        if p is None or target is None:
            return {"error": "no such paragraph; call outline"}
        if p is target:
            return {"error": "a paragraph cannot follow itself"}
        if level_of and self._styled_para(level_of) is None:
            return {"error": f"no paragraph {level_of}; call outline"}
        text = p.text
        # moved without a new reason, the paragraph keeps the comment it carried
        own = [e for e in self.work.edits if e.get("para") == p.id]
        if p.inserted is not None:
            own = [self.work.edits[p.inserted[0]], *own]
        carried = next((e for e in reversed(own) if e.get("reason")), {})
        # a move the insert refuses must not have deleted the paragraph (a definition
        # was once lost so): the state before the delete is put back
        saved = (
            list(self.work.paras),
            list(self.work.deleted),
            list(self.work.edits),
            dict(self._struck_after),
            self._next,
        )
        self.delete(id, reason)
        made = self.insert_after(
            after,
            [text],
            reason or carried.get("reason", ""),
            termsheet_quote="" if reason else carried.get("termsheet_quote", ""),
            level_of=level_of,
        )
        if "error" in made:
            (
                self.work.paras,
                self.work.deleted,
                self.work.edits,
                self._struck_after,
                self._next,
            ) = saved
            return made
        moved = self._para(made["inserted"][0]["id"]) if "inserted" in made else None
        if moved is not None:
            moved.moved_from = p.moved_from or p.id
        if not level_of and p.inserted is not None and "inserted" in made:
            # a paragraph this piece added keeps the style it was given: a heading stays one
            q = self._para(made["inserted"][0]["id"])
            if q is not None:
                q.heading, q.heading_like, q.beneath = p.heading, p.heading_like, p.beneath
                q.beneath_of = p.beneath_of
                q.like, q.like_struck = p.like, p.like_struck
        return made

    def delete(self, id: str, reason: str, termsheet_quote: str = "") -> dict:
        """Remove paragraph `id` from the contract."""
        p = self._para(id)
        if p is None:
            return {"error": f"no paragraph {id}; already removed, or call outline"}
        at = self.work.paras.index(p)
        self.work.paras.remove(p)
        if p.inserted is None:
            self._struck_after[p.id] = self.work.paras[at - 1].id if at > 0 else ""
            self.work.deleted.append(p)
            self.work.edits.append(
                {
                    "kind": "delete",
                    "para": p.id,
                    "old_text": p.original,
                    "new_text": "",
                    "reason": reason,
                    "termsheet_quote": termsheet_quote,
                }
            )
        return {"removed": p.id, "text": p.original[:120], **self._whole_piece_struck()}

    def delete_range(self, first: str, last: str, reason: str, termsheet_quote: str = "") -> dict:
        """Remove every paragraph from `first` to `last` inclusive, in document order: a list
        or a schedule part struck whole."""
        ids = [p.id for p in self.work.paras]
        if first not in ids or last not in ids:
            return {"error": "no such paragraph; already removed, or call outline"}
        i, j = ids.index(first), ids.index(last)
        if j < i:
            return {"error": f"{last} comes before {first}"}
        removed = [self.delete(pid, reason, termsheet_quote)["removed"] for pid in ids[i : j + 1]]
        return {"removed": len(removed), "first": first, "last": last, **self._whole_piece_struck()}

    def _whole_piece_struck(self) -> dict:
        """Where no form paragraph of the piece stands but its title, a note to search for its
        uses: a precedent's appendix struck whole may be a form another piece still calls on."""
        standing = [p for p in self.work.paras if p.inserted is None and not self._own_title(p)]
        if standing or not self.work.deleted:
            return {}
        return {
            "whole_piece_struck": f"every form paragraph of {self.heading} is now struck. A form,"
            " schedule or provision the contract still calls on stands, its old deal's facts"
            f" replaced: search_contract its title and its defined terms, and where another piece"
            " names or relies on it, restore its paragraphs and change only what this deal changes."
        }

    def strike_empty_after(self, id: str, reason: str) -> dict:
        """Strike, as tracked deletions, the paragraphs without words that follow form
        paragraph `id` in the contract, up to the next paragraph with words: pictures (a form the
        precedent carries as images), drawings, page breaks and
        empty numbered paragraphs that print a number with nothing after it. None of them is in
        your view, so this is the only way to remove them. A section break stays. A pictured
        form is never retyped as text: where the playbook keeps it under an open point, leave
        its pictures. restore(id) puts back what this struck. Refused for an empty numbered paragraph the next ones are limbs of, or one a
        cross-reference names: striking it renumbers the clause. Returns the paragraph the
        strike stops at; your view cannot count the pictures."""
        p = self._para(id) or next((q for q in self.work.deleted if q.id == id), None)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        if p.inserted is not None:
            return {"error": f"{id} was added by this piece; name the form paragraph before them"}
        later = [
            q
            for q in [*self.work.paras, *self.work.deleted]
            if q.inserted is None and (_number(q.id) or 0) > (_number(id) or 0)
        ]
        nxt = min(later, key=lambda q: _number(q.id) or 0) if later else None
        if nxt is not None and p.heading and nxt.heading > p.heading + 1:
            # the next paragraph shown sits two levels below this one: a paragraph the view
            # does not show heads it, and striking that renumbers the clause
            return {
                "error": f"the paragraph after {id} sits {nxt.heading - p.heading} levels below it,"
                " so what lies between is a numbered paragraph its limbs hang from; it stays"
            }
        if not any(e.get("kind") == "strike_empty" and e.get("after") == id for e in self.work.edits):
            self.work.edits.append(
                {
                    "kind": "strike_empty",
                    "after": id,
                    "old_text": p.original,
                    "new_text": "",
                    "reason": reason,
                    "termsheet_quote": "",
                }
            )
        following = nxt.original if nxt is not None else None
        if following is None:
            units = [t for t, _ in self.contract_units]
            k = units.index(self.heading) + 1 if self.heading in units else len(units)
            following = next(
                (document_text(x) for _, u in self.contract_units[k:] for x in u.split("\n") if document_text(x)),
                "",
            )
        return {"struck_after": id, "up_to": following[:120]}

    def restore(self, id: str, reason: str = "") -> dict:
        """Put back a form paragraph this piece struck, where it stood, as the form has it,
        and the pictures and empty paragraphs it struck after paragraph `id`
        (strike_empty_after). For a fault that says a struck paragraph or a pictured form must
        stay: nothing is retyped."""
        struck = next((d for d in self.work.deleted if d.id == id), None)
        gone = [
            n
            for n, e in enumerate(self.work.edits)
            if (e.get("kind") == "strike_empty" and e.get("after") == id)
            or (struck is not None and e.get("kind") == "delete" and e.get("para") == id)
        ]
        if not gone:
            return {"error": f"{id} is not a paragraph this piece struck; see the faults' ids"}
        self.work.edits[:] = [e for n, e in enumerate(self.work.edits) if n not in gone]
        for q in self.work.paras:
            # an added paragraph names its edit by place in the list, which has just shortened
            if q.inserted is not None:
                ei, k = q.inserted
                q.inserted = (ei - sum(g < ei for g in gone), k)
        if struck is None:
            return {"restored": f"what follows {id}"}
        self.work.deleted.remove(struck)
        ids = [p.id for p in self.work.paras]
        after = self._struck_after.get(id, "")
        self.work.paras.insert(ids.index(after) + 1 if after in ids else 0, struck)
        return {"restored": struck.id, "text": struck.text[:120]}

    def open_point(self, summary: str, termsheet_quote: str = "", at: str = "") -> dict:
        """A point the term sheet does not settle, for the lawyer, in one sentence they can act on.
        `at` is the id of the paragraph it is about; the lawyer reads the point there."""
        point = {"summary": summary, "termsheet_quote": termsheet_quote}
        if at:
            p = self._para(at)
            if p is None:
                return {"error": f"no paragraph {at}; call outline"}
            point["at"], point["at_id"] = p.text, p.id  # its words, and which paragraph
        self.open_points.append(point)
        return {"recorded": len(self.open_points)}

    def revise_point(
        self, n: int, at: str = "", drop: bool = False, summary: str = "", termsheet_quote: str = ""
    ) -> dict:
        """Open point number `n` (as numbered in your task, from 1): with `at`, the id of the
        paragraph it is about, the lawyer reads it there; with `summary`, it is reworded, where
        it stands; with `termsheet_quote`, it quotes those words of the term sheet instead; with
        drop=True it is withdrawn, where the draft now settles it or another point already
        says it."""
        if not 1 <= n <= len(self.open_points):
            return {"error": f"there are {len(self.open_points)} open points; n counts from 1"}
        if drop:
            gone = self.open_points.pop(n - 1)
            return {"withdrawn": gone["summary"][:120]}
        if not (at or summary.strip() or termsheet_quote.strip()):
            return {"error": "give at, summary or termsheet_quote, or drop=True"}
        p = self._para(at) if at else None
        if at and p is None:
            return {"error": f"no paragraph {at}; call outline"}
        point = dict(self.open_points[n - 1])
        if p is not None:
            point["at"], point["at_id"] = p.text, p.id
        if summary.strip():
            point["summary"] = summary
        if termsheet_quote.strip():
            point["termsheet_quote"] = termsheet_quote
        self.open_points[n - 1] = point
        return {"revised": n, "summary": point["summary"][:120], **({"at": p.text[:120]} if p else {})}

    def note(self, id: str, reason: str, termsheet_quote: str = "") -> dict:
        """Rewrite the comment the lawyer reads beside paragraph `id`, a paragraph this piece
        changed or added, without changing its words: one sentence on what the change does and
        why, in the contract's terms. An empty reason removes the comment. termsheet_quote, the term sheet's
        exact words the new reason rests on, replaces the words the comment quoted; without it
        the comment keeps them. An open point is not a comment: move or withdraw it with
        revise_point."""
        p = self._para(id)
        if p is None:
            p = next((q for q in self.work.deleted if q.id == id), None)
        if p is None:
            return {"error": f"no paragraph {id}; call outline"}
        own = [e for e in self.work.edits if e.get("para") == id or e.get("cell") == id]
        if p.inserted is not None and p.inserted[1] == 0:
            own = [self.work.edits[p.inserted[0]], *own]
        elif p.inserted is not None and not own:
            # a later paragraph of an insert: the insert's reason is its first paragraph's
            # comment, so this paragraph's comment is an edit of its own
            own = [
                {
                    "kind": "note",
                    "para": id,
                    "old_text": p.text,
                    "new_text": p.text,
                    "reason": "",
                    "termsheet_quote": "",
                }
            ]
            self.work.edits.append(own[0])
        if not own:
            return {"error": f"{id} reads as the form has it and carries no comment"}
        if not reason.strip():
            for e in own:  # the comment is the last reason given: none is left to show
                e["reason"], e["termsheet_quote"] = "", ""
            return {"removed comment": id}
        # the comment the writer shows is the last reason on an edit that changed the words: a
        # replace writing them back as they stood is passed over, so the reason never goes there
        target = own[-1]
        before = p.original if p.inserted is None else _inserted_text(self.work.edits[p.inserted[0]], p)
        for e in own:
            if e.get("kind") == "modify" and e.get("new_text") is not None:
                same = _letters(e["new_text"]) == _letters(before)
                before = e["new_text"]
                if same and not e.get("like"):
                    continue
            target = e
        target["reason"] = reason
        if termsheet_quote.strip():
            for e in own:
                e["termsheet_quote"] = ""
            target["termsheet_quote"] = termsheet_quote
        return {"noted": id, "reason": reason, **({"termsheet_quote": termsheet_quote} if termsheet_quote else {})}

    # -- observing -----------------------------------------------------------------------
    def form_defects(self) -> list[dict]:
        """Defects of the base document's copy the piece's own structure shows: a table row holding
        all its words in one cell where the table's other rows fill the cells left empty (a
        label run together with its values), a cell repeating its row's label, a number typed
        by hand on a paragraph Word numbers."""
        found: list[dict] = []
        paras = self.work.paras
        k = 0
        while k < len(paras):
            if paras[k].cells is None:
                k += 1
                continue
            end = k
            while end < len(paras) and paras[end].cells is not None:
                end += 1
            # the render's bold marks, and its "1." for a cell Word numbers, are not words
            rows = [
                (r, [re.sub(r"^\d+\.$", "", c.replace("**", "").strip()) for c in r.text.split(" | ")])
                for r in paras[k:end]
            ]
            values: dict[int, set[str]] = {}  # by width, what rows of several cells hold after their label
            for _, cells in rows:
                filled = [c for c in cells if c]
                if len(filled) > 1:
                    values.setdefault(len(cells), set()).update(filled[1:])
            for r, cells in rows:
                filled = [c for c in cells if c]
                if len(filled) == 1 and any(
                    filled[0].endswith(" " + v) for v in values.get(len(cells), set())
                ):
                    defect = "a label and its values run together in one cell; the table's other rows give each value its own cell"
                elif len(filled) > 1 and any(c.startswith(filled[0] + " ") for c in filled[1:]):
                    defect = "a cell repeats the row's label"
                else:
                    continue
                found.append({"id": r.id, "text": r.text[:160], "defect": defect})
            k = end
        for p in paras:
            # "SCHEDULE 5" typed on a heading whose style prints "Schedule 4" (or "Part A")
            words = [c.split()[0].lower() for c in self.numbers.get(p.id, "").split(",") if len(c.split()) > 1]
            if p.cells is None and words and re.match(rf"(?:{'|'.join(map(re.escape, words))})\s+\S+", p.text, re.I):
                found.append(
                    {
                        "id": p.id,
                        "text": p.text[:160],
                        "defect": f"a number typed by hand where Word numbers the heading {self.numbers[p.id]}",
                    }
                )
                continue
            if p.cells is None and p.id in self.numbers and re.match(r"\(?[0-9A-Za-z]{1,5}[.)]\s", p.text):
                found.append(
                    {
                        "id": p.id,
                        "text": p.text[:160],
                        "defect": f"a number typed by hand where Word numbers the paragraph {self.numbers[p.id]}",
                    }
                )
        return found

    def check(self) -> dict:
        """Facts about the clause as it now stands: paragraphs still holding a bracket, a changed
        paragraph whose round brackets no longer close, footnote notes
        still attached, a term defined twice, the same paragraph twice, a form
        paragraph deleted and its words typed back in as a new one, a defined term written with
        another hyphen or spacing, a list the piece changed whose
        punctuation no longer closes it, a form paragraph struck alone
        with no reason and none of the term sheet's words, a precedent paragraph left standing among the
        term sheet's copied text, defects of the base document's copy still standing
        (form_defects), and on a precedent each number the term sheet never states (old_deal_numbers):
        the old deal's serial numbers, years and amounts, to replace or strike unless a rule keeps
        them. Reports only."""
        brackets = [
            {"id": p.id, "text": p.text[:160]}
            for p in self.work.paras
            if "[" in p.text or "]" in p.text
        ]
        notes = [
            {"id": p.id, "notes": self._notes_of(p.text)}
            for p in self.work.paras
            if self._notes_of(p.text)
        ]
        seen: dict[str, list[str]] = {}
        for p in self.work.paras:
            head = p.text.lstrip('#* ["“')
            if '" means' in head[:120] or "” means" in head[:120]:
                term = head.split('" means')[0].split("” means")[0].strip()
                seen.setdefault(term, []).append(p.id)
        doubled = [{"term": t, "ids": ids} for t, ids in seen.items() if len(ids) > 1]
        same: dict[str, list[str]] = {}
        for p in self.work.paras:
            same.setdefault(" ".join(p.text.lower().split()), []).append(p.id)
        repeated = [
            {"text": t[:120], "ids": ids} for t, ids in same.items() if len(ids) > 1 and len(t) > 20
        ]
        now = {" ".join(p.text.lower().split()): p for p in self.work.paras if p.inserted}
        retyped = [
            {
                "deleted": d.id,
                "inserted_as": now[" ".join(d.text.lower().split())].id,
                "text": d.text[:120],
            }
            for d in self.work.deleted
            if " ".join(d.text.lower().split()) in now
            and now[" ".join(d.text.lower().split())].moved_from != d.id
            # the term sheet's own words copied in, not a precedent paragraph retyped
            and now[" ".join(d.text.lower().split())].id not in self.__dict__.get("_copied", set())
        ]
        unbalanced = [
            {"id": p.id, "text": p.text[:160]}
            for p in self.work.paras
            if p.text.count("(") != p.text.count(")")
            and (p.inserted or (p.original and p.original.count("(") == p.original.count(")")))
        ]
        among = []  # a form paragraph left standing inside a stretch of the term sheet's copied text
        for copy_ids in self.__dict__.get("_copies", []):
            at = [k for k, q in enumerate(self.work.paras) if q.id in copy_ids]
            if at:
                among += [
                    {"id": q.id, "text": q.text[:120]}
                    for q in self.work.paras[at[0] : at[-1] + 1]
                    if q.inserted is None
                ]
        # the lists of the piece: siblings at one level, deeper limbs between them
        runs: list[list[Para]] = []
        open_: dict[int, list[Para]] = {}
        for q in self.work.paras:
            level = q.heading if q.cells is None else 0
            for deeper in [k for k in open_ if not level or k > level]:
                runs.append(open_.pop(deeper))
            if level:
                open_.setdefault(level, []).append(q)
        runs += open_.values()
        joined = re.compile(r";\s*(?:and|or)?\s*$")
        punctuation = []
        for run in runs:
            touched = any(q.inserted is not None or q.text != q.original for q in run)
            semis = sum(bool(joined.search(q.text)) for q in run[:-1])
            stops = sum(q.text.rstrip().endswith(".") for q in run[:-1])
            if len(run) < 2 or not touched or semis <= stops:
                continue  # sentences, not a list's limbs; or a list the piece left as it was
            for q in run[:-1]:
                if q.text.rstrip().endswith("."):
                    punctuation.append({"id": q.id, "text": q.text[-80:], "fault": "ends with a full stop before the list's last limb"})
            if not re.search(r"[.;:,]\s*$|\b(?:and|or)\s*$", run[-1].text):  # "," runs on to the closing words
                punctuation.append({"id": run[-1].id, "text": run[-1].text[-80:], "fault": "the list's last limb has no closing punctuation"})
        bare = {
            e["para"]
            for e in self.work.edits
            if e["kind"] == "delete"
            and not e.get("by_copy")
            and not (e.get("reason") or "").strip()
            and not (e.get("termsheet_quote") or "").strip()
        }
        # a form paragraph struck alone, the form paragraphs either side of it standing: a list
        # or a schedule cleared to be replaced is not this
        order = sorted(
            [q for q in self.work.paras if q.inserted is None] + list(self.work.deleted),
            key=lambda q: int(q.id[1:]) if q.id[1:].isdigit() else 0,
        )
        gone = {d.id for d in self.work.deleted}
        alone = {
            q.id
            for k, q in enumerate(order)
            if 0 < k < len(order) - 1 and order[k - 1].id not in gone and order[k + 1].id not in gone
        }
        # a defined term written with another hyphen or spacing ("Sub Contractor" for "Sub-Contractor")
        defined = {_bare(t["term"]): t["term"] for t in self.defined_terms() if len(_bare(t["term"])) > 3}
        variants = []
        for q in self.work.paras:
            if q.inserted is None and q.text == q.original:
                continue
            for m in re.finditer(r"\b[A-Z][\w’']*(?:[ -][A-Z][\w’']*)+", q.text):
                written = re.sub(r"(?:['’]s?)$", "", m.group(0))  # "Customer Parties'" is the term's possessive
                term = defined.get(_bare(written))
                same = term is not None and term.replace("’", "'") == written.replace("’", "'")
                if term and not same and not written.isupper():  # a heading's capitals are its style
                    variants.append({"id": q.id, "written": written, "defined": term})
        return {
            "defined_term_variants": variants,
            "list_punctuation": punctuation,
            "struck_without_reason": [
                {"id": d.id, "text": d.text[:120]}
                for d in self.work.deleted
                if d.id in bare and d.id in alone and re.search(r"\w", d.text)
            ],
            "precedent_among_copied_text": among,
            "repeated_paragraphs": repeated,
            "unbalanced_parentheses": unbalanced,
            "brackets": brackets,
            "footnote_notes": notes,
            "defined_twice": doubled,
            "deleted_then_retyped": retyped,
            "precedent_defects": self.form_defects(),
            "old_deal_numbers": self._old_deal_numbers(),
        }

    def _old_deal_numbers(self) -> list[dict]:
        """On a precedent, each paragraph's numbers of four or more digits (a serial number, a
        year, an amount) that the term sheet states nowhere: the old deal's, unless a rule keeps
        them. A precedent's own asset details can survive into the new draft."""
        if not self.precedent:
            return []

        def numbers(text: str) -> list[str]:
            return [n.replace(",", "") for n in re.findall(r"\d[\d,]*\d", text) if len(n.replace(",", "")) >= 4]

        termsheet = set()
        for s in self.record.sections:
            termsheet.update(numbers(s.heading + " " + s.text))
        for f in self.record.facts:
            termsheet.update(numbers(" ".join(str(v) for v in vars(f).values())))
        out = []
        for p in self.work.paras:
            kept = [n for n in dict.fromkeys(numbers(p.text)) if n not in termsheet]
            if kept:
                out.append({"id": p.id, "numbers": kept, "text": p.text[:160]})
        return out

    def tools(self) -> list[StructuredTool]:
        return [
            StructuredTool.from_function(self._traced(m))
            for m in (
                self.outline,
                self.read,
                self.search_contract,
                self.defined_terms,
                self.template_clause,
                self.template_definition,
                self.search_record,
                self.read_termsheet_section,
                self.replace,
                self.insert_after,
                self.insert_table,
                self.copy_termsheet_table,
                self.reconcile_list,
                self.copy_termsheet_text,
                self.move,
                self.delete,
                self.restore,
                self.delete_range,
                self.strike_empty_after,
                self.open_point,
                self.revise_point,
                self.note,
                self.check,
            )
        ]

    def _traced(self, method):
        import functools

        @functools.wraps(method)
        def wrapped(*args, **kwargs):
            result = method(*args, **kwargs)
            self.trace.append(
                {
                    "tool": method.__name__,
                    "args": {k: (str(v)[:200]) for k, v in kwargs.items()},
                    "result": json.dumps(result, default=str)[:300],
                }
            )
            return result

        return wrapped


class RuleAccount(BaseModel):
    rule: str = Field(description="the rule's opening words, enough to know which one")
    where: str = Field(
        description="the paragraph ids where it is met, or why it does not reach this piece"
    )
    status: Literal["done", "not_applicable", "outstanding"] = Field(
        default="done",
        description="done: the piece carries it; not_applicable: it does not reach this piece,"
        " or its work belongs in another piece (a definition outside this piece's alphabetical"
        " stretch, a clause another piece holds); outstanding: this piece should carry it and"
        " does not yet",
    )


class Outcome(BaseModel):
    note: str = Field(description="two sentences: what changed in this clause, and what is open")
    account: str = Field(
        description="your account of the work, item by item, for a colleague checking it"
    )
    rules_accounted: list[RuleAccount] = Field(
        default_factory=list,
        description="one entry for every rule you were given, after re-reading the piece",
    )


class RuleVerdict(BaseModel):
    rule: str = Field(description="the rule's opening words, enough to know which one")
    met: Literal["yes", "no", "does_not_reach"] = Field(
        description="yes: the piece as it reads carries it; no: this piece should carry it and"
        " does not, or carries it wrongly; does_not_reach: it is not about this piece, or its"
        " work belongs in another piece"
    )
    fault: str = Field(
        default="",
        description="where no: each paragraph id that fails it and the exact words that paragraph"
        " must say instead, as a lawyer marks a draft; never 'check', 'confirm', 'if' or"
        " 'unless': read it now and say",
    )
    breaks: str = Field(
        default="",
        description="where no: the rule's own words the piece fails, copied exactly from the task",
    )
    reads: str = Field(
        default="",
        description="where no and a paragraph says it wrongly: that paragraph's words, copied"
        " exactly; empty where the fault is that the piece lacks something",
    )


class RuleCheck(BaseModel):
    verdicts: list[RuleVerdict] = Field(description="one for every rule in the task")


def check_rules(agent: ClauseAgent, llm, max_turns: int = 40) -> list[str]:
    """The finished piece read afresh against the task it was drafted from, rule by rule. The
    drafter's own account is not shown: only an account of "outstanding" was ever acted on, and
    a piece accounting every rule done still left rules unmet. Its tools only read. Returns
    each rule not met, as a fault for the drafter."""
    struck = [(d.id, d.text) for d in agent.work.deleted if d.inserted is None]
    reading = (agent.outline, agent.read, agent.search_contract, agent.defined_terms,
               agent.template_clause, agent.template_definition, agent.search_record,
               agent.read_termsheet_section, agent.check)
    outcome = run_agent(
        system_prompt=agent.system_brief or brief(),
        task_message=(
            f"A DRAFTER WAS GIVEN THIS TASK:\n\n{agent.task}\n\n"
            "THE PIECE AS IT NOW READS:\n"
            + "\n".join(f"{p.id}: {p.text}" for p in agent.work.paras)
            + (
                "\n\nFORM PARAGRAPHS IT STRUCK:\n" + "\n".join(f"{i}: {t}" for i, t in struck)
                if struck
                else ""
            )
            + "\n\nGo down every rule in the task and judge it against the piece as it now"
            " reads, with the term sheet's sections the rule depends on. A rule that reaches several"
            " paragraphs is met only where every one of them meets it: go through each. For each"
            " rule not met, name every paragraph id that fails it and what it must say instead."
            " A rule this piece meets, or one only a fact the term sheet lacks could meet, is yes."
            " Before you say a piece lacks something (a definition, a provision), search_contract"
            " for it: what another piece already holds is yes. A doubt or a preference is yes."
            " Then submit_outcome."
        ),
        tools=[StructuredTool.from_function(m) for m in reading],
        final_output_schema=RuleCheck,
        max_iterations=max_turns,
        llm=llm,
    )
    agent.rule_check = [v.model_dump() for v in outcome.verdicts]
    # a fresh read can call a right piece wrong (a term another piece defines, called
    # undefined): a fault stands only where the rule's words it quotes are the task's and
    # the words it quotes are the piece's
    task = _bare(agent.task)
    piece = _bare(" ".join(p.text for p in agent.work.paras))
    proven = [
        v
        for v in outcome.verdicts
        if v.met == "no" and _bare(v.breaks) and _bare(v.breaks) in task and _bare(v.reads) in piece
    ]
    return [f"{v.rule}: {v.fault}" for v in proven]


def drafter_task(
    heading: str,
    rules: list[str],
    termsheet_sections: list[int],
    record: Record,
    house_facts: dict[str, str] | None,
    precedent: bool = False,
    defects: list[dict] | None = None,
    term_homes: list[list] | None = None,
) -> str:
    """What one drafter is told: its piece, the rules the planner copied for it, the term sheet's
    sections it needs opened in full, the deal's shape, the client's own details, and the defects
    of the base document's copy found in the piece. Nothing about other pieces."""
    opened = []
    for n in termsheet_sections:
        section = next((s for s in record.sections if s.index == n), None)
        if section is not None:
            homes = [(term, piece) for m, term, piece in term_homes or [] if m == n]
            mine = [term for term, piece in homes if piece in (heading, EVERY_PIECE)]
            theirs = [f"{term} ({piece})" for term, piece in homes if piece not in (heading, EVERY_PIECE)]
            opened.append(
                f"[{n}] {section.heading.strip()}\n{plain(section.text).strip()}"
                + (
                    f"\nTHIS PIECE WRITES, where the playbook writes it at all: {'; '.join(mine)}."
                    " Where the base document already provides for a term, its words stand; a term"
                    " the playbook spends or keeps out is not written."
                    if mine
                    else ""
                )
                + (
                    f"\nWRITTEN BY ANOTHER PIECE, never here (change only what your own words must"
                    f" for it): {'; '.join(theirs)}."
                    if theirs
                    else ""
                )
            )
    named = document_being_edited(precedent)
    return (
        f"YOUR PIECE: {heading}.\n"
        + (f"{named}\n\n" if named else "\n")
        + "RULES FOR THIS PIECE:\n"
        + ("\n".join(f"- {r}" for r in rules) or "- (none)")
        + "\n\nTHE TERM SHEET'S SECTIONS FOR THIS PIECE"
        + (
            " (search_record finds a term elsewhere in the term sheet):\n\n" + "\n\n".join(opened)
            if opened
            else ": none. The term sheet says nothing about this piece."
        )
        + f"\n\nTHE DEAL, as the term sheet reads:\n{deal_lines(record)}\n"
        + (
            "\nTHE CLIENT'S OWN DETAILS, for where the form asks for its particulars:\n"
            + "\n".join(f"- {k}: {v}" for k, v in house_facts.items())
            + "\n"
            if house_facts
            else ""
        )
        + (
            "\nDEFECTS OF THE BASE DOCUMENT'S COPY FOUND IN THIS PIECE, each mended with the"
            " smallest change:\n"
            + "\n".join(f"- {d['id']}: {d['defect']}: {d['text']}" for d in defects)
            + "\n"
            if defects
            else ""
        )
        + "\nStart with outline. Draft it. Run check and put right what it reports. Then read"
        " the whole piece again and go down every rule above: where it is met, note the"
        " paragraph ids; where it is not, fix it now; where it does not reach this piece, say"
        " why. Only then submit_outcome, with that account in rules_accounted."
    )


def draft_clause(
    heading: str,
    clause_text: str,
    anchored_text: str,
    contract_text: str,
    contract_units: list[tuple[str, str]],
    record: Record,
    footnotes: dict[str, str],
    rules: list[str],
    termsheet_sections: list[int],
    house_facts: dict[str, str] | None,
    llm,
    max_turns: int = 150,
    precedent: bool = False,
    numbers: dict[tuple[str, int], str] | None = None,
    template_text: str = "",
    term_homes: list[list] | None = None,
    form_bytes: bytes = b"",
    before_text: str = "",
) -> tuple[ClauseAgent, str]:
    """One clause drafted by the agent in one loop, on the instructions the planner wrote for
    it: its rules, the term sheet's sections it needs, the deal's shape. Returns the agent's state
    and its note."""
    agent = ClauseAgent(
        heading, clause_text, anchored_text, contract_text, record, footnotes, contract_units, numbers
    )
    agent.template_text = template_text
    agent.form_bytes, agent.before_text = form_bytes, before_text
    agent.system_brief = brief(precedent)
    agent.precedent = precedent
    task = drafter_task(
        heading,
        rules,
        termsheet_sections,
        record,
        house_facts,
        precedent,
        defects=agent.form_defects(),
        term_homes=term_homes,
    )
    agent.task = task
    note = ""
    try:
        outcome = run_agent(
            system_prompt=agent.system_brief,
            task_message=task,
            tools=agent.tools(),
            final_output_schema=Outcome,
            max_iterations=max_turns,
            llm=llm,
        )
        note = outcome.note
        agent.rules_accounted = [r.model_dump() for r in outcome.rules_accounted]
        left = [r for r in outcome.rules_accounted if r.status == "outstanding"]
        if left:
            # a piece that submits saying its work is outstanding (a schedule's
            # Part B left empty) is sent back once to do it
            note = _finish(
                agent,
                "You submitted with rules outstanding; what you wrote stands. Do each now that"
                " this piece holds the place for: "
                + "; ".join(f"{r.rule} ({r.where})" for r in left)
                + ". Where a rule's work belongs in another piece (a definition outside this"
                " piece's stretch of the alphabet, a clause another piece holds), change nothing"
                " for it and account it not_applicable. Then submit_outcome.",
                llm,
                note,
            )
    except Exception as error:  # the budget ran out: what was recorded stands
        # a piece left half done went out as finished: it is
        # sent back once, on what it has written, to finish
        note = _finish(
            agent,
            "You ran out of turns before you finished this piece; what you wrote stands."
            " Finish it: take each rule and each term sheet section you were given that the"
            " piece does not yet carry, then submit_outcome.",
            llm,
            f"budget spent: {type(error).__name__}",
        )
    # one fix pass put right some of what the check found and left the rest: the piece
    # is checked again after each fix, until the check passes or the rounds run out
    for _ in range(_CHECK_ROUNDS if rules else 0):
        try:
            unmet = check_rules(agent, llm)
        except AgentDidNotCompleteError:  # the piece stands as drafted, and says it went unchecked
            note = f"{note} (its rules went unchecked: the check ran out of turns)"
            break
        if not unmet:
            break
        note = _finish(
            agent,
            "A fresh read of your piece against your rules found the rules below not met;"
            " what you wrote stands. Put each right, in every paragraph it names, before you"
            " submit_outcome:",
            llm,
            note,
            items=unmet,
        )
    return agent, note


_CHECK_ROUNDS = 1


def _finish(agent: ClauseAgent, instruction: str, llm, note: str, items: list[str] | None = None) -> str:
    """The piece sent back once, on what it has written, to finish, each of `items` a fault of
    its own; its note and account are the finished ones, or, where the second pass also runs
    out, the first ones stand and the note says the pass did not finish."""
    finished = fix_clause(agent, [instruction, *(items or [])], llm)
    if finished.startswith("fix budget spent"):
        return f"{note} (sent back to put right what it missed, it ran out of turns: some may still be missed)"
    agent.rules_accounted = agent.fix_accounted
    return finished


def fix_task(
    heading: str,
    faults: list[str],
    drafted_from: str = "",
    open_points: list[dict] | None = None,
    struck: list[tuple[str, str]] | None = None,
) -> str:
    """What the drafter is told when its piece comes back from review: the task it drafted
    from (its rules, the term sheet's sections, the client's details, the document it edits), the
    form paragraphs it struck (read finds only standing ones, so restore needs these ids),
    the faults, and that the fix is the smallest change that puts each right."""
    return (
        (
            f"YOU DRAFTED THIS PIECE FROM THIS TASK, and all of it still holds:\n\n{drafted_from}\n\n"
            if drafted_from
            else ""
        )
        + f"YOUR PIECE: {heading}, as you drafted it, stands in your tools.\n\n"
        + (
            "ITS OPEN POINTS, each a comment the lawyer reads (revise_point moves, rewords or"
            " withdraws one):\n"
            + "\n".join(
                f"  {n}. {o.get('summary', '')}"
                + (f' (on "{o["at"][:80]}")' if o.get("at") else "")
                for n, o in enumerate(open_points, 1)
            )
            + "\n\n"
            if open_points
            else ""
        )
        + (
            "FORM PARAGRAPHS IT STRUCK (restore(id) puts one back):\n"
            + "\n".join(f"  {i}: {text[:90]}" for i, text in struck)
            + "\n\n"
            if struck
            else ""
        )
        + "A READ OF IT against the whole draft found these faults:\n"
        + "\n".join(f"- {f}" for f in faults)
        + "\n\nA reviewer can be wrong. Where a fault asks you to strike, narrow, move or reverse"
        " words you wrote from the term sheet or from a rule for this piece, read that term sheet"
        " provision or rule again first: where it calls for the words as written, they stay and"
        " your note quotes it. A fact of the client's own is the one in your task, never one you"
        " make up or leave blank."
        "\n\nEach reason you give replaces the comment the lawyer reads beside that change: say"
        " why the words now read as they do, from the term sheet or the form, never that a fault, a"
        " read or a request called for it. A comment that is wrong about words you are not"
        " changing is rewritten with note; an empty note removes it."
        "\n\nPut each right with the smallest change: replace the words, insert or delete the"
        " paragraph the fault names, and nothing else. Nothing the faults do not name is"
        " re-drafted, re-worded or moved. Where a fault says a name, term or figure should match"
        " another paragraph, read that paragraph and copy it exactly as it is written there."
        " Where a fault asks you to change the form's own words"
        " on a point the term sheet does not reach, or a figure the term sheet gives, they stay and"
        " your note says why. Where a fault says a value or alternative is not the term sheet's,"
        " search_record for the term sheet's own words first: if they say what your piece says,"
        " change nothing and your note says so. Where a fault says something is missing, find"
        " the words that call for it (search_record, search_contract): if they are not in that"
        " provision, or the thing it ties to (an option, a schedule) is not in the draft,"
        " change nothing and your note says so. Before you decline any fault because a term is used or unused, or a paragraph exists, settle it with search_contract: what you remember of the draft is not evidence. A fault that says a form paragraph the piece struck must stay is put right with restore (never by typing its words back in). Read each paragraph you changed, then"
        " read again every paragraph each fault names: where it does not yet read as the fault's"
        " right text, the fault is not put right, so put it right or say in its entry why not;"
        " an entry never says a change was made that read does not show. Then submit_outcome:"
        " your note, and one rules_accounted entry per fault saying what you did."
    )


def fix_clause(agent: ClauseAgent, faults: list[str], llm, max_turns: int = 60) -> str:
    """The piece's own drafter puts right the faults a reviewer found, working on the piece as
    it drafted it: the same agent, the same tools, the smallest change. Returns its note."""
    note = ""
    try:
        outcome = run_agent(
            system_prompt=agent.system_brief or brief(),
            task_message=fix_task(
                agent.heading,
                faults,
                agent.task,
                agent.open_points,
                [(d.id, d.text) for d in agent.work.deleted if d.inserted is None],
            ),
            tools=agent.tools(),
            final_output_schema=Outcome,
            max_iterations=max_turns,
            llm=llm,
        )
        note = outcome.note
        agent.fix_accounted = [r.model_dump() for r in outcome.rules_accounted]
    except Exception as error:  # the budget ran out: what was recorded stands
        note = f"fix budget spent: {type(error).__name__}"
    return note


def _form_words_changed(form: str, now: str) -> list[dict[str, str]]:
    """The form's words a replace changed, run by run: what the form said and what it says
    now, so the drafter sees every word it touched, not only the ones it meant to."""
    was, new = form.split(), now.split()
    out = []
    matcher = difflib.SequenceMatcher(None, was, new, autojunk=False)
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op != "equal":
            out.append({"form": " ".join(was[i1:i2]), "now": " ".join(new[j1:j2])})
    return out
