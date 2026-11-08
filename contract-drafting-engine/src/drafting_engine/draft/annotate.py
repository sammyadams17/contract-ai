"""The annotated term sheet: the term sheet itself, with a green note after each item naming
where the contract carries it.

It is the term sheet's Word file untouched, with `[Clause 5.4(a)(i)]`, `["Commencement
Date"]` or `[Schedule 1 - Fees]` after the items. A note points at the final contract, never at what a drafter did on the way.

Code does three structural things here: it names every place of the final contract the way a
lawyer cites it (from Word's own numbering, with every tracked change accepted); it lists the
term sheet's items in the term sheet's order; and it writes each note after its item. Which place
carries which item is read by an agent, a batch of items at a time, choosing only from the
contract's real places, so no note can cite a clause the contract does not have.

A term sheet that exists only as a signed PDF has no Word file to mark up; for it the term sheet's
text is laid out afresh, section by section, with the same notes.
"""

from __future__ import annotations

import io
import pathlib
import re
from collections.abc import Callable, Iterator, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import ClassVar

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph as DocxParagraph
from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field, model_validator

from drafting_engine.draft.annotate_footnotes import FootnoteNotes
from drafting_engine.draft.annotate_widths import note_fits
from drafting_engine.llm.agent import AgentDidNotCompleteError, run_agent
from drafting_engine.termsheet.document import unwrap_smart_tags
from drafting_engine.word.numbering import (
    ClauseNumber,
    ClauseNumberer,
    _paragraph_numbering,
    load_numbering_definitions,
)

BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "annotator.md"
GREEN = RGBColor(0x00, 0xB0, 0x50)
#: How much term sheet text one agent reads at a time: enough for a few sections in context,
#: small enough that a batch stays a short task.
BATCH_CHARS = 5000

_CONTAINER = re.compile(r"^(Schedule|Appendix)\s*(\d+|[A-Z])\b")
_PART = re.compile(r"^Part\s*(\d+|[A-Z](?![a-z]))")
# a schedule heading typed as text, not numbered by Word: "SCHEDULE 3Service Levels"
_TYPED_CONTAINER = re.compile(r"^(SCHEDULE|APPENDIX)\s*(\d+|[A-Z])(?![\d.)])")
_DEFINITION = re.compile(r'^\s*["“]([^"”]+)["”]\s+(?:means|has the meaning)\b')
_WS = re.compile(r"\s+")
_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def brief() -> str:
    return BRIEF_FILE.read_text()


def _norm(text: str) -> str:
    return _WS.sub(" ", text).strip().casefold()


# ---------------------------------------------------------------- the contract's places


@dataclass(frozen=True)
class Location:
    label: str
    text: str
    comments: tuple[str, ...] = ()  # the draft's comments on this paragraph or row


def own_number(number: ClauseNumber, container: str = "", part: str = "") -> str:
    """A schedule paragraph's number within its schedule and part, as the page shows it:
    "1.3(d)", never Word's composed "Schedule 31.3(d)" or "Part A1". Empty where the number
    hangs off something unnumbered (a definition's limb composes as ".(a)"): such a
    paragraph has no name of its own."""
    if number.full.strip().startswith("."):
        return ""
    own = number.local.strip()
    for prefix in (container, part):
        if prefix and own.startswith(prefix):
            own = own[len(prefix) :].strip()
    if not part and (borrowed := _PART.match(own)):
        # A form in an appendix numbered on from the schedule list before it: Word composes
        # the schedule's Part into the number, the page shows only the paragraph's own.
        own = own[borrowed.end() :].strip()
    return own.replace(".(", "(").rstrip(".")


def _visible(element) -> str:  # type: ignore[no-untyped-def]
    """The text as it reads once every change is accepted: inserted text counts, deleted
    and moved-away text (w:delText, w:moveFrom) does not."""
    return "".join(
        node.text or ""
        for node in element.iter(qn("w:t"))
        if not any(a.tag == qn("w:moveFrom") for a in node.iterancestors())
    )


def _mark_deleted(element) -> bool:  # type: ignore[no-untyped-def]
    """A paragraph whose mark is a tracked deletion: Word removes it on accepting, so it
    takes no number."""
    return element.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}") is not None


def _row_deleted(row) -> bool:  # type: ignore[no-untyped-def]
    return row._tr.find(f"{qn('w:trPr')}/{qn('w:del')}") is not None


def contract_locations(docx_bytes: bytes) -> list[Location]:
    """Every paragraph and table row of the contract as it reads with every change accepted,
    each named as a lawyer cites it: Clause 4.3(d), Paragraph 2 of Part A of Schedule 5,
    "Commencement Date", a row of a key terms table (Fees), Cover page. A paragraph with no
    name of its own (a definition's limb, an unnumbered line) takes the name of the place
    it sits in. The tracked file is read as Word numbers it once the changes are accepted:
    a paragraph whose mark is deleted and a deleted table row are not there."""
    document = Document(io.BytesIO(docx_bytes))
    numberer = ClauseNumberer(load_numbering_definitions(docx_bytes))
    said = _comments(docx_bytes)

    def comments_on(element) -> tuple[str, ...]:  # type: ignore[no-untyped-def]
        return tuple(
            said[i]
            for i in dict.fromkeys(r.get(qn("w:id")) for r in element.iter(qn("w:commentReference")))
            if i in said
        )

    out: list[Location] = []
    current = "Cover page"
    container: str | None = None
    part: str | None = None
    in_body = False
    for element in document.element.body.iterchildren():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "tbl":
            rows = Table(element, document).rows
            first = [_visible(c._tc).strip() for c in rows[0].cells] if rows else []
            terms_table = [_norm(c) for c in first[:2]] == ["term", "meaning"]
            for row in rows[1:] if terms_table else rows:
                if _row_deleted(row):
                    continue
                cells = list(dict.fromkeys(_visible(c._tc).strip() for c in row.cells))
                text = " | ".join(c for c in cells if c)
                if not text:
                    continue
                label = f"Key Terms \u2013 {cells[0].strip('[] ')}" if terms_table else current
                out.append(Location(label, text, comments_on(row._tr)))
            continue
        if tag != "p" or _mark_deleted(element):
            continue
        paragraph = DocxParagraph(element, document)
        number = numberer.advance(*_paragraph_numbering(paragraph))
        text = _visible(element).strip()
        if not text:
            continue
        if number is not None and len(number.path) == 1 and _CONTAINER.match(number.full.strip()):
            container, part = number.full.strip(), None
            current = container
        elif number is not None and container and _PART.fullmatch(number.marker.strip()):
            part = number.marker.strip()
            current = f"{part} of {container}"
        elif number is not None:
            if container is None and not in_body and number.full.strip().startswith("("):
                current = "Parties"  # the parties are numbered (1), (2) before Clause 1
            elif container is None:
                in_body = True
                current = "Clause " + number.full.strip().replace(".(", "(").rstrip(".")
            else:
                own = own_number(number, container, part or "")
                if own:
                    where = f"{part} of {container}" if part else container
                    current = f"Paragraph {own} of {where}"
        elif (
            number is None
            and in_body  # the contents list before the clauses names every schedule too
            and (m := _TYPED_CONTAINER.match(text)) is not None
            and not (paragraph.style.name or "").startswith("TOC")
            and not text[-1].isdigit()  # a contents line ends in its page number
        ):
            container, part = f"{m.group(1).title()} {m.group(2)}", None
            current = container
        elif (m := _DEFINITION.match(text)) is not None:
            current = f'"{m.group(1).strip()}"'
        elif not in_body and container is None:
            current = "Cover page"
        out.append(Location(current, text, comments_on(element)))
    return out


def _comments(docx_bytes: bytes) -> dict[str, str]:
    """The draft's comments by id, each as one line of text."""
    import zipfile

    from lxml import etree

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if "word/comments.xml" not in z.namelist():
            return {}
        root = etree.fromstring(z.read("word/comments.xml"))
    return {
        c.get(qn("w:id")): " ".join("".join(t.text or "" for t in c.iter(qn("w:t"))).split())
        for c in root.iter(qn("w:comment"))
    }


# ---------------------------------------------------------------- the term sheet's items


@dataclass
class Item:
    id: str
    text: str
    anchor: DocxParagraph | None = None
    section: str = ""


def _blocks(element, document: DocumentType) -> Iterator[DocxParagraph | Table]:  # type: ignore[no-untyped-def]
    """Paragraphs and tables in layout order, opening content controls and the text boxes
    the lawyer sees (not Word's shadow copies of them)."""
    for child in element.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, document)
            for box in child.iterdescendants(qn("w:txbxContent")):
                if not any(a.tag == _FALLBACK for a in box.iterancestors()):
                    yield from _blocks(box, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)
        elif child.tag == qn("w:sdt"):
            for content in child.findall(qn("w:sdtContent")):
                yield from _blocks(content, document)


def _table_items(
    table: Table, document: DocumentType, seen: set[object]
) -> Iterator[tuple[str, DocxParagraph]]:
    """A row is one item, its note after its last filled cell; a table inside a cell is
    walked for its own rows. A merged cell is listed under every column it spans, so it is
    read once."""
    for row in table.rows:
        cells = []
        for cell in row.cells:
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            cells.append(cell)
        filled = [c for c in cells if c.text.strip()]
        if filled:
            last = filled[-1].paragraphs[-1]
            yield " | ".join(c.text.strip() for c in filled), last
        for cell in cells:
            for nested in _blocks(cell._tc, document):
                if isinstance(nested, Table):
                    yield from _table_items(nested, document, seen)


def termsheet_items(document: DocumentType) -> list[Item]:
    """The term sheet's Word file as items in order: each filled paragraph, each table row.
    What has been read is held as the elements themselves: an id() is reused once its
    proxy is collected, and paragraphs read later would pass as seen."""
    out: list[Item] = []
    seen: set[object] = set()
    unwrap_smart_tags(document.element.body)
    for block in _blocks(document.element.body, document):
        if isinstance(block, Table):
            for text, anchor in _table_items(block, document, seen):
                out.append(Item(f"L{len(out) + 1}", text, anchor))
        elif block.text.strip() and block._p not in seen:
            seen.add(block._p)
            out.append(Item(f"L{len(out) + 1}", block.text.strip(), block))
    return out


def record_items(sections: Sequence[dict]) -> list[Item]:
    """A term sheet held only as text (a signed PDF): its lines, section by section."""
    out: list[Item] = []
    for s in sections:
        for line in s["text"].split("\n"):
            if line.strip():
                out.append(Item(f"L{len(out) + 1}", line.strip(), None, s.get("heading", "")))
    return out


def with_sections(items: Sequence[Item], sections: Sequence[dict]) -> None:
    """A Word term sheet's items take the record's section that holds them, the nearest to where
    the reading is: an item that is a section's heading opens it, any other stays in the
    section it reads in. Nearest, not only onward: the record need not keep the term sheet's order."""
    heads = [_norm(s.get("heading", "")) for s in sections]
    texts = [_norm(s["text"]) for s in sections]

    def nearest(k: int, fits: Callable[[int], bool]) -> int | None:
        order = sorted(range(len(sections)), key=lambda j: (abs(j - k), j < k))
        return next((j for j in order if fits(j)), None)

    k, opened = 0, False
    for item in items:
        words = _norm(item.text)
        at = nearest(k, lambda j: heads[j] == words and not (opened and j == k))
        if at is not None:
            k, opened = at, True
        else:
            at = nearest(k, lambda j: words in texts[j])
            if at is not None and at != k:
                k, opened = at, False
        if sections:
            item.section = sections[k].get("heading", "")


def batches(items: Sequence[Item], size: int = BATCH_CHARS) -> list[list[Item]]:
    out: list[list[Item]] = [[]]
    used = 0
    for item in items:
        if out[-1] and used + len(item.text) > size:
            out.append([])
            used = 0
        out[-1].append(item)
        used += len(item.text)
    return [b for b in out if b]


# ---------------------------------------------------------------- the agent


class Note(BaseModel):
    item: str = Field(description="the term sheet item's id, as listed (L12)")
    places: list[str] = Field(
        default_factory=list,
        description="where the contract carries this item: each a place name copied exactly from "
        "search_contract or read_place, most specific first",
    )
    confirm: str = Field(
        default="",
        description="only where the contract has a blank where the term sheet gives a value, or an open "
        "point comment there leaving it to the lawyer: what the reviewer is to confirm, "
        "in a few words that read on from 'please confirm' ('whether the Supplier may be an affiliate', "
        "'the specification to be attached'); empty where the contract carries the item as the term sheet states it",
    )
    reason: str = Field(
        default="",
        description="only where no place carries the item: why, in a few words",
    )


class Notes(BaseModel):
    notes: list[Note] = Field(default_factory=list)


def _bracket(note: Note) -> str:
    if note.places:
        check = note.confirm.strip()
        return ", ".join(note.places) + (f" \u2013 please confirm {check}" if check else "")
    return note.reason.strip()


def annotate_batch(
    batch: Sequence[Item],
    locations: Sequence[Location],
    leads: Sequence[str],
    deal: str,
    llm,
    max_turns: int = 40,
    termsheet: Sequence[Item] = (),
    every: bool = False,
) -> dict[str, str]:
    """One agent reads a stretch of the term sheet and names, for each item that sets a term,
    the places of the final contract that carry it. Returns item id -> note. `termsheet` is the
    whole term sheet, shown in outline so the agent knows what heading an item sits under and
    what a heading covers beyond its stretch. With `every`, each item is the heading of a
    section no note reached, and each must be answered."""
    labels = list(dict.fromkeys(loc.label for loc in locations))
    known = set(labels)
    ids = {i.id for i in batch}

    def search_contract(words: str) -> list[dict]:
        """Places of the final contract whose text holds all these words (case ignored): the
        place's name, the start of its text, and any comment the draft puts on it. Search on a
        figure or a distinctive term."""
        wanted = _norm(words).split()
        hits = [loc for loc in locations if all(w in _norm(loc.text) for w in wanted)]
        return [
            {"place": h.label, "text": h.text[:300], **({"comments": list(h.comments)} if h.comments else {})}
            for h in hits[:12]
        ] or [
            {"place": "", "text": "(nowhere in the contract)"}
        ]

    def read_place(place: str) -> str:
        """The whole text of one place of the final contract, by its exact name, each comment the
        draft puts on a paragraph after it."""
        text = "\n".join(
            loc.text + "".join(f"\n  [comment: {c}]" for c in loc.comments)
            for loc in locations
            if loc.label == place
        )
        return text[:4000] if text else f"no place named {place!r}"

    class Checked(Notes):
        places_known: ClassVar[set[str]] = known
        items_known: ClassVar[set[str]] = ids

        @model_validator(mode="after")
        def real(self) -> Checked:
            for n in self.notes:
                if n.item not in self.items_known:
                    raise ValueError(f"{n.item!r} is not an item of this stretch of the term sheet")
                bad = [p for p in n.places if p not in self.places_known]
                if bad:
                    raise ValueError(
                        f"{bad[0]!r} is not a place of this contract; copy a name exactly as"
                        " search_contract or read_place gives it"
                    )
                if not n.places and not n.reason.strip():
                    raise ValueError(f"{n.item}: name a place, or give the reason none carries it")
            if every:
                left = sorted(self.items_known - {n.item for n in self.notes})
                if left:
                    raise ValueError(
                        f"{', '.join(left)}: each heads a section of the term sheet with no note yet;"
                        " name the places that carry it, or 'Not carried' with the reason"
                    )
            return self

    top = re.compile(r"(Clause \d+|Schedule \d+|Appendix [A-Z])")
    headings = [lab for lab in labels if top.fullmatch(lab)]
    first = {}
    for loc in locations:
        first.setdefault(loc.label, loc.text)
    outline = "\n".join(f"  {h}: {first[h][:70]}" for h in headings)
    task = (
        (
            f"THE WHOLE TERM SHEET for {deal}, in outline (each item's opening words;"
            " for where your stretch sits, not to be noted):\n"
            + "\n".join(f"  {i.id}: {_WS.sub(' ', i.text)[:90]}" for i in termsheet)
            + "\n\n"
            if termsheet
            else ""
        )
        + "YOUR STRETCH OF THE TERM SHEET, item by item, to be noted:\n"
        + "\n".join(f"  {i.id}: {i.text}" for i in batch)
        + "\n\nTHE FINAL CONTRACT, its clauses and schedules:\n"
        + outline
        + (
            "\n\nWHAT THE DRAFTERS SAID THEY WROTE FROM THIS STRETCH (leads only: a lead names"
            " the drafter's piece, not a place; find the place with search_contract):\n"
            + "\n".join(f"  - {ld}" for ld in leads)
            if leads
            else ""
        )
        + (
            "\n\nEACH ITEM ABOVE heads a section of the term sheet that no note reached. Give every"
            " one a note: the places of the contract that carry its section, or 'Not carried' with"
            " the reason (as the brief says for the term sheet's own housekeeping). Finish with"
            " submit_outcome."
            if every
            else "\n\nNote each item as the brief says. Finish with submit_outcome."
        )
    )
    out = run_agent(
        system_prompt=brief(),
        task_message=task,
        tools=[StructuredTool.from_function(f) for f in (search_contract, read_place)],
        final_output_schema=Checked,
        max_iterations=max_turns,
        llm=llm,
    )
    assert isinstance(out, Notes)
    # the reviewer reads each confirm as a question to the deal team: it is asked only where
    # the draft itself leaves the point open, as the writer marks an open point
    open_at = {loc.label for loc in locations if any("Open point:" in c for c in loc.comments)}
    notes = [
        n if open_at & set(n.places) else n.model_copy(update={"confirm": ""}) for n in out.notes
    ]
    return {n.item: _bracket(n) for n in notes if _bracket(n)}


def _fragments(quote: str) -> list[str]:
    return [f for f in (_norm(p).strip(" .\"'") for p in re.split(r"\.\.\.|…", quote)) if f]


def leads_for(batch: Sequence[Item], edits: dict) -> list[str]:
    """The drafters' edits whose term sheet quote is, word for word, in this stretch."""
    text = _norm(" ".join(i.text for i in batch))
    out = []
    for result in edits.get("results", []):
        for e in result.get("edits", []):
            frags = _fragments(e.get("termsheet_quote") or "")
            if frags and all(f in text for f in frags):
                words = (e.get("new_text") or e.get("old_text") or "").replace("\n", " ")
                out.append(f"{result['title']}: {words[:160]}")
    return list(dict.fromkeys(out))[:20]


def annotate(
    items: Sequence[Item],
    contract_docx: bytes,
    edits: dict,
    deal: str,
    llm,
    workers: int = 8,
) -> dict[str, str]:
    """Item id -> note, for the whole term sheet, the batches read side by side."""
    locations = contract_locations(contract_docx)
    work = batches(items)
    notes: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for found in pool.map(
            lambda b: annotate_batch(b, locations, leads_for(b, edits), deal, llm, termsheet=items),
            work,
        ):
            notes.update(found)
    # a section of the term sheet no note reached is answered once more, on its heading: the place
    # that carries it, or "Not carried" and why: every item gets a response
    # each run of items under one heading is its own section: a term sheet may repeat a
    # heading for a second set of terms
    runs: list[list[Item]] = []
    for item in items:
        if item.section and runs and runs[-1][0].section == item.section:
            runs[-1].append(item)
        elif item.section:
            runs.append([item])
    heads = [run[0] for run in runs if not any(i.id in notes for i in run)]
    for head in batches(heads):
        try:
            notes.update(annotate_batch(head, locations, [], deal, llm, termsheet=items, every=True))
        except AgentDidNotCompleteError:
            continue
    return notes


# ---------------------------------------------------------------- the document


def _safe(text: str) -> str:
    """Control characters a Word file cannot hold."""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def _append(paragraph: DocxParagraph, note: str) -> None:
    paragraph.add_run(" ")
    run = paragraph.add_run(f"[{_safe(note)}]")
    run.bold = True
    run.font.color.rgb = GREEN


def for_the_lawyer(note: str) -> bool:
    """Whether a note leaves something to the lawyer: a place to confirm, or a term no place
    carries because it is an open point. Read from the wording this module's notes take."""
    return "please confirm" in note or "for the lawyer" in note


def _points_first(document: DocumentType, items: Sequence[Item], notes: dict[str, str]) -> None:
    """A list at the head of the term sheet of every item a note leaves to the lawyer, each with
    its note, in the term sheet's order, so the reviewer can work through them."""
    flagged = [i for i in items if i.id in notes and for_the_lawyer(notes[i.id])]
    body = document.element.body
    first = next((c for c in body.iterchildren() if c.tag in (qn("w:p"), qn("w:tbl"))), None)
    if not flagged or first is None:
        return
    lines: list[DocxParagraph] = []
    heading = document.add_paragraph()
    heading.add_run("Points for the lawyer").bold = True
    lines.append(heading)
    for item in flagged:
        line = document.add_paragraph(_safe(_WS.sub(" ", item.text).strip()))
        _append(line, notes[item.id])
        lines.append(line)
    for line in lines:
        first.addprevious(line._p)


def render_in_place(
    document: DocumentType, source: bytes, items: Sequence[Item], notes: dict[str, str]
) -> bytes:
    """The term sheet's own Word file with each note after its item, and the points left to the
    lawyer listed at its head; nothing else touched. A note too wide for any line its
    paragraph has (a narrow cell, a text box) goes to the foot of the page, so the term sheet
    keeps its shape."""
    footnotes = FootnoteNotes(source)
    for item in items:
        if item.id in notes and item.anchor is not None:
            note = notes[item.id]
            if note_fits(item.anchor, f"[{_safe(note)}]"):
                _append(item.anchor, note)
            else:
                footnotes.add(item.anchor, f"[{_safe(note)}]")
    _points_first(document, items, notes)
    buffer = io.BytesIO()
    document.save(buffer)
    return footnotes.into(buffer.getvalue())


def render_rebuilt(title: str, items: Sequence[Item], notes: dict[str, str]) -> bytes:
    """A term sheet held only as text, laid out section by section with its notes."""
    document = Document()
    document.add_paragraph().add_run(title).bold = True
    section = None
    for item in items:
        if item.section != section:
            section = item.section
            if _norm(item.text) == _norm(section or ""):
                heading = document.add_paragraph()
                heading.add_run(_safe(item.text)).bold = True
                if item.id in notes:
                    _append(heading, notes[item.id])
                continue
        paragraph = document.add_paragraph(_safe(item.text))
        if item.id in notes:
            _append(paragraph, notes[item.id])
    _points_first(document, items, notes)
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()
