"""The paragraphs of a clause as the drafter sees them and the writer places them.

The render of the Word document is markdown with marks the document does not hold: heading
hashes, footnote markers, cross-reference fields, cell anchors, comment bodies. This module
knows that syntax, and only that syntax, so a clause can be shown to the drafter as numbered
paragraphs and written back by the writer. Nothing here reads content.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from drafting_engine.termsheet.record import plain
from drafting_engine.word.span_grounding import normalise_for_comparison

# heading hashes, the generic "1. " or "* " the render gives a numbered or bulleted paragraph, footnote
# markers, bookmark anchors, comment marks. The render marks every numbered paragraph "1. ": a
# "2. " is a number typed into the paragraph's own words, and stays
_MARKUP = re.compile(r"^\s*#+\s*|^\s*1\.\s+|^\s*\*\s+|\[\^fn-\d+\]|\{#[^}]*\}|\[Com:\d+\]", re.M)
_XREF = re.compile(r"\[~([^~\]]*)~\]\(#[^)]*\)")  # a Word cross-reference field in the render
_LINK = re.compile(r"\[([^\]]+)\]\((?:mailto:|https?://)[^)]*\)")  # a hyperlink in the render
_CELL_ANCHOR = re.compile(r"\{#cell:[A-Za-z0-9]+\}")
_CRITIC = re.compile(r"\{==(.*?)==\}")  # a highlighted span in the render; its words stay
_COMMENT = re.compile(r"\{>>.*?<<\}")  # a comment body in the render; not document text
_FOOTNOTE = re.compile(r"\[\^fn-\d+\]")
PARA_BREAK = " // "  # how a paragraph break inside a table cell is shown to the model


def document_text(text: str) -> str:
    """The clause as the Word document holds it: the render's heading marks, footnote markers,
    anchors and bold/italic marks are not in the document."""
    text = _MARKUP.sub("", text)
    text = _XREF.sub(r"\1", text)
    text = _LINK.sub(r"\1", text)
    text = _COMMENT.sub("", text)
    text = _CRITIC.sub(r"\1", text)
    return plain(text).strip()


@dataclass
class Para:
    id: str
    text: str  # as the paragraph now reads (no rendering marks)
    cells: list[str] | None = None  # anchored cells, for a table row
    inserted: tuple[int, int] | None = None  # (edit index, paragraph index) when this pass added it
    original: str = ""  # as the form has it; every change to a form paragraph is derived from this
    xrefs: list[str] = field(default_factory=list)  # visible text of each cross-reference field
    heading: int = 0  # the render's heading level (# = 1), 0 for a body paragraph or a row
    heading_like: str = ""  # for a new heading: the form heading whose style it takes
    beneath: bool = False  # for a new body paragraph: one numbering level below its anchor
    beneath_of: str = ""  # that anchor's id, where the piece added the anchor itself
    like: str = ""  # restyled: the form paragraph whose style and numbering it takes
    like_struck: bool = False  # that form paragraph is one the piece struck
    moved_from: str = ""  # for a paragraph move() placed: the id of the one it was

    def __post_init__(self) -> None:
        if not self.original and self.inserted is None:
            self.original = self.text


@dataclass
class Rewriting:
    """The working state of one clause: its paragraphs as they now read, the edits so far."""

    paras: list[Para]
    edits: list[dict] = field(default_factory=list)
    clause_text: str = ""
    deleted: list[Para] = field(default_factory=list)  # form paragraphs the drafter removed


def _accepted(line: str) -> str:
    """A line of the anchored render as it reads with every tracked change accepted: the
    render of an already-edited document marks deletions, insertions and comments inline."""
    line = re.sub(r"\{--.*?--\}", "", line)
    line = re.sub(r"\{>>.*?<<\}", "", line)
    return re.sub(r"\{\+\+(.*?)\+\+\}", r"\1", line)


def _table_rows(anchored_text: str) -> list[tuple[str, int, list[str]]]:
    """Each table row of the anchored render as (key of its first line, lines it spans, cells).
    A cell holding several paragraphs spans several lines; its anchor sits after the last, so the
    row is complete when every cell has its anchor. The first cell's earlier paragraphs come
    before the row's line: a table starts with a header and a separator, so a line between two
    rows of a table, the second no header, is part of the second."""
    lines = [_accepted(x) for x in anchored_text.split("\n")]
    rows = []
    lead: list[str] = []  # lines since the last row of a table
    table = False
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.fullmatch(r"[\s|:-]+", line) and "-" in line:
            table, lead = True, []
        elif " | " in line and _CELL_ANCHOR.search(line) and not re.fullmatch(r"[\s|:-]+", line):
            joined, n, j = line, 1, i + 1
            while len(_CELL_ANCHOR.findall(joined)) < joined.count(" | ") + 1 and j < len(lines):
                if lines[j].strip():
                    joined += "\n\n" + lines[j]
                    # a line that is only a cell's closing anchor has no counterpart in the
                    # clean text, so it is not a line the row spans there; nor one of only
                    # anchors and separators (" | " closing empty cells): the clean text's
                    # lines drop it, and counting it swallowed the next row's first line
                    if re.sub(r"[\s|]", "", _CELL_ANCHOR.sub("", lines[j])):
                        n += 1
                j += 1
            following = next((x for x in lines[j:] if x.strip()), "")
            header = bool(re.fullmatch(r"[\s|:-]+", following)) and "-" in following
            if table and lead and not header:
                joined, n, line = "\n\n".join([*lead, joined]), n + len(lead), lead[0]
            key = normalise_for_comparison(document_text(_CELL_ANCHOR.sub("", line)))
            rows.append((key, n, [c.strip() for c in joined.split(" | ")]))
            table, lead = True, []
            i = j
            continue
        elif _CELL_ANCHOR.search(line):
            lead = []  # a row of one cell: complete in itself
        elif line.strip() and table:
            lead.append(line)
        i += 1
    return rows


def paragraph_text(line: str) -> str:
    """The paragraph as the engine's text holds it: rendering marks removed, but footnote
    markers kept, because they sit in the text and a target that skips them is not found."""
    kept = _FOOTNOTE.sub(lambda m: f"\x00{m.group(0)[2:-1]}\x00", line)
    return re.sub(r"\x00(fn-\d+)\x00", lambda m: f"[^{m.group(1)}]", document_text(kept)).strip()


def paragraphs_of(clause_text: str, anchored_text: str) -> list[Para]:
    """One Para per paragraph of the clause; a table row is one Para carrying its anchored cells,
    however many lines its cells span. Rows are matched to the anchored render in order, so two
    rows reading alike each keep their own cells."""
    anchored = _table_rows(anchored_text)
    cursor = 0
    lines = [x for x in clause_text.split("\n") if x.strip()]
    # The render appends the footnote texts, endnotes and a read-only structure listing; none
    # is a paragraph of the document.
    for i, ln in enumerate(lines):
        if ln.strip() == "<!-- READONLY_BOUNDARY_START -->" or re.fullmatch(
            r"#*\s*(Footnotes|Endnotes)\s*", ln
        ):
            lines = lines[:i]
            break
    out: list[Para] = []
    matched: list[tuple[int, int]] = []  # (anchored row, the paragraph it is) in order
    i = 0
    while i < len(lines):
        line = lines[i]
        i += 1
        if re.fullmatch(r"[\s|:-]+", line):
            # a table's separator, or a header of empty cells the render made up; a row of
            # empty cells the document has is a row, and its cells carry anchors
            real = cursor < len(anchored) and anchored[cursor][0] == normalise_for_comparison(
                document_text(line)
            )
            if "-" in line or not real:
                continue
        if re.match(
            r"^\[\^fn-\d+\]:|^- _Ref\w+ →|^\s+- Referenced from:|^- \[(?:Error|Info|Warning)\]",
            line,
        ):
            continue  # the render's footnote list, reference index and diagnostics
        text = paragraph_text(line)
        if not text:
            continue
        cells = None
        key = normalise_for_comparison(document_text(line))
        hit = None
        if " | " in line:
            hit = next((k for k in range(cursor, len(anchored)) if anchored[k][0] == key), None)
        elif i > 1 and " | " in lines[i - 2]:
            # the first paragraph of a row's first cell, on a line of its own after a row; only
            # such a row is keyed by a line without a cell separator
            hit = next((k for k in range(cursor, len(anchored)) if anchored[k][0] == key), None)
        if hit is not None:
            _, n, cells = anchored[hit]
            cursor = hit + 1
            matched.append((hit, len(out)))
            i += n - 1  # the row's further lines are part of this paragraph
            text = " | ".join(
                PARA_BREAK.join(
                    piece
                    # stripped first: a cell reading "2. " is a quantity, not a list marker
                    for piece in (paragraph_text(x.strip()) for x in cell.split("\n\n"))
                    if piece.strip()
                )
                for cell in (_CELL_ANCHOR.sub("", c) for c in cells)
            )
        # A footnote reference is a field too: the engine cannot move or retype it, so a change
        # splits around it exactly as around a cross-reference.
        marks = re.match(r"^\s*(#+)\s", line)
        out.append(
            Para(
                id=f"p{len(out) + 1}",
                text=text,
                cells=cells,
                heading=len(marks.group(1)) if marks and cells is None else 0,
                xrefs=_XREF.findall(line) + _FOOTNOTE.findall(line),
            )
        )
    return _with_empty_rows(out, anchored, matched)


def _with_empty_rows(paras: list[Para], anchored: list, matched: list[tuple[int, int]]) -> list[Para]:
    """The rows of empty cells the clean render leaves out, where the anchored render has them
    between this clause's rows, each before the next row it shows: ids of their own ("e1"), so
    the paragraphs keep theirs."""
    if not matched:
        return paras
    before: dict[int, list[Para]] = {}  # paragraph position -> the empty rows just before it
    count = 0
    for row in range(matched[0][0] + 1, matched[-1][0]):
        _, _, cells = anchored[row]
        if row in {r for r, _ in matched} or any(_CELL_ANCHOR.sub("", c).strip() for c in cells):
            continue
        at = next(p for r, p in matched if r > row)
        count += 1
        before.setdefault(at, []).append(
            Para(id=f"e{count}", text=" | ".join("" for _ in cells), cells=cells)
        )
    return [x for k, p in enumerate(paras) for x in [*before.get(k, []), p]]


def shown(paras: list[Para]) -> str:
    return "\n".join(f"[{p.id}] {p.text}" for p in paras)
