"""The term sheet as its author laid it out: paragraphs and tables in body order, table
cells kept apart, text boxes recovered where they sit.

Why not plain text: a table flattened to text comes out as a block of labels followed by a
block of amounts, and a reader assigns the amounts by guesswork. A record is only as good as
the page it was read from.
"""

from __future__ import annotations

import io
import pathlib
import re

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from drafting_engine.word.pdf_ingest import extract_pdf_text


def render_termsheet(path: str | pathlib.Path) -> str:
    """Reading-order text of a term sheet. Word files are walked element by element; a PDF
    falls back to the existing text extraction (with OCR where there is no text layer)."""
    path = pathlib.Path(path)
    if path.suffix.lower() == ".pdf":
        return extract_pdf_text(str(path))
    return _render_docx(path)


def _accept_tracked_changes(document) -> None:
    """A term sheet sent as a redline reads as its author meant it: insertions kept, deletions
    gone. `Paragraph.text` reads only a paragraph's own runs, so an insertion left wrapped in
    `w:ins` would be lost."""
    body = document.element.body
    for ins in list(body.iter(qn("w:ins"))):
        parent = ins.getparent()
        for child in list(ins):
            ins.addprevious(child)
        parent.remove(ins)
    for dele in list(body.iter(qn("w:del"))):
        dele.getparent().remove(dele)
    unwrap_smart_tags(body)


def unwrap_smart_tags(body) -> None:
    """A smart tag (Word's "place", "City") wraps runs the way an insertion does, and nests:
    `Paragraph.text` would read "Paris, France" as ", France". Its runs take its place."""
    for tag in reversed(list(body.iter(qn("w:smartTag")))):
        for child in list(tag):
            if child.tag != qn("w:smartTagPr"):
                tag.addprevious(child)
        tag.getparent().remove(tag)


def render_docx_bytes(data: bytes) -> str:
    """``render_termsheet`` for a Word term sheet held in memory."""
    return _render_document(Document(io.BytesIO(data)))


def _render_docx(path: pathlib.Path) -> str:
    return _render_document(Document(str(path)))


def _render_document(document) -> str:
    _accept_tracked_changes(document)
    lines: list[str] = []
    seen_boxes: set[str] = set()
    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            paragraph = Paragraph(child, document)
            text = paragraph.text.strip()
            boxes = _text_boxes(child, seen_boxes)
            if text:
                lines.append(text)
            lines.extend(boxes)
        elif tag == "tbl":
            lines.append("")
            lines.extend(_table_rows(Table(child, document)))
            lines.append("")
    return "\n".join(lines)


def _table_rows(table: Table) -> list[str]:
    """One line per row, cells separated by ' | ', a merged cell written once, an empty cell
    kept empty in its column, only a trailing column empty in every row dropped, nested tables
    read in place.

    A row whose cells each hold the same number of lines is several rows the author stacked in
    one: a rate table may have "Year 1 / Year 2 | 100 / 200" in one row, and pairing them by
    position is what a reader does by eye. A layout grid of tables nested in one, most cells
    empty, is lost entirely by `cell.text` alone.
    """
    rows: list[str] = []
    # A cell merged downwards is the same element on every row it spans. Its paragraphs are
    # written once: one per row where the author wrote one line per spanned row (a
    # date over the next row's count), otherwise all on the first row.
    # The elements are held in a list and compared by identity: lxml hands out a fresh proxy
    # on each access, so neither `id()` nor hashing tells one cell from another reliably.
    elements: list = []
    span_rows: list[list[int]] = []
    for r, row in enumerate(table.rows):
        for cell in row.cells:
            k = next((i for i, e in enumerate(elements) if e is cell._tc), None)
            if k is None:
                elements.append(cell._tc)
                span_rows.append([r])
            else:
                span_rows[k].append(r)
    table_cells: list[list[list[str]]] = []
    for r, row in enumerate(table.rows):
        cells: list[list[str]] = []
        placed: list = []
        for cell in row.cells:
            if any(cell._tc is e for e in placed):
                continue
            placed.append(cell._tc)
            spanned = span_rows[next(i for i, e in enumerate(elements) if e is cell._tc)]
            lines = _cell_lines(cell._tc)
            if len(spanned) > 1:
                if len(lines) == len(spanned):
                    lines = [lines[spanned.index(r)]]
                elif r != spanned[0]:
                    lines = []
            cells.append(lines)
        if any(cells):
            table_cells.append(cells)
    # a trailing column empty in every row goes; a row's own blank last cell stays, so every
    # row keeps the table's width (rows with no reference number came out a cell short)
    width = max((max(i for i, c in enumerate(row) if c) + 1 for row in table_cells), default=0)
    for cells in table_cells:
        cells = cells[:width]
        filled = [c for c in cells if c]
        counts = {len(c) for c in filled}
        if len(counts) == 1 and counts.pop() > 1 and len(filled) > 1:
            for i in range(len(filled[0])):
                rows.append(" | ".join(c[i] if c else "" for c in cells))
        else:
            rows.append(" | ".join(" / ".join(c) for c in cells))
    return rows


def _cell_lines(tc) -> list[str]:
    """The lines a cell shows: its paragraphs, and the rows of any table nested inside it."""
    lines: list[str] = []
    for child in tc.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            if text:
                lines.append(text)
        elif child.tag == qn("w:tbl"):
            lines.extend(_table_rows(Table(child, None)))
    return lines


def _text_boxes(paragraph_element, seen: set[str]) -> list[str]:
    """Text boxes anchored in this paragraph, once each. Word stores a drawing twice - a
    modern Choice and a legacy Fallback - and both carry the same words."""
    out: list[str] = []
    for box in paragraph_element.iter(qn("w:txbxContent")):
        parts: list[str] = []
        for child in box.iterchildren():
            if child.tag == qn("w:p"):
                text = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
                if text:
                    parts.append(text)
            elif child.tag == qn("w:tbl"):
                parts.extend(_table_rows(Table(child, None)))
        text = "\n".join(parts).strip()
        key = re.sub(r"[^A-Za-z0-9]", "", text)
        if key and key not in seen:
            seen.add(key)
            out.append(text)
    return out
