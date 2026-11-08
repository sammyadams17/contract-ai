"""How much room a paragraph has for a note, and whether the note fits in it.

A table cell is not the only shape that takes width away, which is what the first
version of this got wrong. A note can land in an ordinary body
paragraph that sits in a two-column section and carries a deep right indent, so it has
a few centimetres to render a long note in and Word sets it one character per line.
Nothing about those paragraphs is a cell or a text box.

So ask the question that actually matters - does the note fit on the line? - and let the
document answer it. Page size, margins, column count and indents are all stated in the
document, and the only estimate here is how wide a character is, which is a property of
type rather than of any one drafting house.
"""

from __future__ import annotations

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

__all__ = ["available_width", "note_fits"]

#: Mean character advance as a fraction of the font's point size. A property of
#: proportional type, not a threshold tuned to a document: half the point size is the
#: usual approximation for mixed-case text in a serif face.
_MEAN_ADVANCE = 0.5
#: Word's default when a document states no font size, in half-points.
_DEFAULT_HALF_POINTS = 20
_TWIPS_PER_POINT = 20
#: A text box is drawn at a size somebody chose and does not reflow, and its width is
#: not stated where a paragraph can read it. Nothing inline is safe inside one.
_NO_ROOM = 0


def _twips(node: object, attribute: str) -> int | None:
    value = node.get(qn(attribute)) if node is not None else None  # type: ignore[attr-defined]
    try:
        return int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None


def _column_width(sectPr: object) -> int:
    """The width one column of this section gives its text."""
    page_size = sectPr.find(qn("w:pgSz"))  # type: ignore[attr-defined]
    margins = sectPr.find(qn("w:pgMar"))  # type: ignore[attr-defined]
    width = (
        (_twips(page_size, "w:w") or 11906)
        - (_twips(margins, "w:left") or 0)
        - (_twips(margins, "w:right") or 0)
    )
    columns = sectPr.find(qn("w:cols"))  # type: ignore[attr-defined]
    count = _twips(columns, "w:num") or 1
    if count > 1:
        width = (width - (_twips(columns, "w:space") or 0) * (count - 1)) // count
    return max(width, 0)


def _section_of(paragraph: Paragraph) -> object | None:
    """The section a paragraph belongs to: the next `sectPr` at or after it."""
    node = paragraph._p
    while node is not None:
        pPr = node.find(qn("w:pPr")) if node.tag == qn("w:p") else None
        inner = pPr.find(qn("w:sectPr")) if pPr is not None else None
        if inner is not None:
            return inner
        node = node.getnext()
    body = paragraph._p.getroottree().getroot().find(qn("w:body"))
    return body.find(qn("w:sectPr")) if body is not None else None


def _enclosing_width(paragraph: Paragraph) -> int | None:
    """The width of the nearest shape that sets one, or None if nothing does."""
    node = paragraph._p.getparent()
    while node is not None:
        if node.tag == qn("w:txbxContent"):
            return _NO_ROOM
        if node.tag == qn("w:tc"):
            tcPr = node.find(qn("w:tcPr"))
            cell_width = tcPr.find(qn("w:tcW")) if tcPr is not None else None
            if cell_width is not None and cell_width.get(qn("w:type")) in (None, "dxa"):
                return _twips(cell_width, "w:w") or _NO_ROOM
            return _NO_ROOM
        node = node.getparent()
    return None


def available_width(paragraph: Paragraph) -> int:
    """The twips this paragraph has to set text in."""
    enclosed = _enclosing_width(paragraph)
    if enclosed is not None:
        return enclosed
    section = _section_of(paragraph)
    width = _column_width(section) if section is not None else 0
    pPr = paragraph._p.find(qn("w:pPr"))
    indent = pPr.find(qn("w:ind")) if pPr is not None else None
    if indent is not None:
        width -= (_twips(indent, "w:left") or 0) + (_twips(indent, "w:right") or 0)
    return max(width, 0)


def _point_size(paragraph: Paragraph) -> float:
    """The size the paragraph's own text is set in."""
    for run in paragraph.runs:
        if run.font.size is not None:
            return float(run.font.size.pt)
    size = paragraph.style.font.size if paragraph.style is not None else None
    return float(size.pt) if size is not None else _DEFAULT_HALF_POINTS / 2


def note_fits(paragraph: Paragraph, note: str) -> bool:
    """Would the note set on a line of its own in the width the paragraph has?

    Not "is there room beside the existing text" - wrapping onto the next line is what
    prose does and costs nothing. What breaks a layout is a note too wide for any line,
    because then Word sets it one character per line and the column has to open. So the
    question is only whether the note is narrower than the paragraph's own width.
    """
    return _text_width(paragraph, note) <= available_width(paragraph)


def _text_width(paragraph: Paragraph, text: str) -> int:
    advance = _point_size(paragraph) * _MEAN_ADVANCE * _TWIPS_PER_POINT
    return int(len(text) * advance)
