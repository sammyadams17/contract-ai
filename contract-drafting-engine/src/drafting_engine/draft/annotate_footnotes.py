"""Notes that cannot live inline, moved to the foot of the page.

A green bracket appended to the paragraph is the right annotation in running prose: most
notes sit there and read well. The rest do not: they are in table cells and text boxes,
shapes whose width their author fixed. Most carry a note longer than the cell's own text,
and a narrow cell may hold a single short word - so no wording is short enough to fit,
and the column has to open. The
table the lawyer laid out is then not the table he gets back.

A footnote reference is one superscript numeral, so the paragraph keeps its width
whatever the note says, and Word lays the note out at the foot of the page across the
page's own width.

`annotate_widths` decides which notes come here: the ones too wide for any line the
paragraph has.

Word will not let python-docx manage the footnotes part, so the entries are written into
the saved package directly. A document that has no footnotes part yet is given one,
declared in `[Content_Types].xml` and related from the document.
"""

from __future__ import annotations

import io
import zipfile
from xml.sax.saxutils import escape

from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.text.paragraph import Paragraph
from lxml import etree

__all__ = ["FootnoteNotes"]

_W = nsmap["w"]
_PART = "word/footnotes.xml"
_RELS = "word/_rels/document.xml.rels"
_TYPES = "[Content_Types].xml"
_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
#: Word's own character style for the reference mark, so the numeral superscripts and
#: the document's look is its own rather than ours.
_REFERENCE_STYLE = "FootnoteReference"

#: The declaration is left off: `etree.tostring` writes one when the part is serialised,
#: and lxml refuses to parse a string that carries its own.
_EMPTY_PART = (
    f'<w:footnotes xmlns:w="{_W}">'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p>'
    "</w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)

_FOOTNOTE = (
    f'<w:footnote xmlns:w="{_W}" w:id="{{id}}">'
    '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
    f'<w:r><w:rPr><w:rStyle w:val="{_REFERENCE_STYLE}"/></w:rPr><w:footnoteRef/></w:r>'
    '<w:r><w:rPr><w:color w:val="00B050"/></w:rPr>'
    '<w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:footnote>'
)


def _reference_run(footnote_id: int) -> object:
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), _REFERENCE_STYLE)
    rPr.append(style)
    run.append(rPr)
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), str(footnote_id))
    run.append(reference)
    return run


class FootnoteNotes:
    """The notes bound for the foot of the page, and the part that will carry them."""

    def __init__(self, source_docx: bytes) -> None:
        self._existing = self._read_part(source_docx)
        self._notes: list[tuple[int, str]] = []
        self._next_id = self._first_free_id(self._existing)

    @staticmethod
    def _read_part(source_docx: bytes) -> bytes | None:
        with zipfile.ZipFile(io.BytesIO(source_docx)) as package:
            if _PART not in package.namelist():
                return None
            return package.read(_PART)

    @staticmethod
    def _first_free_id(part: bytes | None) -> int:
        """Word reserves -1 and 0 for the separators, so a real footnote starts at 1."""
        if part is None:
            return 1
        used = [
            int(node.get(qn("w:id")))
            for node in parse_xml(part).findall(qn("w:footnote"))
            if node.get(qn("w:id"))
        ]
        return max([*used, 0]) + 1

    def add(self, paragraph: Paragraph, note: str) -> None:
        """Mark the paragraph with a reference and hold the note for the part."""
        paragraph._p.append(_reference_run(self._next_id))
        self._notes.append((self._next_id, note))
        self._next_id += 1

    def __len__(self) -> int:
        return len(self._notes)

    def into(self, saved_docx: bytes) -> bytes:
        """The saved package with the notes written into its footnotes part."""
        if not self._notes:
            return saved_docx
        part = self._build_part()
        with zipfile.ZipFile(io.BytesIO(saved_docx)) as source:
            names = source.namelist()
            written = io.BytesIO()
            with zipfile.ZipFile(written, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    data = source.read(item.filename)
                    if item.filename == _PART:
                        data = part
                    elif item.filename == _TYPES and _PART not in names:
                        data = _declare_content_type(data)
                    elif item.filename == _RELS and _PART not in names:
                        data = _relate_part(data)
                    target.writestr(item, data)
                if _PART not in names:
                    target.writestr(_PART, part)
        return written.getvalue()

    def _build_part(self) -> bytes:
        root = parse_xml(self._existing if self._existing is not None else _EMPTY_PART)
        for footnote_id, note in self._notes:
            root.append(parse_xml(_FOOTNOTE.format(id=footnote_id, text=escape(note))))
        xml: bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        return xml


def _declare_content_type(types_xml: bytes) -> bytes:
    root = parse_xml(types_xml)
    override = etree.SubElement(root, f"{{{root.nsmap[None]}}}Override")
    override.set("PartName", f"/{_PART}")
    override.set("ContentType", _CONTENT_TYPE)
    xml: bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return xml


def _relate_part(rels_xml: bytes) -> bytes:
    root = parse_xml(rels_xml)
    namespace = root.nsmap[None]
    used = {child.get("Id", "") for child in root if child.tag == f"{{{namespace}}}Relationship"}
    number = 1
    while f"rId{number}" in used:
        number += 1
    relationship = etree.SubElement(root, f"{{{namespace}}}Relationship")
    relationship.set("Id", f"rId{number}")
    relationship.set("Type", _REL_TYPE)
    relationship.set("Target", "footnotes.xml")
    xml: bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return xml
