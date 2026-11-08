"""A document as pages, and its pages as clauses.

A PDF is read page by page from its text layer, so every clause knows the pages it is written
on and a citation can open the page. A Word file has no pages until it is printed: each of its
paragraphs is placed on page 1, and its citations name the clause rather than a page.
"""

from __future__ import annotations

import pathlib
import re
from dataclasses import dataclass

# A clause begins at a line that numbers itself: "5", "5.3", "12.1.4", with or without a stop,
# then its heading or its first words. The numbering is the document's own structure.
_NUMBERED = re.compile(r"^\s*(\d{1,3}(?:\.\d{1,3}){0,3})\.?\s+(\S.{0,200})$")
_SCHEDULE = re.compile(r"^\s*((?:SCHEDULE|Schedule|APPENDIX|Appendix|ANNEX|Annex)\s+[\dA-Z]{1,3})\b\s*[-:–]?\s*(.*)$")


@dataclass
class Clause:
    number: str  # as the document numbers it; "" for text before the first numbered line
    heading: str
    text: str
    first_page: int
    last_page: int


def pages_of(path: pathlib.Path) -> list[str]:
    if path.suffix.lower() == ".pdf":
        from pypdf import PdfReader

        return [page.extract_text() or "" for page in PdfReader(str(path)).pages]
    if path.suffix.lower() == ".docx":
        from docx import Document

        return ["\n".join(p.text for p in Document(str(path)).paragraphs)]
    if path.suffix.lower() in (".txt", ".md"):
        return path.read_text().split("\f")  # a form feed is a page break, as printed text marks one
    raise ValueError(f"cannot read {path.suffix} files: {path.name}")


def clauses_of(pages: list[str]) -> list[Clause]:
    """The document cut at each numbered line and schedule heading, in reading order. A heading
    whose words run on to the next line keeps them in its text."""
    out: list[Clause] = []
    current = Clause("", "", "", 1, 1)
    headings: dict[str, str] = {}
    for page_no, page in enumerate(pages, start=1):
        for line in page.split("\n"):
            m = _NUMBERED.match(line) or _SCHEDULE.match(line)
            if m:
                if current.text.strip():
                    out.append(current)
                number, rest = m.group(1), m.group(2).strip()
                top = "." not in number
                if top:
                    heading = rest if len(rest) <= 80 and not rest.endswith((".", ";", ",")) else ""
                    headings[number.split(".")[0]] = heading
                else:
                    # a sub-clause is searched under its clause's heading: "4.1" is the Deposit
                    heading = headings.get(number.split(".")[0], "")
                current = Clause(number, heading, line.strip(), page_no, page_no)
                continue
            if line.strip():
                current.text += ("\n" if current.text else "") + line.strip()
                current.last_page = page_no
    if current.text.strip():
        out.append(current)
    return out
