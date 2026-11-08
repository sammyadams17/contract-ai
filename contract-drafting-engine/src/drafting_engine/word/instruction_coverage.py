"""The term sheet cut into the sections it numbers itself: the checklist's raw material.

The checklist proposer names the headings; this finds them in the text and slices between them,
so each section carries exactly the words beneath its heading.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# A numbered heading at the start of a line: "11. Payment Terms".
#
# The number is captured for display only and is never parsed or validated. A
# term sheet may RESTART its numbering in an annex - a second "1." after "20." - which is
# the same thing contracts do, restarting
# numbering in each Schedule. It may also SKIP a number while that section's content is
# present. Treating either as corruption would refuse to slice a perfectly sliceable
# document.
_HEADING_RE = re.compile(r"^[ \t]*(\d{1,2}\.[ \t]+[A-Z][^\n]{0,60}?)[ \t]*$", re.M)

# A Markdown heading, which is how the DOCX converter renders a heading whose number Word
# supplies automatically: "# Payment Terms". Such numbers are not in the text at all, so
# nothing may depend on them - the same holds in the contracts.
#
# Preferred over the numbered pattern rather than merged with it, which matters: an
# appendix's checklist is full of lines that DO carry literal numbers ("1. Signed copy
# received"), so merging would readmit exactly the items that caused the
# silent failure. When a document marks its structure, that marking is the better evidence.
_MARKDOWN_HEADING_RE = re.compile(r"^[ \t]*#{1,6}[ \t]+(\S[^\n]{0,70}?)[ \t]*$", re.M)


@dataclass(frozen=True)
class TermSheetSection:
    """One numbered part of a term sheet, as the document itself numbers them."""

    heading: str
    text: str
    start: int
    end: int


def cut_sections(termsheet_text: str) -> list[TermSheetSection]:
    """Cut the term sheet at each heading, in reading order.

    Two conventions, because real documents use both. A term sheet whose numbering is
    literal text slices on "11. Payment Terms"; one whose numbering Word supplies
    automatically arrives from the converter as "# Payment Terms" and has no numbers to slice
    on at all. A term sheet of the second kind, against the numbered pattern alone,
    yields ZERO body sections.

    Returns an empty list when there are no headings to cut at, so a caller can fall
    back to treating the document as one unit rather than receiving a single bogus
    section covering everything.
    """
    matches = list(_MARKDOWN_HEADING_RE.finditer(termsheet_text)) or list(_HEADING_RE.finditer(termsheet_text))
    sections = []
    for i, match in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(termsheet_text)
        sections.append(
            TermSheetSection(
                heading=match.group(1).strip(),
                text=termsheet_text[match.start() : end],
                start=match.start(),
                end=end,
            )
        )
    return sections
