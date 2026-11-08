"""The term sheet as a list of sections, each a unit of work to be accounted for.

Lawyers already do this on paper - they annotate the term sheet clause by clause as they
draft, walking it as a list. This is that list, made explicit, so that "we do not know what
happened to clause 14" cannot be said: a section exists whether or not anything was done
with it, so it cannot vanish by being skipped.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass


@dataclass(frozen=True)
class ChecklistItem:
    """One term sheet Section as a unit of work.

    Identified by index rather than by heading text, because an annex may reuse a title
    and a term sheet may number two sections the same - two clause 1s.
    """

    index: int
    heading: str
    text: str
    fact_keys: tuple[str, ...] = ()


def build_checklist(termsheet_text: str, headings: list[str]) -> tuple[list[ChecklistItem], list[str]]:
    """Turn proposed headings into items, cutting each body from the term sheet.

    A heading must be found verbatim in the term sheet to become an item. The model
    proposing headings may read tables, unnumbered titles and titles of any length - the
    freedom a model-read proposal bought - but it may not supply body text, because text it wrote is
    text it could have garbled or invented. Bodies are cut by position instead.

    Returns the items in document order, and the proposed headings that could not be
    located. An unlocatable heading is reported rather than turned into an empty item:
    that is the model naming something that is not there, and it should be visible.
    """
    located: list[tuple[int, str]] = []
    unlocatable: list[str] = []
    taken: list[int] = []
    for heading in headings:
        # A repeated heading is a distinct item - a term sheet may number two sections "1." and an
        # annex reuses titles - so each search starts after the last position claimed.
        # An occurrence that heads its own line wins over one inside a sentence: a
        # body may say "set out in Appendix 2" long before the heading "# Appendix 2",
        # and cutting at the mention filed the whole appendix under the section before it.
        occurrences = _occurrences(termsheet_text, heading)
        anchored = [at for at in occurrences if _heads_a_line(termsheet_text, at, heading)]
        # Two headings on one line are one heading: the proposer emits a section twice, once
        # as "# Sanctions Compliance" and once bare, two characters along the same line.
        candidates = [at for at in (anchored or occurrences) if _line_of(termsheet_text, at) not in taken]
        # A heading that heads a line somewhere never falls back to a mention inside a sentence:
        # a proposer named "Equipment" twice, and the second cut the preamble at "the
        # supply of the Equipment to the Buyer". Only a heading with no line of its own anywhere
        # may be found mid-line, which is how a term sheet with run-in headings still cuts.
        if not candidates:
            unlocatable.append(heading)
            continue
        at = candidates[0]
        taken.append(_line_of(termsheet_text, at))
        located.append((at, heading))

    located.sort()
    items = []
    for i, (at, heading) in enumerate(located):
        end = located[i + 1][0] if i + 1 < len(located) else len(termsheet_text)
        items.append(ChecklistItem(index=i, heading=heading, text=termsheet_text[at:end]))
    return items, unlocatable


def _line_of(text: str, at: int) -> int:
    return text.rfind("\n", 0, at) + 1


def _occurrences(text: str, needle: str) -> list[int]:
    found, start = [], 0
    while (at := text.find(needle, start)) >= 0:
        found.append(at)
        start = at + 1
    return found


# What may sit between a line break and a heading, or after it: Markdown heading, bold and
# italic marks, whitespace, the converter's page-break tag, and an image (an "Appendix 2" may
# follow a logo on the same line).
_LINE_DRESSING = re.compile(r"^(?:\s|[#*_]|<w:br[^>]*/>|!\[[^\]]*\]\([^)]*\))*$")


def _heads_a_line(text: str, at: int, heading: str) -> bool:
    line_start = text.rfind("\n", 0, at) + 1
    line_end = text.find("\n", at + len(heading))
    line_end = len(text) if line_end < 0 else line_end
    before, after = text[line_start:at], text[at + len(heading) : line_end]
    return bool(_LINE_DRESSING.match(before) and _LINE_DRESSING.match(after))


_APPENDIX_START = re.compile(r"^(appendix|annex|schedule)\b", re.I)
# "Appendix 1 - Equipment Specification": it names itself, so it is the schedule and not a
# sentence pointing at one.
_APPENDIX_TITLED = re.compile(r"^(?:appendix|annex|schedule)\s*\d*\s*[-\u2013\u2014:]\s*\S", re.I)
_APPENDIX_NUMBER = re.compile(r"^(?:appendix|annex|schedule)\s*(\d+)", re.I)
# Nothing but punctuation, whitespace and Markdown fences - what is left of "Appendix 1  **"
# once its own heading is removed.
_MARKUP_ONLY = re.compile(r"^[\s\W_]*$")


def merge_headings(primary: list[str], secondary: list[str]) -> list[str]:
    """Combine two proposers' headings, keeping whatever either one saw.

    Neither proposer dominates. A model reads appendices, table headings and headings
    numbered with commas ("12, Fees") that a line-based pattern cannot see; the pattern
    finds a renumbered annex the model folds into its parent heading.

    Over-inclusion is nearly free here: a spurious section is read and found to state
    nothing. Omission is invisible. So two cheap proposers beat one clever one.

    Merged by count rather than by set union, because a repeated heading is a distinct item
    - a term sheet may number two sections "1." and an annex reuses titles - and collapsing those
    would drop an item silently, the very failure being guarded against. Order follows
    `primary`, with anything only `secondary` saw appended; final document order is settled
    by build_checklist, which locates each heading anyway.
    """
    merged = list(primary)
    primary_counts = Counter(primary)
    for heading, count in Counter(secondary).items():
        merged += [heading] * max(0, count - primary_counts[heading])
    return merged
