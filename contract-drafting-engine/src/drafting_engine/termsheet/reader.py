"""The term sheet reader agent: one agent per section, with the whole term sheet in reach.

The model has tools over the entire document, follows a reference to where it points instead of recording the
reference, records each term through a tool that checks the quote on the spot, and may not finish
until a code check has named every span with a figure or a defined term that no quote covers.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from drafting_engine.llm.agent import run_agent
from drafting_engine.termsheet.brief import system_prompt as reader_brief
from drafting_engine.termsheet.record import Fact, Section, plain
from drafting_engine.word.span_grounding import normalise_for_comparison

TOOLS = """

HOW YOU WORK
You have the whole term sheet in reach through tools. Read your section with read_section.
Where it points elsewhere, follow it with read_neighbours, search_termsheet or read_section. Record every
row with record_term. It checks the quote against the term sheet at once and refuses a quote
that is not verbatim: fix the quote and record again. Before you finish, call check_coverage: it
lists every sentence and table row of your section carrying a figure, a date, a percentage or a
defined term that none of your quotes covers. Either it belongs in a row you have not yet written
- write it, whole - or skip_span it with the reason a lawyer would give (a heading, a signature
line, a page-break artefact, words already covered by a row's quote). You may not submit while
check_coverage still lists spans you have neither recorded nor skipped. A section with no heading
of its own ("where", a lone formula row) belongs to the section before it: read the neighbour and
label the row by what it is about. Where several rows fall under one heading, label each by
what it settles, in the term sheet's words: "Fees: amount", "Fees: when payable", "Fees:
late payment". Never three rows all labelled "Fees".
"""


_FIGURE_OR_TERM = re.compile(
    r"\d|%|US\$|\$|\bTBD\b|to be (?:advised|agreed|confirmed)|as per\b"
    r"|[\"“(]\s*[A-Z][A-Za-z ]{2,}[\"”)]"
)
_SENTENCE_END = re.compile(r"(?<=[.;:])\s+(?=[A-Z(\"“])")
_OVERLAP_CHARS = 40


def _spans(text: str) -> list[str]:
    """Sentences and table rows of a section, as a reader would count them."""
    spans: list[str] = []
    for line in plain(text).splitlines():
        line = line.strip()
        if not line or re.fullmatch(r"[-|\s]+", line):
            continue
        if "|" in line:
            spans.append(line)
        else:
            spans.extend(s.strip() for s in _SENTENCE_END.split(line) if s.strip())
    return spans


def _covered(span: str, quotes: list[str]) -> bool:
    s = normalise_for_comparison(span)
    for q in quotes:
        if not q:
            continue
        if s in q or q in s:
            return True
        for start in range(0, len(s) - _OVERLAP_CHARS + 1):
            if s[start : start + _OVERLAP_CHARS] in q:
                return True
    return False


def uncovered_spans(section_text: str, quotes: list[str]) -> list[str]:
    """Spans of the section carrying a figure, a date, a percentage or a defined term that no
    quote covers. The first span is the heading line, which is skipped."""
    normalised = [normalise_for_comparison(q) for q in quotes]
    spans = _spans(section_text)[1:] if _spans(section_text) else []
    return [
        s
        for s in spans
        if _FIGURE_OR_TERM.search(s)
        and len(re.sub(r"\W", "", s)) >= 4  # "1." is numbering, not a term
        and not _covered(s, normalised)
    ]


@dataclass
class SectionReading:
    """One section's reading in progress: the tools' shared state."""

    sections: list[Section]
    current: int
    facts: list[Fact] = field(default_factory=list)
    skipped: dict[str, str] = field(default_factory=dict)
    refusals: int = 0

    def _section(self, index: int) -> Section | None:
        return next((s for s in self.sections if s.index == index), None)

    def _where(self, quote: str) -> int | None:
        """The section the quote occurs in: the current one first, then any."""
        q = normalise_for_comparison(quote)
        if not q:
            return None
        order = [self._section(self.current)] + [
            s for s in self.sections if s.index != self.current
        ]
        for section in order:
            if section and (
                q in normalise_for_comparison(section.text)
                or q in normalise_for_comparison(plain(section.text))
            ):
                return section.index
        return None

    # tools

    def list_sections(self) -> list[dict]:
        """Every section of the term sheet: index, heading, body or appendix, length."""
        return [
            {"index": s.index, "heading": s.heading, "part": s.part, "chars": len(s.text)}
            for s in self.sections
        ]

    def read_section(self, index: int) -> str:
        """The full text of one section by index. Your own section is `current`."""
        section = self._section(index)
        return plain(section.text) if section else f"no section {index}"

    def read_neighbours(self) -> dict:
        """The sections immediately before and after yours, for context a fragment lacks."""
        before, after = self._section(self.current - 1), self._section(self.current + 1)
        return {
            "before": plain(before.text)[-3000:] if before else "",
            "after": plain(after.text)[:3000] if after else "",
        }

    def search_termsheet(self, query: str, max_matches: int = 8) -> list[dict]:
        """Case-insensitive search of the whole term sheet, with the section and context."""
        matches: list[dict] = []
        needle = query.casefold()
        for section in self.sections:
            text = plain(section.text)
            low = text.casefold()
            start = 0
            while (at := low.find(needle, start)) >= 0 and len(matches) < max_matches:
                matches.append(
                    {
                        "section": section.index,
                        "heading": section.heading,
                        "context": text[max(0, at - 200) : at + len(query) + 300],
                    }
                )
                start = at + 1
        return matches

    def record_term(
        self,
        term: str,
        kind: Literal["value", "position"],
        role: Literal["term", "termsheet_only"],
        value: str,
        quote: str,
        open: bool = False,
        note: str = "",
    ) -> dict:
        """Record one row of the term sheet: the term sheet's own label, the value whole, the verbatim
        quote (checked now), role `termsheet_only` if it governs the term sheet itself, `open` if the term sheet
        leaves the point to the contract, and a note for anything a lawyer should be told."""
        where = self._where(quote)
        if where is None:
            self.refusals += 1
            return {
                "status": "refused",
                "reason": "quote not found verbatim in the term sheet; copy it exactly",
            }
        if where != self.current:
            # Every section has its own reader. A reader that follows a reference into another
            # section and records what it finds there makes the record carry that term twice.
            self.refusals += 1
            return {
                "status": "refused",
                "reason": (
                    f"that text is in section {where}, which has its own reader and will record"
                    " it. Record only what YOUR section states; where it points elsewhere, record"
                    " the pointer as your row's value."
                ),
            }
        self.facts.append(
            Fact(
                section=where,
                term=term,
                kind=kind,
                value=value,
                quote=quote,
                quote_verified=True,
                note=note,
                role=role,
                open=open,
            )
        )
        return {"status": "recorded", "section": where, "recorded_so_far": len(self.facts)}

    def check_coverage(self) -> dict:
        """Spans of your section with a figure or defined term that no recorded quote covers."""
        section = self._section(self.current)
        quotes = [f.quote for f in self.facts]
        spans = uncovered_spans(section.text, quotes) if section else []
        pending = [s for s in spans if s not in self.skipped]
        return {"uncovered": pending, "count": len(pending)}

    def skip_span(self, span: str, reason: str) -> dict:
        """Declare a check_coverage span as carrying no term, with a reason a lawyer accepts."""
        self.skipped[span] = reason
        return {"status": "skipped", "remaining": self.check_coverage()["count"]}


class ReadingOutcome(BaseModel):
    summary: str = Field(description="One sentence: what this section agrees")
    states_nothing: bool = Field(
        default=False, description="True if the section states no term at all"
    )


def build_tools(reading: SectionReading) -> list[StructuredTool]:
    return [
        StructuredTool.from_function(reading.list_sections),
        StructuredTool.from_function(reading.read_section),
        StructuredTool.from_function(reading.read_neighbours),
        StructuredTool.from_function(reading.search_termsheet),
        StructuredTool.from_function(reading.record_term),
        StructuredTool.from_function(reading.check_coverage),
        StructuredTool.from_function(reading.skip_span),
    ]


def read_section_with_tools(
    llm,
    sections: list[Section],
    section: Section,
    max_iterations: int = 60,
    vocabulary: str = "",
) -> SectionReading:
    reading = SectionReading(sections=sections, current=section.index)
    task = (
        f"Your section is index {section.index}: {section.heading!r} ({section.part}). "
        "Read it with read_section, follow what it points to, record every term with record_term, "
        "then call check_coverage and clear it before you submit.\n\n=== YOUR SECTION ===\n\n"
        + plain(section.text)
    )
    run_agent(
        system_prompt=reader_brief(vocabulary) + TOOLS,
        task_message=task,
        tools=build_tools(reading),
        final_output_schema=ReadingOutcome,
        max_iterations=max_iterations,
        llm=llm,
    )
    return reading
