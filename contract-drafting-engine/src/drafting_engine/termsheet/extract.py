"""Read a term sheet into a Record: one reader agent per section, in parallel. Every quote is
checked by code before the record is frozen.

Why per section and not whole-document: a whole-document reading returns a different number of
facts from byte-identical input on every run, and drops the ones a section read on its own finds.
A section read on its own competes with nothing.
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor

from drafting_engine.termsheet.reader import SectionReading, read_section_with_tools
from drafting_engine.termsheet.record import Fact, Record, Section, dedupe, verify_facts
from drafting_engine.termsheet.sections import propose_checklist


def enumerate_sections(termsheet_text: str, llm) -> tuple[list[Section], list[str]]:
    """The term sheet's own sections, headings verbatim, bodies cut by position."""
    proposed = propose_checklist(termsheet_text, llm=llm)
    return sections_from(termsheet_text, proposed), list(proposed.unlocatable)


def sections_from(termsheet_text: str, proposed) -> list[Section]:
    """Sections from the proposed headings, plus the text above the first heading.

    A term sheet's opening paragraph often names the parties before any heading. The cutter
    starts at the first heading, so without a preamble section those terms reach no reader.
    """
    items = list(proposed.items)
    # The first item's text starts where the cutter anchored its heading; searching for the
    # heading alone can find the same word earlier, in a letterhead.
    first_at = termsheet_text.find(items[0].text) if items else len(termsheet_text)
    preamble = termsheet_text[: max(first_at, 0)].strip()
    parts: list[tuple[str, str]] = ([("Preamble", preamble)] if preamble else []) + [
        (item.heading, item.text) for item in items
    ]
    appendix_from = next(
        (i for i, (heading, _) in enumerate(parts) if _names_an_appendix(heading)), len(parts)
    )
    return [
        Section(
            index=i,
            heading=heading,
            text=text,
            part="appendix" if i >= appendix_from else "body",
        )
        for i, (heading, text) in enumerate(parts)
    ]


def _names_an_appendix(heading: str) -> bool:
    words = heading.lower().lstrip("#*_ ").split()
    return bool(words) and words[0] in ("appendix", "annex", "schedule")


def build_record(
    deal: str,
    termsheet_text: str,
    llm,
    model_name: str,
    workers: int = 6,
    vocabulary: str = "",
) -> Record:
    sections, unlocatable = enumerate_sections(termsheet_text, llm)
    notes = [f"heading named by the proposer but not found in the term sheet: {h}" for h in unlocatable]

    def one(section: Section) -> SectionReading | str:
        try:
            return read_section_with_tools(llm, sections, section, vocabulary=vocabulary)
        except Exception as error:  # a dead section is reported, not lost
            return (
                f"section {section.index} '{section.heading}' did not read:"
                f" {type(error).__name__}: {error}"
            )

    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(pool.map(one, sections))

    facts: list[Fact] = []
    gaps: list[dict] = []
    for result in results:
        if isinstance(result, str):
            notes.append(result)
            continue
        facts.extend(result.facts)
        # What the reader declared as carrying no term, and what it left uncovered when it
        # stopped: both are for a person to look at, neither is silently fine.
        gaps += [
            {"section": result.current, "span": span, "reason": reason}
            for span, reason in result.skipped.items()
        ]
        gaps += [
            {"section": result.current, "span": span, "reason": "UNRESOLVED: the reader stopped"}
            for span in result.check_coverage()["uncovered"]
        ]
        if result.refusals:
            notes.append(
                f"section {result.current}: {result.refusals} quote(s) refused as not verbatim"
            )

    return Record(
        deal=deal,
        model=model_name,
        sections=sections,
        facts=verify_facts(sections, dedupe(facts)),
        notes=notes,
        gaps=gaps,
    )
