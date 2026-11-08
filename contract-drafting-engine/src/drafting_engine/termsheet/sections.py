"""Proposes the term sheet's headings by reading it.

A line-based rule cannot do this. A length cap drops a long heading; a structural rule (a
numbered line alone in its own block) cuts some term sheets to a fraction of their sections
and reads another's appendix table rows as headings. Layouts differ too much for one rule.

So a model reads the document. What it may return is deliberately narrow: **headings
only, verbatim**. It does not return body text, because text a model writes is text it
could have garbled or invented, and the body is what every later stage quotes from.
Bodies are cut from the term sheet by position instead (checklist.build_checklist).

It must also report what it declined, so a near-miss can be seen rather than vanish.
"""

from __future__ import annotations

from pydantic import BaseModel, Field

from drafting_engine.llm.agent import run_agent
from drafting_engine.llm.prompts import build_system_prompt
from drafting_engine.termsheet.checklist import ChecklistItem, build_checklist, merge_headings
from drafting_engine.word.instruction_coverage import cut_sections

PROMPT_EXTENSION = """
ROLE
You are the Checklist Proposal Agent.

OBJECTIVE
List every heading in the term sheet under which a term is stated, so each becomes
one item on a checklist a lawyer will walk while drafting.

REQUIRED BEHAVIOUR
- Return each heading EXACTLY as it appears in the term sheet, character for
  character, including its number and any punctuation. It is located by exact search; a
  heading you paraphrase or tidy cannot be found and is reported as an error.
- Include headings of any length, headings that are not numbered, and headings that
  introduce a table. Terms stated inside appendix tables count - an appendix listing
  the services and their fees is a term of the deal.
- Do NOT return body text. Only headings.
- A numbered line that continues a sentence is not a heading. "4. Completion actually takes
  place, then and only then the Buyer undertakes to pay..." is prose inside a clause.
- List in `declined` any line you considered and rejected, with a short reason, so a
  near-miss appears here rather than vanishing.
- Do not judge whether a term matters. Completeness is the whole purpose; a heading you
  omit is work nobody will be asked to do.
"""


class DeclinedLine(BaseModel):
    line: str = Field(description="The line, as it appears in the term sheet")
    reason: str = Field(description="Why it is not a heading")


class ChecklistProposalOutcome(BaseModel):
    headings: list[str] = Field(
        default_factory=list,
        description="Every heading, verbatim from the term sheet, in reading order",
    )
    declined: list[DeclinedLine] = Field(
        default_factory=list, description="Lines considered and rejected, with reasons"
    )
    summary: str = Field(default="", description="One sentence on how the document was laid out")


class ProposedChecklist(BaseModel):
    items: list[ChecklistItem] = Field(default_factory=list)
    # Headings the model named that are not in the term sheet. Reported rather than
    # turned into empty items: a model naming something absent is exactly the failure
    # mode of letting it write the list, and the person confirming should see it.
    unlocatable: list[str] = Field(default_factory=list)
    declined: list[DeclinedLine] = Field(default_factory=list)
    summary: str = ""

    model_config = {"arbitrary_types_allowed": True}


def propose_checklist(termsheet_text: str, llm=None) -> ProposedChecklist:
    outcome = run_agent(
        system_prompt=build_system_prompt(PROMPT_EXTENSION),
        task_message=(
            "The term sheet follows. List every heading under which a term is stated, "
            "verbatim.\n\n" + termsheet_text
        ),
        tools=[],
        final_output_schema=ChecklistProposalOutcome,
        llm=llm,
    )
    # Both proposers contribute. The model reads structures a pattern cannot - appendices,
    # tables, headings numbered with a comma - and the pattern catches the renumbered annex
    # the model folds into its parent. Each proposer finds headings the other misses.
    # See merge_headings.
    headings = merge_headings(list(outcome.headings), [s.heading for s in cut_sections(termsheet_text)])
    items, unlocatable = build_checklist(termsheet_text, headings)
    return ProposedChecklist(
        items=items,
        unlocatable=unlocatable,
        declined=list(outcome.declined),
        summary=outcome.summary,
    )
