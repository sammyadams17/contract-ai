"""The whole pipeline, offline: plan, draft, check, write, comment, with a scripted model."""

from __future__ import annotations

import io
import pathlib

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.draft.run import draft_contract
from drafting_engine.termsheet.checklist import build_checklist
from drafting_engine.termsheet.document import render_termsheet
from drafting_engine.termsheet.extract import sections_from
from drafting_engine.termsheet.record import Fact, Record, verify_facts
from drafting_engine.termsheet.sections import ProposedChecklist
from scripted_llm import PURPOSE, ScriptedLLM

ROOT = pathlib.Path(__file__).resolve().parents[1]
DEMO = ROOT / "examples" / "nda"
PLAYBOOK = ROOT / "playbooks" / "example-nda" / "nda.md"


def _record() -> Record:
    text = render_termsheet(DEMO / "termsheet.docx")
    headings = [line for line in text.split("\n") if line[:3].strip().rstrip(".").isdigit()]
    items, _ = build_checklist(text, headings)
    sections = sections_from(text, ProposedChecklist(items=items))
    purpose = next(s for s in sections if "Purpose" in s.heading)
    quote = "Evaluating a possible joint development of warehouse picking software."
    facts = [Fact(section=purpose.index, term="Purpose", kind="value", value=quote, quote=quote)]
    return Record(deal="demo", model="scripted", sections=sections, facts=verify_facts(sections, facts))


def _xml(docx: bytes) -> str:
    with io.BytesIO(docx) as f:
        import zipfile

        with zipfile.ZipFile(f) as z:
            return z.read("word/document.xml").decode() + z.read("word/comments.xml").decode()


def test_a_term_sheet_value_reaches_the_redline_as_a_tracked_change_with_its_open_point(tmp_path):
    llm = ScriptedLLM()
    drafted = draft_contract(
        _record(),
        (DEMO / "template.docx").read_bytes(),
        PLAYBOOK,
        llm,
        author="Drafting engine",
        workers=2,
        plan_workers=2,
        plan_path=tmp_path / "plan.json",
    )

    assert llm.calls[:3] == ["planner", "instructions", "drafter"]
    assert drafted.failed == []
    xml = _xml(drafted.docx)
    assert "<w:ins " in xml and "<w:del " in xml  # the change is tracked, not overwritten
    assert PURPOSE in xml
    assert "please confirm it" in xml  # the open point is a Word comment
    accepted = Document(io.BytesIO(drafted.docx))
    body = [
        "".join(t.text or "" for t in p.iter(qn("w:t")) if p is not None)
        for p in accepted.element.body.iter(qn("w:p"))
    ]
    assert any("Each party shall keep the other party" in t for t in body)  # untouched text stands
    assert (tmp_path / "plan.json").exists()  # a rerun reuses the plan


def test_a_rerun_reads_the_saved_plan_instead_of_planning_again(tmp_path):
    first, second = ScriptedLLM(), ScriptedLLM()
    for llm in (first, second):
        draft_contract(
            _record(),
            (DEMO / "template.docx").read_bytes(),
            PLAYBOOK,
            llm,
            workers=1,
            plan_workers=1,
            plan_path=tmp_path / "plan.json",
        )
    assert "planner" in first.calls
    assert "planner" not in second.calls
