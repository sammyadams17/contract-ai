"""A playbook is a file a lawyer edits: sections by heading, General rules shared by include."""

from __future__ import annotations

import pathlib

from drafting_engine.draft.planner import playbook_sections, rules_for

ROOT = pathlib.Path(__file__).resolve().parents[1]
NDA = ROOT / "playbooks" / "example-nda" / "nda.md"


def test_the_included_general_playbook_and_the_clients_general_rules_are_one_section():
    book = playbook_sections(NDA)
    general = book["General"]
    assert "Every term the term sheet sets is written into the contract" in general  # general.md
    assert "The agreement is mutual unless" in general  # the client's own General (NDA)
    assert "Confidentiality Obligations" in book


def test_a_piece_gets_its_own_section_then_every_section_that_names_it_then_every_general_rule():
    book = playbook_sections(NDA)
    rules = rules_for(book, "Return or Destruction")
    assert rules[0].startswith("- The template's clause stands")
    assert any("The term is the term sheet's" in r for r in rules)  # "Read also by:" reached it
    general = [r for r in rules if r.startswith("Every term the term sheet sets")]
    assert len(general) == 1


def test_include_is_resolved_beside_the_including_file(tmp_path: pathlib.Path):
    (tmp_path / "shared.md").write_text("## General\n\n- Shared rule.\n")
    (tmp_path / "client.md").write_text("@include shared.md\n\n## Fees\n\n- Fee rule.\n")
    book = playbook_sections(tmp_path / "client.md")
    assert book["General"].strip() == "- Shared rule."
    assert book["Fees"].strip() == "- Fee rule."
