"""A paragraph the drafter added is marked inserted, never paired with the form's own beside it."""

from __future__ import annotations

import io
import pathlib

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.draft.writer import accept_all_changes, redline, restore_insert_order, unpair_inserts

DEMO = pathlib.Path(__file__).resolve().parents[1] / "examples" / "nda"


def _marks(docx: bytes) -> dict[str, set[str]]:
    """Each paragraph's accepted words, with the kinds of run it holds."""
    out = {}
    for p in Document(io.BytesIO(docx)).element.body.iter(qn("w:p")):
        kinds, words = set(), []
        for r in p.iter(qn("w:r")):
            tracked = [a for a in r.iterancestors() if a.tag in (qn("w:ins"), qn("w:del"))]
            kinds.add(tracked[0].tag.rsplit("}", 1)[1] if tracked else "same")
            if not any(a.tag == qn("w:del") for a in tracked):
                words.append("".join(t.text or "" for t in r.iter(qn("w:t"))))
        out["".join(words)] = kinds
    return out


def test_a_clause_added_beside_an_amended_one_reads_as_added_and_the_amended_one_as_amended():
    base = (DEMO / "template.docx").read_bytes()
    final = accept_all_changes((DEMO / "sample-run" / "draft.docx").read_bytes())
    marks = _marks(unpair_inserts(restore_insert_order(redline(base, final), final), base, final))
    added = next(k for t, k in marks.items() if "prospective funders" in t)
    amended = next(k for t, k in marks.items() if t.startswith("Each party shall protect"))
    assert added == {"ins"}
    assert "same" in amended  # the form's own words stand, only the bracket is marked
