"""A term is trusted only where its quote is in its section; a reader cannot skip a figure."""

from __future__ import annotations

from drafting_engine.termsheet.reader import uncovered_spans
from drafting_engine.termsheet.record import Fact, Section, plain, verify_facts

SECTION = Section(
    index=3,
    heading="6. Duration",
    text="6. Duration\nThe agreement lasts three (3) years. The obligations survive for five (5) years.",
)


def test_a_quote_found_in_its_own_section_is_verified_and_a_retyped_one_is_not():
    facts = [
        Fact(3, "Duration", "value", "three (3) years", "The agreement lasts three (3) years."),
        Fact(3, "Duration", "value", "three years", "The agreement lasts 3 years."),
    ]
    verified = verify_facts([SECTION], facts)
    assert [f.quote_verified for f in verified] == [True, False]


def test_the_renderers_marks_are_not_words():
    assert plain("**Fees** are payable[^fn-2] {#_Ref1}monthly") == "Fees are payable monthly"
    assert plain("signed on this _____ day") == "signed on this _____ day"  # a blank stays


def test_every_sentence_with_a_figure_must_be_quoted_or_skipped():
    quoted = ["The agreement lasts three (3) years."]
    assert uncovered_spans(SECTION.text, quoted) == ["The obligations survive for five (5) years."]
    assert uncovered_spans(SECTION.text, quoted + ["The obligations survive for five (5) years."]) == []
