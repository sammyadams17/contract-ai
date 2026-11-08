"""The record of a term sheet: its sections, every term they state, and the deal facts given for it.

Pure data and pure functions. Reading the term sheet with a model lives in `extract.py`;
this module owns what a record IS, what makes a fact trustworthy (its quote occurs in its
section), and how a record is frozen and read back. Nothing here calls a model or touches a document.
"""

from __future__ import annotations

import json
import pathlib
import re
from dataclasses import asdict, dataclass, field
from typing import Literal

from drafting_engine.word.span_grounding import normalise_for_comparison

Kind = Literal["value", "position"]

_SCAFFOLDING = [
    (re.compile(r"\[\^fn-\d+\]"), ""),  # footnote references
    (re.compile(r"\[~([^~\]]*)~\]\(#[^)]*\)"), r"\1"),  # cross-reference links, keep the text
    (re.compile(r"\{#[^}]*\}"), ""),  # bookmarks
    (re.compile(r"!\[[^\]]*\]\([^)]*\)"), ""),  # images
    (re.compile(r"<w:br[^>]*/>"), " "),  # page breaks
    (re.compile(r"^#+\s*", re.M), ""),  # heading markers
    (re.compile(r"\*\*|__|(?<![A-Za-z0-9])_|_(?![A-Za-z0-9])"), ""),  # bold and italic marks
]


def plain(markdown: str) -> str:
    """The render with its marks removed: what a reader sees, not how it was drawn.

    A model quotes what it reads and the renderer interleaves `**`, `[^fn-7]` and `{#OLE_LINK}` with
    the words. Most quotes that fail exact search are exact once these are removed.
    """
    # a run of three or more underscores is a blank of the form ("on this _____ day of"), never
    # an emphasis mark: held aside while the marks go (a form full of blanks
    # would otherwise lose every one)
    blanks = re.findall(r"_{3,}", markdown)
    markdown = re.sub(r"_{3,}", "\x00", markdown)
    for pattern, replacement in _SCAFFOLDING:
        markdown = pattern.sub(replacement, markdown)
    for blank in blanks:
        markdown = markdown.replace("\x00", blank, 1)
    return markdown.replace("[●]", "[•]")


def _occurs(quote: str, text: str) -> bool:
    q = normalise_for_comparison(quote)
    return bool(q) and (
        q in normalise_for_comparison(text) or q in normalise_for_comparison(plain(text))
    )


@dataclass(frozen=True)
class Section:
    """One part of the term sheet as it numbers itself, body cut by position."""

    index: int
    heading: str
    text: str
    part: Literal["body", "appendix"] = "body"


@dataclass
class Fact:
    """One agreed point, quoted from the section it was read in."""

    section: int
    term: str
    kind: Kind
    value: str
    quote: str
    quote_verified: bool = False
    note: str = ""
    #: term: a term of the deal; termsheet_only: governs the term sheet itself, not the contract.
    role: str = "term"
    #: The term sheet leaves this to the contract or to later agreement ("to be defined in the contract").
    open: bool = False


@dataclass
class Record:
    deal: str
    model: str
    sections: list[Section]
    facts: list[Fact]
    #: Deal facts a person gives for the whole run, not read from any one section. Each a dict
    #: with `value` and `quote`, and `quote_verified` once checked.
    structure: dict[str, dict] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    #: Spans with a figure or defined term the reader skipped (with its reason) or left uncovered.
    gaps: list[dict] = field(default_factory=list)



def verify_facts(sections: list[Section], facts: list[Fact]) -> list[Fact]:
    """Mark each fact by whether its quote occurs in its own section. Nothing is dropped."""
    by_index = {s.index: s.text for s in sections}
    for fact in facts:
        fact.quote_verified = _occurs(fact.quote, by_index.get(fact.section, ""))
    return facts


def verify_structure(termsheet_text: str, structure: dict[str, dict]) -> dict[str, dict]:
    for entry in structure.values():
        entry["quote_verified"] = _occurs(str(entry.get("quote", "")), termsheet_text)
    return structure


def save_record(record: Record, path: pathlib.Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(record), indent=1, ensure_ascii=False))


def load_record(path: pathlib.Path) -> Record:
    raw = json.loads(path.read_text())
    return Record(
        deal=raw["deal"],
        model=raw["model"],
        sections=[Section(**s) for s in raw["sections"]],
        facts=[Fact(**f) for f in raw["facts"]],
        structure=raw.get("structure", {}),
        notes=raw.get("notes", []),
        gaps=raw.get("gaps", []),
    )


def dedupe(facts: list[Fact]) -> list[Fact]:
    """One fact per (section, term, value, quote), case and whitespace aside. The reader may record
    the same term twice when it follows a reference back into its own section."""
    seen: set[tuple] = set()
    kept: list[Fact] = []
    for f in facts:
        key = (f.section, *(normalise_for_comparison(x) for x in (f.term, f.value, f.quote)))
        if key not in seen:
            seen.add(key)
            kept.append(f)
    return kept
