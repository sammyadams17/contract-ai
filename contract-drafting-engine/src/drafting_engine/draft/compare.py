"""Every difference between the version of a contract the lawyers last saw and a new one, as they
read it: changes accepted, paragraph by paragraph and comment by comment. A read that has this
list has to account for each item; a read without it relies on noticing, which missed lost
sentences and re-worded approved text."""

from __future__ import annotations

import difflib
import io
import pathlib
from dataclasses import dataclass

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.draft.writer import _key, accept_all_changes, clause_numbers
from drafting_engine.word.adeu_engine import render_markdown

_PAIRED = 0.5  # a removed and an added paragraph this alike are one paragraph re-worded


@dataclass(frozen=True)
class Change:
    kind: str  # changed, removed, added, moved, numbering changed, or a comment changed, removed, added
    number: str  # the printed clause number where the paragraph has one
    before: str
    after: str


@dataclass(frozen=True)
class _Unit:
    text: str
    number: str
    comments: tuple[str, ...]


def _units(docx_bytes: bytes) -> list[_Unit]:
    """The contract's paragraphs and table rows in reading order, changes accepted, each with its
    printed number and the comments anchored on it."""
    accepted = accept_all_changes(docx_bytes)
    doc = Document(io.BytesIO(accepted))
    comments: dict[str, str] = {}
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/comments"):
            from lxml import etree

            part = rel.target_part
            root = getattr(part, "_element", None)
            tree = root if root is not None else etree.fromstring(part.blob)
            for c in tree.iter(qn("w:comment")):
                paras = [
                    "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in c.iter(qn("w:p"))
                ]
                comments[c.get(qn("w:id")) or ""] = " ".join(" ".join(paras).split())
    numbers = clause_numbers(accepted, render_markdown(accepted, clean_view=True))
    seen: dict[str, int] = {}
    out: list[_Unit] = []

    def add(text: str, element) -> None:
        text = " ".join(text.split())
        notes = tuple(
            comments[i]
            for i in (r.get(qn("w:id")) for r in element.iter(qn("w:commentReference")))
            if i in comments
        )
        if not text and not notes:
            return
        key = _key(text)
        rank = seen.get(key, 0)
        seen[key] = rank + 1
        out.append(_Unit(text, numbers.get((key, rank), ""), notes))

    for el in doc.element.body:
        if el.tag == qn("w:p"):
            add("".join(t.text or "" for t in el.iter(qn("w:t"))), el)
        elif el.tag == qn("w:tbl"):
            for row in el.iter(qn("w:tr")):
                cells = [
                    " ".join(
                        "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in tc.iter(qn("w:p"))
                    )
                    for tc in row.iter(qn("w:tc"))
                ]
                add(" | ".join(cells), row)
    return out


def _comment_changes(before: _Unit, after: _Unit) -> list[Change]:
    number = after.number or before.number
    gone = [c for c in before.comments if c not in after.comments]
    new = [c for c in after.comments if c not in before.comments]
    out = [Change("comment changed", number, b, a) for b, a in zip(gone, new, strict=False)]
    out += [Change("comment removed", number, b, "") for b in gone[len(new) :]]
    out += [Change("comment added", number, "", a) for a in new[len(gone) :]]
    return out


def compare_versions(before_docx: bytes, after_docx: bytes) -> list[Change]:
    """Every paragraph and comment that differs, in reading order. Numbers that moved with an
    unchanged paragraph are not a change: renumbering follows from changes listed elsewhere."""
    before, after = _units(before_docx), _units(after_docx)
    matcher = difflib.SequenceMatcher(
        None, [_key(u.text) for u in before], [_key(u.text) for u in after], autojunk=False
    )
    out: list[Change] = []
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == "equal":
            for b, a in zip(before[i1:i2], after[j1:j2], strict=True):
                if bool(b.number) != bool(a.number):
                    # the same words, numbered in one version and not the other: a body line
                    # styled as a numbered paragraph, or a numbered paragraph made body text
                    out.append(Change("numbering changed", a.number or b.number, b.text, a.text))
                out += _comment_changes(b, a)
            continue
        olds, news = list(before[i1:i2]), list(after[j1:j2])
        while olds:
            old = olds.pop(0)
            scored = [
                (difflib.SequenceMatcher(None, old.text, n.text, autojunk=False).ratio(), k)
                for k, n in enumerate(news)
            ]
            best = max(scored, default=(0.0, -1))
            if best[0] >= _PAIRED:
                for n in news[: best[1]]:
                    out.append(Change("added", n.number, "", n.text))
                    out += [Change("comment added", n.number, "", c) for c in n.comments]
                new = news[best[1]]
                news = news[best[1] + 1 :]
                out.append(Change("changed", new.number or old.number, old.text, new.text))
                out += _comment_changes(old, new)
            else:
                out.append(Change("removed", old.number, old.text, ""))
                out += [Change("comment removed", old.number, c, "") for c in old.comments]
        for n in news:
            out.append(Change("added", n.number, "", n.text))
            out += [Change("comment added", n.number, "", c) for c in n.comments]
    return _moves(out)


def _moves(changes: list[Change]) -> list[Change]:
    """A paragraph removed in one place and added in another is one paragraph moved: listed
    once, as moved (or moved and changed), at its new number."""
    added = [k for k, c in enumerate(changes) if c.kind == "added"]
    used: set[int] = set()
    out: list[Change] = []
    for c in changes:
        if c.kind != "removed":
            out.append(c)
            continue
        scored = [
            (difflib.SequenceMatcher(None, c.before, changes[k].after, autojunk=False).ratio(), k)
            for k in added
            if k not in used
        ]
        best = max(scored, default=(0.0, -1))
        if best[0] >= 0.8:
            used.add(best[1])
            new = changes[best[1]]
            kind = "moved" if _key(c.before) == _key(new.after) else "moved and changed"
            out.append(Change(kind, new.number or c.number, c.before, new.after))
        else:
            out.append(c)
    kept = [c for c in out if not (c.kind == "added" and _index(changes, c) in used)]
    return _comment_moves(kept)


def _comment_moves(changes: list[Change]) -> list[Change]:
    """A comment removed from one paragraph and added, word for word, on another moved with
    its text: not a change of its own."""
    removed = [c for c in changes if c.kind == "comment removed"]
    added = [c for c in changes if c.kind == "comment added"]
    gone: set[int] = set()
    for r in removed:
        twin = next((a for a in added if id(a) not in gone and a.after == r.before), None)
        if twin is not None:
            gone.update({id(r), id(twin)})
    return [c for c in changes if id(c) not in gone]


def _index(changes: list[Change], change: Change) -> int:
    return next(k for k, c in enumerate(changes) if c is change)


BRIEF_FILE = pathlib.Path(__file__).with_name("briefs") / "compare.md"
_BATCH = 12
