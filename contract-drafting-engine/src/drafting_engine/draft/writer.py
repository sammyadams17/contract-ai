"""Writing the draft into Word: plain edits, then a comparison for the redline.

The drafter ends with each paragraph of a clause as it should finally read. This module puts
those paragraphs into a plain copy of the form (no tracked changes), then a document-comparison
engine marks the copy up against the form. That is how Word's own Compare works, and how the
legal-tech tools do it: the redline comes from comparing two finished documents, not from
hand-writing each change. Tables, footnote references and cross-reference fields are the
engine's problem, and it handles them.

Paragraphs are found by alignment, not by text search: the form's clean render and the Word file
list the same paragraphs in the same order, so the first still-unmatched paragraph reading as a
given original is that paragraph. Cross-reference fields inside a changed paragraph are kept live
where the visible text survives in the new wording.
"""

from __future__ import annotations

import copy
import datetime
import difflib
import io
import pathlib
import re
from itertools import takewhile

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.draft.paragraphs import _FOOTNOTE, PARA_BREAK, document_text
from drafting_engine.word.adeu_engine import MATH, math_text, outer_equations
from drafting_engine.word.numbering import ClauseNumber, accepted_numbers_by_paragraph
from drafting_engine.word.span_grounding import normalise_for_comparison

# A vertical tab or form feed is a line break in disguise. The rest are control
# characters lxml will not put in a w:t — one of them used to abort the whole
# draft after every clause had already been written.
_LINE_BREAK_CONTROLS = re.compile("[\x0b\x0c]")
_XML_ILLEGAL = re.compile("[\x00-\x08\x0e-\x1f\ufffe\uffff]")


def _xml_safe(text: str) -> str:
    return _XML_ILLEGAL.sub("", _LINE_BREAK_CONTROLS.sub(" ", text))


def snapshot(
    work, title: str = "", whole_document: bool = False, before_text: str | None = None
) -> dict:
    """Serialise a Rewriting's final state for the writer and the record: what each paragraph
    finally reads, what was removed, what was added, and the reviewer's note for each. With
    `before_text`, the render of everything in the document before this unit, every changed
    and deleted paragraph carries its rank among paragraphs reading the same in the whole
    document, so the writer changes that one and never an identical paragraph elsewhere."""
    changed, inserted, rows, tables, notes = [], [], [], [], []
    order: list[tuple[str, int]] = []  # insert groups, rows and tables in document order
    insert_ranks: list[int] = []  # which occurrence of its anchor text each group follows
    row_ranks: list[int] = []
    table_ranks: list[int] = []
    changed_paras: list = []
    previous_original = ""
    previous_para = None
    offsets: dict[str, int] = {}
    if before_text is not None:
        from drafting_engine.draft.paragraphs import document_text

        for line in before_text.split("\n"):
            k = _key(document_text(line))
            if k:
                offsets[k] = offsets.get(k, 0) + 1
                # an addition before the unit's first paragraph is anchored to the last
                # paragraph before the unit, never to an empty anchor that matches anywhere
                previous_original = document_text(line)
    everything = [*work.paras, *getattr(work, "deleted", [])]
    # the paragraphs this piece added, by their words, where no form paragraph reads so
    form_words = {x.original for x in everything if x.inserted is None}
    added = {
        x.text: x for x in work.paras if x.inserted is not None and x.text not in form_words
    }
    by_id = {x.id: x for x in work.paras}

    def whole_rank(para) -> int:
        if _number(para.id) is None and para.inserted is None:
            return -1  # a row the render leaves out: no count of the render's lines names it
        r = _rank(everything, para)
        return -1 if r < 0 else r + offsets.get(_key(para.original), 0)

    def anchor_rank() -> int:
        """Which occurrence, in the whole document, of the anchor text the next addition
        follows: the paragraph before it in this unit, or the last line before the unit. A
        heading reading "Term" is told from a table cell reading "Term" by its place."""
        if before_text is None:
            return -1
        if previous_para is not None:
            return whole_rank(previous_para)
        return offsets.get(_key(previous_original), 0) - 1

    numbered = sorted(
        (x for x in everything if _number(x.id) is not None), key=lambda x: _number(x.id) or 0
    )

    def context(para) -> list:
        """The paragraph before it reading differently, and how many reading as it does come
        between: an empty row is told from the other empty rows by what it follows, not by a
        count from the top of a document whose other tables the render may read otherwise."""
        number, want = _number(para.id), _key(para.original)
        if number is None and para.inserted is None and para in work.paras:
            # a row of empty cells the view lists with an id of its own ("e1"): by the form
            # paragraphs before it where it stands
            k = 0
            for q in reversed(work.paras[: work.paras.index(para)]):
                if q.inserted is not None:
                    continue
                if _key(q.original) == want:
                    k += 1
                elif _key(q.original):
                    return [q.original, k]
            return ["", -1]
        if number is None:
            return ["", -1]
        k = 0
        for q in reversed(numbered):
            n = _number(q.id) or 0
            if n >= number:
                continue
            if _key(q.original) == want:
                k += 1
            elif _key(q.original):
                return [q.original, k]
        return ["", -1]

    anchor_context: dict[str, list] = {"p": [], "r": [], "t": []}

    def anchor_ctx() -> list:
        """Where the paragraph an addition follows sits, as context() tells it."""
        return context(previous_para) if previous_para is not None else ["", -1]

    why, substantive = _why, _substantive

    # each table this piece added, by its edit: its rows as the agent's state lists them, rows
    # it inserted between them included, so they are written where the agent put them
    table_rows: dict[int, list] = {}
    for ei in {x.inserted[0] for x in work.paras if x.inserted is not None}:
        if work.edits[ei].get("kind") != "table":
            continue
        at = [k for k, x in enumerate(work.paras) if x.inserted is not None and x.inserted[0] == ei]
        table_rows[ei] = [
            x
            for x in work.paras[at[0] : at[-1] + 1]
            if x.inserted is not None
            and (x.inserted[0] == ei or work.edits[x.inserted[0]].get("kind") == "insert_row")
        ]
    in_table = {id(x) for rows_of in table_rows.values() for x in rows_of}

    def row_cells(q) -> list[str]:
        """An added row's cells as it now reads: a replace after the insert corrects them."""
        made = work.edits[q.inserted[0]]
        if made.get("kind") == "insert_row":
            return made["cells"] if q.text == made["new_text"] else q.text.split(" | ")
        if q.cells is not None and q.text == " | ".join(q.cells):
            return q.cells
        return q.text.split(" | ")

    for p in work.paras:
        if p.inserted is not None:
            ei = p.inserted[0]
            edit = work.edits[ei]
            if edit.get("kind") == "insert_row" and id(p) in in_table:
                # a row of a table this piece added: written with that table
                cells = row_cells(p)
                own = [e for e in work.edits if e.get("para") == p.id]
                notes.append((" | ".join(cells), why([edit, *own])))
            elif edit.get("kind") == "insert_row":
                cells = row_cells(p)
                rows.append((previous_original, cells))
                row_ranks.append(anchor_rank())
                anchor_context["r"].append(anchor_ctx())
                order.append(("r", len(rows) - 1))
                own = [e for e in work.edits if e.get("para") == p.id]
                notes.append((" | ".join(cells), why([edit, *own])))
            elif edit.get("kind") == "table":
                # written at its first row still standing: a table whose header the drafter
                # struck is the rows left (a whole account-details table can go out)
                if p is table_rows[ei][0]:
                    # its rows as they now read, each replaced row as corrected
                    now = table_rows[ei]
                    grid = [row_cells(q) for q in now]
                    tables.append((previous_original, grid[0], grid[1:]))
                    table_ranks.append(anchor_rank())
                    anchor_context["t"].append(anchor_ctx())
                    order.append(("t", len(tables) - 1))
                    notes.append((" | ".join(grid[0]), why([edit])))
            elif (
                not inserted
                or inserted[-1][0] != previous_original
                or inserted[-1][2] != ei
                # a table or row placed inside the group: what follows it is a group of its
                # own, written after that table, never joined to the paragraphs before it
                or order[-1] != ("p", len(inserted) - 1)
            ):
                # the insert's reason and quote are noted once, on its first paragraph
                later = any(x[2] == ei for x in inserted)
                inserted.append((previous_original, [p.text], ei, [_heading_flag(p, added, by_id)]))
                insert_ranks.append(anchor_rank())
                anchor_context["p"].append(anchor_ctx())
                order.append(("p", len(inserted) - 1))
                # a paragraph the piece added and then replaced or re-noted: the last reason
                own = substantive(
                    [e for e in work.edits if e.get("para") == p.id], _inserted_text(edit, p)
                )
                if not later:
                    notes.append((p.text, why([edit, *own])))
                elif own:
                    notes.append((p.text, why(own)))
            else:
                inserted[-1][1].append(p.text)
                inserted[-1][3].append(_heading_flag(p, added, by_id))
                # a later paragraph of the insert, replaced or re-styled since: that edit's own
                # reason, with its own quote; the insert's quote is about the insert
                own = substantive(
                    [e for e in work.edits if e.get("para") == p.id], _inserted_text(edit, p)
                )
                if own:
                    notes.append((p.text, why(own)))
            continue
        if p.text != p.original or p.like:
            changed.append((p.original, p.text))
            changed_paras.append(p)
            own = [e for e in work.edits if e.get("para") == p.id or e.get("cell") == p.id]
            notes.append((p.text, why(substantive(own, p.original))))
        previous_original = p.original
        previous_para = p
    # the form paragraph whose style and numbering each changed paragraph takes, where the
    # drafter named one; a snapshot that restyles nothing reads as it always did
    restyled = {"changed_like": [x.like for x in changed_paras]} if any(
        x.like for x in changed_paras
    ) else {}
    # the form paragraphs whose following pictures and empty paragraphs the drafter struck,
    # each found as an addition's anchor is; a snapshot that strikes none reads as it did
    struck_empty = []
    for e in work.edits:
        after = next((x for x in everything if x.id == e.get("after")), None)
        if e.get("kind") != "strike_empty" or after is None:
            continue
        struck_empty.append([after.original, whole_rank(after) if before_text is not None else -1, context(after)])
        if e.get("reason"):
            notes.append((after.text, why([e])))
    empties = {"struck_empty": struck_empty} if struck_empty else {}
    # the removed paragraphs in document order, as the changed ones are, not the order the
    # drafter removed them in: each is found after the paragraph before it
    struck = sorted(getattr(work, "deleted", []), key=lambda x: _number(x.id) or 0)
    # a strike carries its reason, as a change does: once for a run of paragraphs struck
    # together for one reason, and none where the paragraph was moved (its comment goes with it)
    moved_away = {getattr(x, "moved_from", None) for x in work.paras}
    place = {x.id: k for k, x in enumerate(numbered)}
    previous: tuple[int, str] | None = None
    for x in struck:
        if x.id in moved_away or not x.original.strip():
            continue
        comment = why([e for e in work.edits if e.get("kind") == "delete" and e.get("para") == x.id])
        at = place.get(x.id, -2)
        if comment and previous != (at - 1, comment):
            notes.append((x.original, comment))
        previous = (at, comment)
    return {
        "changed": changed,
        **restyled,
        **empties,
        "deleted": [x.original for x in struck],
        # which occurrence of that text each deletion is, counting the render's paragraphs in
        # order: identical placeholder lines are told apart by rank
        # ranks count the whole document's paragraphs, so only a snapshot of the whole document
        # can name them; a unit's snapshot is placed in order instead
        "deleted_ranks": (
            [whole_rank(x) for x in struck]
            if whole_document or before_text is not None
            else []
        ),
        "changed_ranks": (
            [whole_rank(x) for x in changed_paras] if before_text is not None else []
        ),
        "changed_context": [context(x) for x in changed_paras],
        "deleted_context": [context(x) for x in struck],
        "inserted": [(a, b, h) for a, b, _, h in inserted],
        "rows_inserted": rows,
        "tables_inserted": tables,
        # the order the agent made them in: several additions after one form paragraph land
        # one after another, each after the last, never each directly after the anchor
        "insert_order": order,
        # which occurrence of the anchor text each addition follows, whole-document, so an
        # anchor whose words appear elsewhere (a table cell reading "Term") is not mistaken
        "insert_ranks": insert_ranks,
        "row_ranks": row_ranks,
        "table_ranks": table_ranks,
        # what each addition's anchor follows, so an anchor reading like others (an empty row)
        # is found by its place, not by a count the render and the file may disagree on
        "anchor_context": anchor_context,
        "notes": notes,
        # where the piece's open points are noted: its first paragraph as it now reads
        "first": next((p.text for p in work.paras if p.text.strip()), ""),
    }


def _comment(edits: list[dict]) -> tuple[str, str]:
    """The reason and term sheet quote of the comment the lawyer reads: the last reason given for
    the paragraph (a fix supersedes the drafting), with the last quote behind it; none where no
    reason is left (the drafter removed it), quote or not."""
    reason = next((e["reason"] for e in reversed(edits) if e.get("reason")), "")
    if not reason:
        return "", ""
    return reason, next((e["termsheet_quote"] for e in reversed(edits) if e.get("termsheet_quote")), "")


def _why(edits: list[dict]) -> str:
    """The comment the lawyer reads, as written: the reason, then the term sheet's words."""
    reason, quote = _comment(edits)
    return reason + (f' | term sheet: "{quote[:240]}"' if quote else "")


def _substantive(edits: list[dict], start: str) -> list[dict]:
    """The edits that changed the words: one writing them back as they stood, or changing
    only their punctuation or spacing, gives no reason to read (a restyle still does)."""
    out, before = [], start
    for e in edits:
        new = e.get("new_text")
        if e.get("kind") == "modify" and new is not None:
            same = _letters(new) == _letters(before)
            before = new
            if same and not e.get("like"):
                continue
        out.append(e)
    return out


def shown_comment(work, p) -> tuple[str, str]:
    """The comment the snapshot gives paragraph `p` of `work`, as (reason, term sheet quote);
    ("", "") for none."""
    own = [e for e in work.edits if e.get("para") == p.id or e.get("cell") == p.id]
    if p.inserted is None:
        if p.text == p.original and not p.like:
            return "", ""
        return _comment(_substantive(own, p.original))
    edit = work.edits[p.inserted[0]]
    if edit.get("kind") == "insert_row":
        return _comment([edit, *[e for e in work.edits if e.get("para") == p.id]])
    if edit.get("kind") == "table":
        first = next(x for x in work.paras if x.inserted is not None and x.inserted[0] == p.inserted[0])
        return _comment([edit]) if p is first else ("", "")
    later = _substantive([e for e in work.edits if e.get("para") == p.id], _inserted_text(edit, p))
    return _comment([edit, *later] if p.inserted[1] == 0 else later)


def piece_notes(result: dict, open_points_only: bool = False) -> list[tuple[str, str]]:
    """The comments one drafted piece puts in the contract: its open points first, at its first
    paragraph, then a comment on each change. With `open_points_only`, only the open points
    are returned."""
    paragraphs = result.get("paragraphs", {})
    # a piece saved by an earlier run wrote "<its title>: <reason> || <reason> ..."; the lawyer
    # reads the last reason only
    named = f"{result.get('title', '')}: "
    notes = [
        (anchor, text.removeprefix(named).rsplit(" || ", 1)[-1] if text.startswith(named) else text)
        for anchor, text in paragraphs.get("notes", [])
    ]
    # a piece saved without its first paragraph, and with no change, is known by its title
    first = paragraphs.get("first") or (notes[0][0] if notes else "") or result.get("title", "")
    paras = (result.get("state") or {}).get("paras") or []
    points = [
        (
            _point_anchor(paras, o["at_id"]) or o.get("at") or first
            if o.get("at_id")
            else o.get("at") or first,
            f"Open point: {o['summary'].removeprefix('Open point:').strip()}",
        )
        for o in result.get("open_points", [])
        if o.get("summary")
    ]
    return [*points] if open_points_only else [*points, *notes]  # type: ignore[list-item]


class Anchor(str):
    """A comment's anchor words that also know where in their piece they stand: `after` is the
    paragraph before them reading otherwise, and how many reading as they do come between. A
    "[•]" is told from every other "[•]" of the contract by what it follows."""

    after: tuple[str, int] | None = None


def _point_anchor(paras: list[dict], pid: str) -> Anchor | None:
    """The words of the paragraph `pid` of a piece's saved state as they now read, with their
    place in the piece; None where the piece no longer holds it."""
    at = next((i for i, d in enumerate(paras) if d.get("id") == pid), None)
    if at is None:
        return None
    anchor = Anchor(paras[at]["text"])
    want, k = _key(anchor), 0
    for d in reversed(paras[:at]):
        if _key(d["text"]) == want:
            k += 1
        elif _key(d["text"]):
            anchor.after = (d["text"], k)
            break
    return anchor


def _letters(text: str) -> str:
    """A paragraph's words without spacing, footnote markers and sentence punctuation: a
    bracket, a blank, a capital or a figure's separator ("12,500") is a change of words, a
    comma or a semicolon between words is not."""
    text = re.sub(r"\s+", "", _FOOTNOTE.sub("", text))
    return re.sub(r"(?<!\d)[,;:.]|[,;:.](?!\d)", "", text)


def _spaced(text: str) -> str:
    """A paragraph's words as written, footnote markers and spacing aside."""
    return " ".join(_FOOTNOTE.sub("", text).split())


def _inserted_text(edit: dict, p) -> str:
    """The words a paragraph was inserted with, before any later edit of it."""
    texts = edit.get("paragraphs") or [edit.get("new_text", "")]
    k = p.inserted[1] if p.inserted is not None else 0
    return texts[k] if k < len(texts) else ""


def _heading_flag(p, added: dict | None = None, by_id: dict | None = None) -> str | bool | dict:
    """How a new paragraph is styled: False for a body paragraph, else the text of the form
    heading whose style the new heading takes (True where none is known); {"like": text} for
    one the drafter restyled as the form paragraph reading so; "beneath" for a limb of the
    form paragraph it follows, {"beneath": text} for a limb of a paragraph the piece added
    (found in `by_id`), reading so. A heading named after a paragraph the piece added itself
    (`added`, by its words) takes the form paragraph that one was styled as: the form has no
    paragraph reading as the added one."""
    if getattr(p, "like", ""):
        return {"like": p.like, **({"struck": True} if getattr(p, "like_struck", False) else {})}
    if not p.heading:
        if not getattr(p, "beneath", False):
            return False
        parent = (by_id or {}).get(getattr(p, "beneath_of", ""))
        return {"beneath": parent.text} if parent is not None else "beneath"
    name, seen = p.heading_like, set()
    while added and name in added and name not in seen:
        seen.add(name)
        q = added[name]
        name = getattr(q, "like", "") or q.heading_like or name
    return name or True


def _key(text: str) -> str:
    """Text as the render and the Word file agree on it: footnote references, rendering marks,
    whitespace, underscores in blanks, literal asterisks the render reads as bold marks, bullet
    glyph and dash style all set aside. A break inside a cell too: the render shows a line
    break as a paragraph break and drops a paragraph of only a blank's underscores."""
    return re.sub(r"\s+", "", _shown(text))


def _shown(text: str) -> str:
    """Text as _key reads it, its spacing kept: each run of whitespace one space."""
    text = _FOOTNOTE.sub("", text).replace(PARA_BREAK.strip(), "")
    text = re.sub(r"\[Com:\d+\]", "", text)
    text = re.sub(r"\[~([^~\]]*)~\]\(#[^)]*\)", r"\1", text)
    text = text.replace("●", "•").replace("\u2013", "-").replace("\u2014", "-").replace("_", "")
    text = text.replace("*", "").replace("\u200b", "")
    text = re.sub(r"^\s*\d+\.\s+", "", text)  # a literal "1. " the render shows as a list marker
    return normalise_for_comparison(text)


class _Body:
    """The document's paragraphs and table rows in document order, each with the text the
    render shows for it, matched to originals in order."""

    def __init__(self, doc) -> None:
        self.doc = doc
        self.items: list[tuple[str, object]] = []  # (key, w:p or w:tr element)
        parts = [doc.element.body]
        for section in doc.sections:
            for part in (
                section.header,
                section.first_page_header,
                section.footer,
                section.first_page_footer,
            ):
                if not part.is_linked_to_previous:
                    parts.append(part._element)
        self.inline: dict[int, str] = {}  # item index -> the line, for a line inside a paragraph
        for root in parts:
            for el in root.iter(qn("w:p"), qn("w:tr")):
                if any(
                    a.tag == qn("w:sdt") and a.getparent().tag in (qn("w:tr"), qn("w:tbl"))
                    for a in el.iterancestors()
                ):
                    # the render lists nothing inside a content control wrapping a table cell
                    # or row (the cover blocks), so the drafter never sees it and it is not
                    # counted here either; a block-level control's paragraphs are rendered
                    continue
                if el.tag == qn("w:tr"):
                    self.items.append((_key(_row_text(el)), el))
                else:
                    # every paragraph, wherever it sits (a cell, a content control, a text box,
                    # the body), is its own item; a row is an item too, so a row-level change
                    # and a paragraph-level change can each find their element
                    # a tab read as the render shows it, a space: "6.<tab>Payment" is the
                    # contents line the view shows as "6. Payment"
                    self.items.append((_key(_element_text(el, tabs=True)), el))
                    # a paragraph with line breaks renders as several lines, each a paragraph
                    # to the drafter: each line is an item too, at its place in the order
                    lines = _lines_of(el)
                    if len(lines) > 1:
                        for line in lines:
                            self.inline[len(self.items)] = line
                            self.items.append((_key(line), el))
        self.used: set[int] = set()

    def _counts(self, i: int) -> bool:
        """Whether an item is one the render lists on its own. A row of several cells is one
        line of the render, so the row is the item and its cells' paragraphs are not. A table of
        one column renders each cell paragraph as its own line, so there the paragraphs are the
        items and the row is not."""
        el = self.items[i][1]
        if el.tag == qn("w:tr"):
            return len(el.findall(qn("w:tc"))) > 1
        parent = el.getparent()
        if el.tag == qn("w:p") and parent is not None and parent.tag == qn("w:tc"):
            row = parent.getparent()
            return row is not None and len(row.findall(qn("w:tc"))) == 1
        return True

    def take_nth(self, original: str, rank: int) -> tuple[int, object] | None:
        """The rank-th item (0-based, in document order) whose text is `original`, counting as
        the render counts, whether or not earlier ones were taken."""
        want = _key(original)
        if not want:
            return None
        seen = 0
        for i, (k, el) in enumerate(self.items):
            if k != want or not self._counts(i):
                continue
            if seen == rank:
                if i in self.used:
                    return None
                self.used.add(i)
                return i, el
            seen += 1
        return None

    def take_following(self, original: str, before: str, k: int) -> tuple[int, object] | None:
        """The item with this text that is the k-th (0-based) of its run directly after the one
        place reading `before`; None unless exactly one place fits. For text that repeats, an
        empty row above all, where a count from the top of the document drifts."""
        i = self.following(original, before, k)
        if i is None or i in self.used:
            return None
        self.used.add(i)
        return i, self.items[i][1]

    def following(self, original: str, before: str, k: int) -> int | None:
        """take_following's place, taken or not."""
        want, prior = _key(original), _key(before)
        if not want or not prior or k < 0 or sum(key == want for key, _ in self.items) < 2:
            return None
        fits = []
        for i, (key, el) in enumerate(self.items):
            if key != prior or not self._counts(i):
                continue
            seen = 0
            for j in range(i + 1, len(self.items)):
                kj, ej = self.items[j]
                if ej is el or not self._counts(j) or not kj:
                    continue  # the anchor's own lines, a row's cell paragraphs, blank lines
                if kj != want:
                    break
                if seen == k:
                    fits.append(j)
                    break
                seen += 1
        return fits[0] if len(fits) == 1 else None

    def in_place(self, original: str, before: str, k: int, rank: int = -1) -> int | None:
        """The paragraph of the body reading `original` that is the k-th (0-based) of its run
        after one reading `before`, counting every paragraph a table cell holds as well as the
        body's: the view shows a cell's paragraphs as lines of their own. Where several fit,
        the rank-th paragraph reading so, if it is one of them; None otherwise."""
        want, prior = _key(original), _key(before)
        if not want or not prior or k < 0:
            return None
        paras = [
            i
            for i, (_, el) in enumerate(self.items)
            if el.tag == qn("w:p") and i not in self.inline and el.getroottree().getroot() is self.doc.element
        ]
        reading = [i for i in paras if self.items[i][0] == want]

        def fits(at: int) -> bool:
            seen = 0
            for i in reversed([x for x in paras if x < at]):
                key = self.items[i][0]
                if key == want:
                    seen += 1
                elif key:
                    return key == prior and seen == k
            return False

        ok = [i for i in reading if fits(i)]
        if len(ok) == 1:
            return ok[0]
        if 0 <= rank < len(reading) and reading[rank] in ok:
            return reading[rank]
        return None

    def take(self, original: str, after: int = 0) -> tuple[int, object] | None:
        """The first untaken item from `after` with this text, preferring items the render lists
        on its own; a paragraph inside a cell is taken only when nothing else matches."""
        want = _key(original)
        if not want:
            return None
        for cells_too in (False, True):
            for i in range(after, len(self.items)):
                if i in self.used or self.items[i][0] != want:
                    continue
                if not cells_too and not self._counts(i):
                    continue
                self.used.add(i)
                return i, self.items[i][1]
        return None


W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def _fresh_ids(element) -> None:
    """Word requires every paragraph's w14:paraId and w14:textId to be unique; a copied
    paragraph or row must not carry its source's."""
    import secrets

    for p in [element, *element.iter(qn("w:p"))]:
        if p.tag != qn("w:p"):
            continue
        for attr in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
            if attr in p.attrib:
                p.set(attr, f"{secrets.randbelow(0x7FFFFFFF):08X}")


def _rank(paras, para) -> int:
    """Which occurrence of its text a paragraph is, counting the render's paragraphs in order.
    The render numbers its paragraphs p1, p2, ... in document order, so the rank is the count
    of paragraphs of the same text with a lower number, deleted ones included; -1 for an
    inserted paragraph, which has no place in the render."""
    number = _number(para.id)
    if number is None:
        return -1
    want = _key(para.original)
    return sum(
        1
        for q in paras
        if _key(q.original) == want and (n := _number(q.id)) is not None and n < number
    )


def _nth(items: list | None, n: int):
    return items[n] if items and n < len(items) else None


def _number(pid: str) -> int | None:
    m = re.fullmatch(r"p(\d+)", pid or "")
    return int(m.group(1)) if m else None


def _element_text(p, tabs: bool = False) -> str:
    """The element's text as the render shows it: a page-number field's cached result is not
    text of the document (the render drops it), so it is not text here either, whether the
    field is simple or complex. With `tabs`, a tab is a space, as the render shows it."""
    out = []
    in_page_field = False
    for el in p.iter(qn("w:t"), qn("w:fldChar"), qn("w:instrText"), qn("w:tab"), MATH + "oMath"):
        if el.tag == MATH + "oMath":
            # an equation reads as the render shows it; one inside another is read with it
            if not any(True for _ in el.iterancestors(MATH + "oMath")):
                out.append(math_text(el))
            continue
        if el.tag == qn("w:tab"):
            if tabs and not in_page_field and el.getparent().tag == qn("w:r"):
                out.append(" ")
            continue
        if el.tag == qn("w:fldChar"):
            kind = el.get(qn("w:fldCharType"))
            if kind == "end":
                in_page_field = False
            continue
        if el.tag == qn("w:instrText"):
            if _is_page_field(el.text or ""):
                in_page_field = True
            continue
        if in_page_field:
            continue
        fld = next((a for a in el.iterancestors(qn("w:fldSimple"))), None)
        if fld is not None and _is_page_field(fld.get(qn("w:instr")) or ""):
            continue
        out.append(el.text or "")
    return "".join(out)


def _is_page_field(instruction: str) -> bool:
    """A field showing a page count or the page it sits on (PAGE, NUMPAGES, SECTIONPAGES), by
    its field type, the code's first word. A PAGEREF shows the page of a bookmark, a number
    the render keeps."""
    kind = instruction.split()[:1]
    return bool(kind) and kind[0].upper() in ("PAGE", "NUMPAGES", "SECTIONPAGES")


def _lines_of(p) -> list[str]:
    """A paragraph's text split at its line breaks, as the render shows it."""
    lines, current = [], ""
    for el in p.iter(qn("w:t"), qn("w:br")):
        if el.tag == qn("w:br"):
            lines.append(current)
            current = ""
        else:
            current += el.text or ""
    lines.append(current)
    return [ln for ln in (x.strip() for x in lines) if ln]


def _set_line(el, line: str, final: str, original: str = "") -> None:
    """One line of a multi-line paragraph rewritten (or removed, when final is empty)."""
    text = _element_text(el)
    plain_line = " ".join(_FOOTNOTE.sub("", line).split())
    new = " ".join(_FOOTNOTE.sub("", final).split())
    if plain_line in text:
        text = text.replace(plain_line, new, 1)
    else:
        pattern = r"\s*".join(re.escape(w) for w in plain_line.split())
        text = re.sub(pattern, lambda m: new, text, count=1)
    set_paragraph_text(el, text, original)


def _row_text(tr) -> str:
    cells = []
    for tc in tr.findall(qn("w:tc")):
        paras = [_element_text(p) for p in tc.findall(qn("w:p"))]
        cells.append(PARA_BREAK.join(x for x in paras if x.strip()) or "")
    return " | ".join(cells)


def _fields_of(p) -> list[tuple[list, str]]:
    """Each cross-reference field of a paragraph as (its runs, its visible text)."""
    fields, current, visible, inside = [], [], "", False
    for r in p.findall(qn("w:r")):
        fld = r.find(qn("w:fldChar"))
        kind = fld.get(qn("w:fldCharType")) if fld is not None else None
        if kind == "begin":
            inside, current, visible = True, [r], ""
            continue
        if inside:
            current.append(r)
            if r.find(qn("w:instrText")) is None:
                visible += _element_text(r)
            if kind == "end":
                fields.append((current, _INVISIBLE.sub("", visible)))
                inside = False
    return fields


# marks that print nothing (a left-to-right mark Word puts at the head of a field's result):
# the drafter's text never carries them, so they are not words to match a field by
_INVISIBLE = re.compile("[\u200b-\u200f\u2060\ufeff]")


def _target_title(p, field_runs) -> str:
    """The text of the paragraph a REF field's bookmark opens, "" for any other field or a
    bookmark the document does not hold."""
    instruction = "".join(it.text or "" for r in field_runs for it in r.iter(qn("w:instrText")))
    m = _REF_FIELD.search(instruction)
    if not m:
        return ""
    mark = next(
        (
            b
            for b in p.getroottree().getroot().iter(qn("w:bookmarkStart"))
            if b.get(qn("w:name")) == m.group(1)
        ),
        None,
    )
    if mark is None:
        return ""
    host = next(mark.iterancestors(qn("w:p")), None)
    if host is None:
        host = mark.getnext()  # a bookmark opened just before its paragraph
    return _element_text(host) if host is not None and host.tag == qn("w:p") else ""


def _names_its_title(title: str, old_after: str, new_after: str) -> bool:
    """Whether a reference still names the heading its field opens. Where the form follows the
    reference with that heading's title, with nothing but spaces and marks between ("Schedule
    3 (Service Levels)"), the new words must too; elsewhere the
    form gives nothing to go by and the field stays."""

    gap = _title_gap(title, old_after)
    return gap is None or _lower_words(new_after).startswith(gap + _lower_words(title))


def _lower_words(x: str) -> str:
    return " ".join(x.lower().split())


def _title_gap(title: str, after: str) -> str | None:
    """What stands between a reference and the title of the heading it opens, where the form
    follows the reference with that title with nothing but spaces and marks between; None
    where it does not, and the words after the reference say nothing of where it points."""
    title, after = _lower_words(title), _lower_words(after)
    if not title or title not in after:
        return None
    gap = after[: after.index(title)]
    return None if any(c.isalnum() for c in gap) else gap


# a footnote marker, a run of letters, of digits or of spaces, or any other one character
_DIFF_TOKEN = re.compile(r"\[\^fn-\d+\]|[^\W\d_]+|\d+|\s+|.")


def _fewest_edits(a: list[str], b: list[str]) -> list[tuple[str, int, int, int, int]]:
    """The changes from `a` to `b` in the fewest characters put in, taken out or replaced, a
    token at a time, as difflib's opcodes. Not difflib's longest match: "(k), (l) or (m)"
    against "(j), (k) or (l)" is three replaced letters, never "(l)" matched across to where
    "(m)" stood."""
    head = 0
    while head < min(len(a), len(b)) and a[head] == b[head]:
        head += 1
    tail = 0
    while tail < min(len(a), len(b)) - head and a[-1 - tail] == b[-1 - tail]:
        tail += 1
    x, y = a[head : len(a) - tail], b[head : len(b) - tail]
    def swap(u: str, v: str) -> int:
        return 0 if u == v else max(len(u), len(v))

    cost = [[0] * (len(y) + 1)]
    for j in range(1, len(y) + 1):
        cost[0][j] = cost[0][j - 1] + len(y[j - 1])
    for i in range(1, len(x) + 1):
        row = [cost[i - 1][0] + len(x[i - 1])]
        for j in range(1, len(y) + 1):
            row.append(
                min(
                    cost[i - 1][j - 1] + swap(x[i - 1], y[j - 1]),
                    cost[i - 1][j] + len(x[i - 1]),
                    row[j - 1] + len(y[j - 1]),
                )
            )
        cost.append(row)
    moves: list[tuple[bool, int, int]] = []  # from the end: (unchanged, old taken, new taken)
    i, j = len(x), len(y)
    while i or j:
        if i and j and cost[i][j] == cost[i - 1][j - 1] + swap(x[i - 1], y[j - 1]):
            moves.append((x[i - 1] == y[j - 1], 1, 1))
            i, j = i - 1, j - 1
        elif i and cost[i][j] == cost[i - 1][j] + len(x[i - 1]):
            moves.append((False, 1, 0))
            i -= 1
        else:
            moves.append((False, 0, 1))
            j -= 1
    groups: list[list] = [[True, 0, head, 0, head]] if head else []
    i, j = head, head
    for same, di, dj in reversed(moves):
        if groups and groups[-1][0] == same:
            groups[-1][2] += di
            groups[-1][4] += dj
        else:
            groups.append([same, i, i + di, j, j + dj])
        i, j = i + di, j + dj
    if tail:
        groups.append([True, len(a) - tail, len(a), len(b) - tail, len(b)])
    ops = []
    for same, i1, i2, j1, j2 in groups:
        op = "equal" if same else "replace" if i2 > i1 and j2 > j1 else "delete" if i2 > i1 else "insert"
        ops.append((op, i1, i2, j1, j2))
    return ops


def _form_characters(
    p, tabs: dict[int, object] | None = None
) -> tuple[str, list, list[tuple[int, object, object]]]:
    """A paragraph's words as one string, spacing as the drafter's text has it, with the run
    each character is in, and each line break with the place in that string it stands. A tab
    is a space, as the render shows it; `tabs`, if given, gets the run of each such space."""
    chars: list[str] = []
    owners: list = []
    breaks: list[tuple[int, object, object]] = []
    for r in p.iter(qn("w:r")):
        for part in r:
            if part.tag == qn("w:br"):
                breaks.append((len(chars), r, part))
            elif part.tag == qn("w:tab") and chars:
                if chars[-1] != " ":
                    chars.append(" ")
                    owners.append(r)
                if tabs is not None:
                    tabs[len(chars) - 1] = r
            elif part.tag == qn("w:t"):
                for c in part.text or "":
                    if _INVISIBLE.match(c):
                        continue
                    if not c.isspace():
                        chars.append(c)
                        owners.append(r)
                    elif chars and chars[-1] != " ":
                        chars.append(" ")
                        owners.append(r)
    while chars and chars[-1] == " ":
        chars.pop()
        owners.pop()
        if tabs is not None:
            tabs.pop(len(chars), None)
    return "".join(chars), owners, [(min(b, len(chars)), r, br) for b, r, br in breaks]


def _aligned(
    old: str, new: str, whole: list[str] | tuple[str, ...] = ()
) -> tuple[list[int | None], list[tuple[int, int, int]]]:
    """For each character of `new`, the character of `old` it is where the words around it are
    the form's unchanged, None for words put in; then each stretch of `old` the new text keeps
    unchanged, as (start, end in `old`, start in `new`). Each of `whole` (a field's words) is
    compared as one piece, so it is kept or changed whole, never matched in parts."""
    pieces = sorted({w for w in whole if w}, key=len, reverse=True)
    token = re.compile("".join(re.escape(w) + "|" for w in pieces) + _DIFF_TOKEN.pattern)
    a = [(m.start(), m.group()) for m in token.finditer(old)]
    b = [(m.start(), m.group()) for m in token.finditer(new)]
    same: list[int | None] = [None] * len(new)
    kept: list[tuple[int, int, int]] = []
    for op, i1, i2, j1, j2 in _fewest_edits([w for _, w in a], [w for _, w in b]):
        if op == "equal":
            kept.append((a[i1][0], a[i2 - 1][0] + len(a[i2 - 1][1]), b[j1][0]))
            for (sa, word), (sb, _) in zip(a[i1:i2], b[j1:j2], strict=True):
                for k in range(len(word)):
                    same[sb + k] = sa + k
    return same, kept


def _written(text: str) -> str:
    """Words as set_paragraph_text writes them: invisible characters out, spacing single."""
    return " ".join(_INVISIBLE.sub("", text).split())


def set_paragraph_text(
    p, text: str, original: str = "", line_starts: list[int] | None = None
) -> None:
    """Replace a paragraph's words, keeping its properties, its opening runs' formatting, the
    line breaks and cross-reference fields of the words the form had and the new text keeps,
    and any footnote
    reference whose marker ([^fn-N]) survives. The paragraph's footnote references pair in order with the
    markers of its original text; a marker the new text drops takes its reference with it,
    and one it keeps stays where the marker stands. `line_starts`, one for each of the form's
    line breaks, is where in the new text each line after the first starts, where the caller
    knows the lines (a cell whose lines the view showed apart)."""
    text = _written(text)
    runs = p.findall(qn("w:r"))
    tabs: dict[int, object] = {}
    form, owners, breaks = _form_characters(p, tabs)
    fields = _fields_of(p)
    same, kept = _aligned(form, text, [" ".join(v.split()) for _, v in fields])
    at_new = {i: j for j, i in enumerate(same) if i is not None}
    # A first run formatted apart from the words after it (a defined term in bold, a label)
    # keeps its formatting on its own words: so does each run the new text still opens with.
    # The rest takes the formatting of the first run it does not
    lead = 0
    while lead < len(form):
        end = lead
        while end < len(form) and owners[end] is owners[lead]:
            end += 1
        if not text.startswith(form[:end]):
            break
        lead = end
    body = None
    if owners:
        body = owners[min(lead, len(owners) - 1)].find(qn("w:rPr"))
    # a definition the new text opens otherwise than the form's runs do (the form's "[" before
    # the quote dropped): its quoted term keeps the form term's formatting, where the form
    # formats its term as one, apart from the words after it
    term_rpr, term_at = None, range(0)
    a, b = _OPENING_TERM.match(form), _OPENING_TERM.match(text)
    if a and b and b.start(2) >= lead and a.end(3) < len(form):
        term_props = {_props(owners[i].find(qn("w:rPr"))) for i in range(a.start(2), a.end(2)) if form[i] != " "}
        if len(term_props) == 1 and term_props != {_props(owners[a.end(3)].find(qn("w:rPr")))}:
            term_rpr, term_at = owners[a.start(2)].find(qn("w:rPr")), range(b.start(2), b.end(2))
    tails = ["".join(_element_text(x) for x in runs[-1].itersiblings()) for runs, _ in fields]
    titles = [_target_title(p, runs) for runs, _ in fields]
    notes = dict(
        zip(
            _FOOTNOTE.findall(original),
            [r for r in runs if r.find(qn("w:footnoteReference")) is not None],
            strict=False,
        )
    )
    # remove every run, hyperlink and simple field; keep pPr and bookmarks. A simple field's
    # cached result is the paragraph's visible text: left in place, the new words would sit
    # beside it and the paragraph would read twice.
    for child in list(p):
        if child.tag in (
            qn("w:r"), qn("w:hyperlink"), qn("w:ins"), qn("w:del"), qn("w:fldSimple"),
            MATH + "oMathPara", MATH + "oMath",
        ):
            p.remove(child)

    def make_run(words: str, props=None):
        r = copy.deepcopy(runs[0]) if runs else p.makeelement(qn("w:r"), {})
        for c in list(r):
            r.remove(c)
        if props is not None:
            r.append(copy.deepcopy(props))
        t = r.makeelement(qn("w:t"), {})
        t.text = _xml_safe(words)
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        return r

    def emit(words: str, props) -> None:
        """Words with their surviving footnote markers turned back into reference runs."""
        pieces = _FOOTNOTE.split(words)
        markers = _FOOTNOTE.findall(words)
        for i, piece in enumerate(pieces):
            if piece:
                p.append(make_run(piece, props))
            if i < len(markers) and markers[i] in notes:
                p.append(copy.deepcopy(notes[markers[i]]))

    def free(a: int, b: int) -> bool:
        return not any(x < b and a < y for x, y, _ in claimed)

    def unchanged_at(first: int, last: int) -> int | None:
        """Where the form's characters first..last stand in the new text, if the new text
        keeps them in one unchanged stretch; None otherwise."""
        for a, b, j in kept:
            if a <= first and last < b:
                return j + first - a
        return None

    # each field, in the form's order, stays where its words stand as the form had them; one
    # the form follows with its heading's title may take the first other place its words stand
    # still naming that heading (a reference the drafter moved within the sentence). Where
    # neither holds, the words are the drafter's, typed: a "(k)" now standing where the form's
    # "(l)" stood is not the form's "(k)"
    claimed: list[tuple[int, int, list]] = []
    for (field_runs, visible), tail, title in zip(fields, tails, titles, strict=True):
        visible = " ".join(visible.split())
        if not visible:
            continue
        own = [i for i, r in enumerate(owners) if form[i] != " " and any(r is x for x in field_runs)]
        start = unchanged_at(own[0], own[-1]) if own else None
        if (
            start is not None
            and text[start : start + len(visible)] == visible
            and free(start, start + len(visible))
            and _names_its_title(title, tail, text[start + len(visible) :])
        ):
            claimed.append((start, start + len(visible), field_runs))
            continue
        if _title_gap(title, tail) is None:
            continue
        at = next(
            (
                m.start()
                for m in re.finditer(re.escape(visible), text)
                if _names_its_title(title, tail, text[m.end() :]) and free(m.start(), m.end())
            ),
            None,
        )
        if at is not None:
            claimed.append((at, at + len(visible), field_runs))
    # each form line break stays where the words beside it stand, never at the paragraph's
    # edge unless the form had it there. One tied to the word after it goes before what the
    # drafter put in against that word's start (a "[" opening it): that stays with its word
    breaks_at: dict[int, list] = {}
    if line_starts is not None and len(line_starts) == len(breaks):
        for (_b, r, br), j in zip(breaks, line_starts, strict=True):
            breaks_at.setdefault(j, []).append((r, br))
        breaks = []
    for b, r, br in breaks:
        if b > 0 and b - 1 in at_new:
            j: int | None = at_new[b - 1] + 1
        else:
            j = at_new.get(b)
            while j is not None and j > (0 if b == 0 else 1) and same[j - 1] is None and text[j - 1] != " ":
                j -= 1
        if j is None or (j == 0 and b > 0) or (j == len(text) and b < len(form)):
            continue
        breaks_at.setdefault(j, []).append((r, br))
    fields_at = {a: (b, field_runs) for a, b, field_runs in claimed}
    words, props = "", None

    def flush() -> None:
        nonlocal words
        if words:
            emit(words, props)
        words = ""

    j = 0
    while j <= len(text):
        for r, br in breaks_at.pop(j, []):
            flush()
            line = copy.deepcopy(r)
            for c in list(line):
                if c.tag != qn("w:rPr"):
                    line.remove(c)
            line.append(copy.deepcopy(br))
            p.append(line)
        if j == len(text):
            break
        if j in fields_at:
            flush()
            end, field_runs = fields_at[j]
            for r in field_runs:
                p.append(copy.deepcopy(r))
            j = end
            continue
        if text[j] == " " and same[j] in tabs:
            # a space where the form had a tab, the view showing a tab as a space: the tab
            flush()
            tab = copy.deepcopy(tabs[same[j]])
            for c in list(tab):
                if c.tag != qn("w:rPr"):
                    tab.remove(c)
            tab.append(tab.makeelement(qn("w:tab"), {}))
            p.append(tab)
            j += 1
            continue
        rpr = owners[j].find(qn("w:rPr")) if j < lead else term_rpr if j in term_at else body
        if words and _props(rpr) != _props(props):
            flush()
        if not words:
            props = rpr
        words += text[j]
        j += 1
    flush()
    if not p.findall(qn("w:r")):
        p.append(make_run(""))


def _props(rpr) -> bytes:
    """A run's formatting, comparable: its properties without revision marks."""
    from lxml import etree

    if rpr is None:
        return b""
    plain = copy.deepcopy(rpr)
    for x in plain.findall(qn("w:rPrChange")):
        plain.remove(x)
    return etree.tostring(plain)


def _most_words_rpr(p):
    """The font face and size that carry most of a paragraph's words, or None where its words
    set none: only the typeface, never bold, italics or colour, which body text takes from
    its style."""
    weight: dict[bytes, int] = {}
    first: dict[bytes, object] = {}
    for r in p.findall(qn("w:r")):
        n = len("".join(t.text or "" for t in r.findall(qn("w:t"))).strip())
        if not n:
            continue
        rpr = r.find(qn("w:rPr"))
        face = r.makeelement(qn("w:rPr"), {})
        for tag in ("w:rFonts", "w:sz", "w:szCs"):
            x = rpr.find(qn(tag)) if rpr is not None else None
            if x is not None:
                face.append(copy.deepcopy(x))
        key = _props(face) if len(face) else b""  # b"": words with no face of their own
        weight[key] = weight.get(key, 0) + n
        first.setdefault(key, face)
    if not weight:
        return None
    best = max(weight, key=lambda k: weight[k])
    return first[best] if best else None


def _reads_as(p, piece: str, original: str) -> bool:
    """A cell paragraph reads as `piece`: its words and their spacing, a tab a space as the
    view shows it, and the footnote markers the piece keeps of the original's."""
    same = _shown(piece) == _shown(_element_text(p, tabs=True))
    return same and (not original or _FOOTNOTE.findall(piece) == _FOOTNOTE.findall(original))


def _by_line_breaks(
    paras: list, pieces: list[str], originals: list[str]
) -> tuple[list[str], list[str], list[list[int]]] | None:
    """Where the view shows a cell's line breaks as its paragraph breaks (and drops its empty
    paragraphs), the pieces regrouped one to each paragraph, a paragraph's lines run together
    as its words stand either side of a break, with where in those words each new line starts;
    None where the cell has no line break or the original's pieces are not the cell's lines."""
    lines = [_lines_of(p) for p in paras]
    shown = [x for ls in lines for x in ls]
    if not any(len(ls) > 1 for ls in lines) or not len(pieces) == len(originals) == len(shown):
        return None
    if any(_key(a) != _key(b) for a, b in zip(originals, shown, strict=True)):
        return None
    grouped, olds, starts, at = [], [], [], 0
    for ls in lines:
        own = [_written(x) for x in pieces[at : at + len(ls)]]
        grouped.append("".join(own))
        olds.append("".join(originals[at : at + len(ls)]))
        starts.append([sum(map(len, own[:k])) for k in range(1, len(own))])
        at += len(ls)
    return grouped, olds, starts


def _set_cell(tc, text: str, original: str = "") -> None:
    pieces = text.split(PARA_BREAK.strip()) if text else [""]
    pieces = [x.strip() for x in pieces]
    originals = [x.strip() for x in original.split(PARA_BREAK.strip())] if original else []
    paras = tc.findall(qn("w:p"))
    regrouped = _by_line_breaks(paras, pieces, originals)
    line_starts: list[list[int]] = []
    if regrouped is not None:
        # each line written into the paragraph holding it, its line break kept: the view's
        # line written into an empty paragraph beside it is a paragraph the form never had
        pieces, originals, line_starts = regrouped
    while len(paras) < len(pieces):
        paras.append(copy.deepcopy(paras[-1]))
        _fresh_ids(paras[-1])
        tc.append(paras[-1])
    for i, (p, piece) in enumerate(zip(paras, pieces, strict=False)):
        before = originals[i] if i < len(originals) else ""
        if _reads_as(p, piece, before):
            continue  # this cell's words are unchanged: its fields (a page number) stay live
        set_paragraph_text(p, piece, before, line_starts[i] if i < len(line_starts) else None)
    for p in paras[len(pieces) :]:
        set_paragraph_text(p, "")


def _kept_paragraphs(olds: list[str], pieces: list[str]) -> dict[int, str]:
    """Which of a cell's paragraphs (`olds`) each of the fewer paragraphs written for it
    (`pieces`) is, in order: those reading alike first, then each piece the most like it of
    the paragraphs between. A paragraph no piece is is the one the cell lost."""
    import difflib

    kept: dict[int, str] = {}
    matcher = difflib.SequenceMatcher(a=[_key(x) for x in olds], b=[_key(x) for x in pieces], autojunk=False)
    for tag, a0, a1, b0, b1 in matcher.get_opcodes():
        if tag == "equal":
            kept.update({a0 + k: pieces[b0 + k] for k in range(a1 - a0)})
            continue
        start = a0
        for b in range(b0, b1):
            room = a1 - (b1 - b)  # leave a paragraph for every piece after this one
            if room < start:
                kept[a1 - 1] = PARA_BREAK.join([kept[a1 - 1], pieces[b]]) if a1 - 1 in kept else pieces[b]
                continue
            best = max(
                range(start, room + 1),
                key=lambda a: difflib.SequenceMatcher(a=_key(olds[a]), b=_key(pieces[b])).ratio(),
            )
            kept[best] = pieces[b]
            start = best + 1
    return kept


def _set_row(
    tr,
    text: str,
    original: str = "",
    deferred: list[tuple[str, str, str]] | None = None,
    struck: list[str] | None = None,
) -> None:
    """With `deferred` given, a cell gaining paragraphs is left for after the compare, which
    runs a cell's new paragraphs into one: its last paragraph is deferred carrying the extra
    pieces, rewritten afterwards by rewrite_paragraphs. With `struck` given too, a cell losing
    paragraphs has the ids of those it lost appended there, struck after the compare, and the
    paragraphs it keeps deferred with their new words."""
    cells = [x.strip() for x in text.split(" | ")]
    originals = [x.strip() for x in original.split(" | ")] if original else []
    for i, (tc, value) in enumerate(zip(tr.findall(qn("w:tc")), cells, strict=False)):
        before = originals[i] if i < len(originals) else ""
        if before and _spaced(value) == _spaced(before):
            continue  # a cell as the form reads, beside one that changed: its line breaks stand
        pieces = [x.strip() for x in value.split(PARA_BREAK.strip())]
        paras = tc.findall(qn("w:p"))
        olds = [x.strip() for x in before.split(PARA_BREAK.strip())] if before else []
        if (
            deferred is not None
            and struck is not None
            and len(pieces) < len(paras) == len(olds)
            and all(map(_para_id, paras))
        ):
            kept = _kept_paragraphs(olds, pieces)
            for k, p in enumerate(paras):
                if k not in kept:
                    struck.append(_para_id(p))
                elif not _reads_as(p, kept[k], olds[k]):
                    deferred.append((_para_id(p), kept[k], olds[k]))
            continue
        if deferred is not None and len(pieces) > len(paras) and all(map(_para_id, paras)):
            last = len(paras) - 1
            for k, p in enumerate(paras):
                piece = pieces[k] if k < last else PARA_BREAK.join(pieces[last:])
                old = olds[k] if k < len(olds) else ""
                if k == last or not _reads_as(p, piece, old):
                    deferred.append((_para_id(p), piece, old))
            continue
        _set_cell(tc, value, before)


def _is_heading(p) -> bool:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return False
    if ppr.find(qn("w:outlineLvl")) is not None:
        return True
    style = ppr.find(qn("w:pStyle"))
    val = style.get(qn("w:val")) if style is not None else ""
    return bool(re.search(r"heading|title|L1$", val or "", re.I))


def _new_paragraph_after(
    anchor, text: str, heading: bool = False, like=None, styles=None, restyled: bool = False
):
    """A new paragraph after `anchor`, formatted like it, or, after a heading, like the body
    paragraph that follows the heading, so a new sub-clause joins the list it belongs to. A
    new heading takes the style of `like`, the form heading the drafter named (the sub-clause
    it follows), or failing that of the nearest heading above the anchor. After a title (a
    heading, or a numbered paragraph whose style `styles` keeps with the next, a schedule's
    Part) a new heading is a level beneath it. A new heading, or a paragraph the drafter
    styled (`restyled`) like a form title, takes that title's run font too."""
    template = anchor
    style_up = False
    sub_level = None
    if heading:
        if like is not None and like.tag == qn("w:p"):
            template = like
        elif (
            anchor.tag == qn("w:p")
            and _level_of(anchor) is not None
            and (_is_heading(anchor) or (styles is not None and _keeps_with_next(anchor, styles)))
        ):
            # a sub-heading under a title: the level the form gives that title's sub-headings
            template = anchor
            sub_level = _child_level(anchor)
        elif anchor.tag == qn("w:p") and _level_of(anchor) is not None:
            style_up = True  # the form's numbering styles end in the level: one level up
        else:
            prev = anchor if anchor.tag == qn("w:p") else anchor.getparent()
            while prev is not None and not (prev.tag == qn("w:p") and _is_heading(prev)):
                prev = prev.getprevious()
            if prev is not None:
                template = prev
                if _level_of(prev) is not None:
                    # the nearest title above is the schedule's or clause's own: a heading
                    # written after it (after a table, say) is its sub-heading, not another title
                    sub_level = _child_level(prev)
    elif like is not None and like.tag == qn("w:p"):
        template = like  # a body paragraph of a new sub-clause: styled like the anchor it followed
    elif anchor.tag == qn("w:p") and _is_heading(anchor):
        # a body paragraph joining a list: styled like the nearest body paragraph, before the
        # heading if there is one (a new sub-clause's first paragraph), else after it
        prev = anchor.getprevious()
        while prev is not None and not (prev.tag == qn("w:p") and not _is_heading(prev)):
            prev = prev.getprevious()
        nxt = anchor.getnext()
        if prev is not None:
            template = prev
        elif nxt is not None and nxt.tag == qn("w:p"):
            template = nxt
    new = (
        copy.deepcopy(template) if template.tag == qn("w:p") else anchor.makeelement(qn("w:p"), {})
    )
    # the font and size the model's words carry where the form sets them on its runs, not in
    # the style: the formatting of most of its words, so a bold term is not spread
    words_rpr = _most_words_rpr(template) if template.tag == qn("w:p") else None
    _fresh_ids(new)
    if anchor.tag != qn("w:p"):
        # after a table row: the paragraph goes after the table
        tbl = anchor.getparent()
        tbl.addnext(new)
    else:
        anchor.addnext(new)
    for child in list(new):
        if child.tag != qn("w:pPr"):
            new.remove(child)
    ppr = new.find(qn("w:pPr"))
    if ppr is not None:
        for tag in ("w:rPr", "w:pPrChange", "w:sectPr"):
            for x in ppr.findall(qn(tag)):
                ppr.remove(x)
    if style_up:
        _set_level(new, (_level_of(anchor) or 2) - 1)
    if sub_level is not None:
        _set_level(new, sub_level)
    set_paragraph_text(new, text)
    columns = _tab_columns(template, " ".join(text.split())) if template.tag == qn("w:p") else None
    if columns is not None:
        # laid out as the line it was modelled on: a typed contents line's number, title and
        # page each in its column
        for r in new.findall(qn("w:r")):
            new.remove(r)
        for k, words in enumerate(columns):
            if k:
                tab = new.makeelement(qn("w:r"), {})
                tab.append(tab.makeelement(qn("w:tab"), {}))
                new.append(tab)
            run = new.makeelement(qn("w:r"), {})
            t = run.makeelement(qn("w:t"), {})
            t.text = _xml_safe(words)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            run.append(t)
            new.append(run)
    term = (
        _term_formatting(template, " ".join(text.split())) or _heading_formatting(template, " ".join(text.split()))
        if template.tag == qn("w:p")
        else None
    )
    if term is not None:
        # a definition among definitions, a limb among headed limbs: its term or heading
        # formatted as the one it was modelled on
        for r in new.findall(qn("w:r")):
            new.remove(r)
        for words, rpr in term:
            run = new.makeelement(qn("w:r"), {})
            if rpr is not None:
                run.append(copy.deepcopy(rpr))
            t = run.makeelement(qn("w:t"), {})
            t.text = _xml_safe(words)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            run.append(t)
            new.append(run)
    own = (
        _title_formatting(like, " ".join(text.split()))
        if (heading or restyled) and term is None and template is like
        else None
    )
    if own is not None:
        # a title styled like a form title: its words in the form title's own font, and its
        # mark (its number) too where the form title's mark is formatted as its words
        for r in new.findall(qn("w:r")):
            if r.find(qn("w:rPr")) is None:
                r.insert(0, copy.deepcopy(own))
        mark = like.find(f"{qn('w:pPr')}/{qn('w:rPr')}")
        if ppr is not None and mark is not None and _props(mark) == _props(own):
            ppr.append(copy.deepcopy(mark))
    if words_rpr is not None:
        # words still without a face of their own: the typeface of the model's words
        for r in new.findall(qn("w:r")):
            if r.find(qn("w:rPr")) is None:
                r.insert(0, copy.deepcopy(words_rpr))
    return new


def _title_formatting(model, text: str):
    """The run formatting of `model`, a title ("Confirmation") formatted as one throughout,
    where `text` reads as a title too, no sentence or caption ending in it; None otherwise."""
    form, owners, _ = _form_characters(model)
    if not all(_is_heading_lead(x) and not _HEADING_END.search(x) for x in (form, text)):
        return None
    props = {_props(owners[i].find(qn("w:rPr"))) for i, c in enumerate(form) if c != " "}
    return owners[0].find(qn("w:rPr")) if len(props) == 1 else None


def _tab_columns(model, text: str) -> list[str] | None:
    """`text` split into the columns `model` lays its words out in with tabs: each column of
    the model keeps its number of words but the longest, which takes the rest. None where the
    model has no tab between its words, the text has too few words to fill its columns, or a
    column other than the longest is not a number shaped as the model's."""
    columns, current = [], ""
    for r in model.findall(qn("w:r")):
        for part in r:
            if part.tag == qn("w:tab"):
                columns.append(current)
                current = ""
            elif part.tag == qn("w:t"):
                current += part.text or ""
    columns = [c.split() for c in [*columns, current]]
    columns = [c for c in columns if c]
    words = text.split()
    if len(columns) < 2 or len(words) < len(columns):
        return None
    wide = max(range(len(columns)), key=lambda k: len(" ".join(columns[k])))
    head = sum(len(c) for c in columns[:wide])
    tail = sum(len(c) for c in columns[wide + 1 :])
    if head + tail >= len(words):
        return None
    out, at = [], 0
    for k, c in enumerate(columns):
        n = len(words) - head - tail if k == wide else len(c)
        out.append(" ".join(words[at : at + n]))
        at += n
    # the narrow columns are numbers of the model's shape ("6." and "7.", "8" and "9"): any
    # other words are not laid out by the model's columns
    if any(
        _number_shape(" ".join(c)) != _number_shape(o) or not re.search(r"\d|\([a-z]+\)", o)
        for k, (c, o) in enumerate(zip(columns, out, strict=True))
        if k != wide
    ):
        return None
    return out


_QUOTED_TERM = re.compile(r"([\"“])([^\"“”]+)([\"”])")
_OPENING_TERM = re.compile(r"[^\w\"“]*([\"“])([^\"“”]+)([\"”])")  # after a "[" or "(", say


def _term_formatting(model, text: str) -> list[tuple[str, object]] | None:
    """Where `model` opens with a quoted term formatted apart from the words after it (a
    definition, its term in bold) and `text` opens with a quoted term too: `text` in pieces,
    each with the formatting of the model's piece in its place (opening quote, term, closing
    quote, the rest). None otherwise, and for text holding a footnote marker."""
    form, owners, _ = _form_characters(model)
    a, b = _QUOTED_TERM.match(form), _QUOTED_TERM.match(text)
    after = a.end() + len(form[a.end() :]) - len(form[a.end() :].lstrip()) if a else 0
    if a is None or b is None or after >= len(form) or _FOOTNOTE.search(text):
        return None

    def rpr(i: int):
        return owners[i].find(qn("w:rPr"))

    # the words after the term are formatted as the first of them, not as a space beside it
    term, rest = rpr(a.start(2)), rpr(after)
    if _props(term) == _props(rest):
        return None
    pieces = [
        (b.group(1), rpr(a.start(1))),
        (b.group(2), term),
        (b.group(3), rpr(a.start(3))),
        (text[b.end() :], rest),
    ]
    out: list[tuple[str, object]] = []
    for words, props in pieces:
        if out and _props(out[-1][1]) == _props(props):
            out[-1] = (out[-1][0] + words, out[-1][1])
        elif words:
            out.append((words, props))
    return out


_HEADING_END = re.compile(r"[:.](?=\s|$)")


def _is_heading_lead(words: str) -> bool:
    """Words shaped as a limb's heading ("Change of Control", "Late Payment"): no
    clause punctuation, first and last word opening with a capital."""
    names = words.split()
    return bool(names) and not re.search(r"[,;()\[\]\"“”]", words) and names[0][0].isupper() and names[-1][0].isupper()


def _heading_formatting(model, text: str) -> list[tuple[str, object]] | None:
    """Where `model` opens with a heading formatted apart from the words after it (a limb's
    "Repudiation" in italic, then ": the Customer rescinds ...") up to its first ":" or ".", and
    `text` opens with a heading too, or is its heading alone ("Guarantee:"): `text` in pieces,
    its heading, the mark ending it and the rest each formatted as the model's are. None
    otherwise, and for text holding a footnote marker."""
    form, owners, _ = _form_characters(model)
    a, b = _HEADING_END.search(form), _HEADING_END.search(text)
    if a is None or b is None or _FOOTNOTE.search(text):
        return None
    if not (_is_heading_lead(form[: a.start()]) and _is_heading_lead(text[: b.start()])):
        return None
    # the words after the heading as the first of the model's, or, where the model is its
    # heading alone, as the mark ending it
    after = a.end() + len(form[a.end() :]) - len(form[a.end() :].lstrip())
    lead = {_props(owners[i].find(qn("w:rPr"))) for i in range(a.start()) if form[i] != " "}
    rest_at = after if after < len(form) else a.start()
    head, mark, rest = (owners[i].find(qn("w:rPr")) for i in (0, a.start(), rest_at))
    if len(lead) != 1 or _props(head) == _props(rest):
        return None
    out: list[tuple[str, object]] = []
    for words, props in ((text[: b.start()], head), (text[b.start()], mark), (text[b.end() :], rest)):
        if out and _props(out[-1][1]) == _props(props):
            out[-1] = (out[-1][0] + words, out[-1][1])
        elif words:
            out.append((words, props))
    return out


def _style_id(p) -> str:
    ppr = p.find(qn("w:pPr"))
    style = ppr.find(qn("w:pStyle")) if ppr is not None else None
    return (style.get(qn("w:val")) or "") if style is not None else ""


def _style_family(p) -> str:
    return re.sub(r"L\d+$", "", _style_id(p))


def _child_level(title) -> int:
    """The numbering level the form gives the sub-headings of a title like `title`: the level of
    the first numbered paragraph under it, or, where the form keeps it empty, under the nearest
    title of its level that has any. A schedule's title (level 1) is followed by level 3, a
    clause's by level 2."""
    level = _level_of(title) or 1
    family = _style_family(title)

    def under(first):
        cur = first
        while cur is not None:
            if cur.tag == qn("w:p") and _style_family(cur) == family:
                lvl = _level_of(cur)
                if lvl is not None:
                    return lvl if lvl > level else None
            cur = cur.getnext()
        return None

    found = under(title.getnext())
    if found is not None:
        return found
    for p in title.getroottree().getroot().iter(qn("w:p")):
        if _style_family(p) == family and _level_of(p) == level and p is not title:
            found = under(p.getnext())
            if found is not None:
                return found
    return level + 1


def _unnumbered_body(title, styles):
    """How the form lays out a sub-clause whose words open unnumbered under its heading
    `title`: its first paragraph (body text, or a list paragraph with its numbering switched
    off) and its first numbered limb after that, or None for that limb. None where the
    sub-clause opens straight into numbered limbs, with a title, or with nothing under it."""
    level = _level_of(title)
    first = _next_with_text(title) if level is not None else None
    if first is None or _is_heading(first) or _keeps_with_next(first, styles):
        return None
    if _level_of(first) is not None and (
        (_level_of(first) or 0) <= (level or 0) or not _numbering_off(first)
    ):
        return None
    limb = None
    cur = first.getnext()
    while cur is not None and cur.tag == qn("w:p") and not _is_heading(cur):
        lvl = _level_of(cur)
        if lvl is not None and lvl <= (level or 0):
            break
        if lvl is not None and not _numbering_off(cur):
            limb = cur
            break
        cur = cur.getnext()
    return first, limb


def _numbering_off(p) -> bool:
    """A list-styled paragraph whose own properties switch its numbering off (numId 0)."""
    num = p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
    return num is not None and num.get(qn("w:val")) == "0"


def _next_with_text(p):
    """The next paragraph that has words in it: the form leaves blank lines under a title it
    keeps empty, and they hide that the next real paragraph is another title."""
    nxt = p.getnext()
    while nxt is not None and nxt.tag == qn("w:p") and not _element_text(nxt).strip():
        nxt = nxt.getnext()
    return nxt if nxt is not None and nxt.tag == qn("w:p") else None


def _nothing_under(anchor) -> bool:
    """The anchor is followed, past any blank lines, by a title or by nothing."""
    nxt = _next_with_text(anchor)
    return nxt is None or _is_heading(nxt)


def _body_template_for(anchor, styles=None):
    """The paragraph whose style a new body paragraph after `anchor` takes: the anchor itself,
    or, where the anchor is a title (a heading style, or a paragraph followed by a deeper list
    level), the first body paragraph under it. A title with nothing under it (an appendix the
    form names and leaves empty, so the next paragraph is a title at the same or a shallower
    level) lends the nearest body paragraph before it, so the paragraphs written under it read
    as body text, not as titles. A title whose next paragraph is itself a title gives, with the
    document's `styles`, a paragraph in the style its own style names to follow it."""
    if anchor.tag != qn("w:p"):
        return None
    nxt = anchor.getnext()
    nxt = nxt if nxt is not None and nxt.tag == qn("w:p") else None
    # an unnumbered body paragraph sits at its enclosing sub-clause's level: the heading that
    # follows it is a sibling of that sub-clause, not a deeper list under the paragraph
    level, _ = _enclosing_level(anchor)
    deeper = nxt is not None and (_level_of(nxt) or 0) > (level or 0)
    # an empty title: a titled paragraph followed straight by a title at a shallower level,
    # sitting in a run of body text (an appendix set out as running text), not in a numbered list
    own = _level_of(anchor)
    prev = anchor.getprevious()
    while prev is not None and prev.tag == qn("w:p") and not _element_text(prev).strip():
        prev = prev.getprevious()
    body_before = (
        prev is not None
        and prev.tag == qn("w:p")
        and _level_of(prev) is None
        and not _is_heading(prev)
    )
    after_blanks = _next_with_text(anchor)
    shallower_title_after = (
        after_blanks is not None
        and own is not None
        and _is_heading(after_blanks)
        and (_level_of(after_blanks) or 0) < own
    )
    # the form's own blank line under an empty title is its body paragraph: its style is what
    # text written under the title takes
    blank_body = (
        nxt
        if nxt is not None
        and not _element_text(nxt).strip()
        and _level_of(nxt) is None
        and not _is_heading(nxt)
        else None
    )
    if shallower_title_after and blank_body is not None:
        return blank_body
    if shallower_title_after and body_before:
        return prev
    if _is_heading(anchor) and nxt is not None and styles is not None and _keeps_with_next(
        nxt, styles
    ):
        # a title straight under the title (a schedule's Part A): it is no body paragraph to
        # lend its style; the form names the style of the paragraph that follows the title
        template = _follower_of(anchor, styles)
        if template is not None:
            return template
    if _is_heading(anchor) or deeper:
        return nxt
    return anchor


def _follower_of(title, styles):
    """A paragraph in the style the title's style names for the paragraph after it (Word's
    "style for following paragraph"), or None where it names none."""
    style = _style_def(_style_id(title), styles)
    follow = style.find(qn("w:next")) if style is not None else None
    if follow is None or not follow.get(qn("w:val")):
        return None
    template = title.makeelement(qn("w:p"), {})
    ppr = template.makeelement(qn("w:pPr"), {})
    ppr.append(ppr.makeelement(qn("w:pStyle"), {qn("w:val"): follow.get(qn("w:val"))}))
    template.append(ppr)
    return template


def _style_def(style_id: str, styles):
    """The w:style of the document's styles part with this id, or None."""
    return next(
        (s for s in styles.iterchildren(qn("w:style")) if s.get(qn("w:styleId")) == style_id),
        None,
    )


def _keeps_with_next(p, styles) -> bool:
    """Whether a paragraph's style keeps it with the next: Word's mark of a title style, which
    never stands at the foot of a page apart from what it heads. A body paragraph kept with
    the next by its own properties (a "where:" before its list) is not a title."""
    style = _style_def(_style_id(p), styles)
    return style is not None and style.find(f"{qn('w:pPr')}/{qn('w:keepNext')}") is not None


def _enclosing_level(p):
    """The list level a paragraph sits at, and the paragraph that carries it: its own, or, for
    an unnumbered body paragraph, the level of the sub-clause it sits in: the shallowest
    numbered paragraph above it up to and including its sub-clause heading (a level 1 or 2
    paragraph, or a heading style). A closing paragraph after (a), (b), (c) limbs sits at the
    sub-clause's level, so new paragraphs after it can join the limbs."""
    if p.tag == qn("w:p") and _level_of(p) is not None:
        return _level_of(p), p
    level, carrier = None, None
    cur = p.getprevious()
    while cur is not None:
        if cur.tag == qn("w:p"):
            lvl = _level_of(cur)
            if lvl is not None and (level is None or lvl < level):
                level, carrier = lvl, cur
            if (lvl is not None and lvl <= 2) or _is_heading(cur):
                break
        cur = cur.getprevious()
    return level, carrier


def _level_of(p) -> int | None:
    """The list level of a paragraph whose style name ends in the level ("Body L3")."""
    ppr = p.find(qn("w:pPr"))
    style = ppr.find(qn("w:pStyle")) if ppr is not None else None
    m = re.search(r"L(\d+)$", style.get(qn("w:val")) or "") if style is not None else None
    return int(m.group(1)) if m else None


def _set_level(p, level: int) -> None:
    style = p.find(qn("w:pPr")).find(qn("w:pStyle"))
    style.set(qn("w:val"), re.sub(r"L\d+$", f"L{max(level, 1)}", style.get(qn("w:val"))))


def _numbering_of(doc):
    """The document's numbering part, or None where it has none."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        return doc.part.part_related_by(RT.NUMBERING).element
    except KeyError:
        return None


def _set_level_below(p, parent, numbering) -> None:
    """Paragraph `p` one list level below `parent`, where Word's numbering links each level of
    `parent`'s list to a style of its own (a form whose level styles are not named "...L<n>"):
    the style linked to the level after `parent`'s. Unchanged where nothing is linked."""
    style = _style_id(parent)
    if numbering is None or not style:
        return
    for abstract in numbering.iter(qn("w:abstractNum")):
        linked = {
            int(lvl.get(qn("w:ilvl")) or 0): ps.get(qn("w:val"))
            for lvl in abstract.findall(qn("w:lvl"))
            if (ps := lvl.find(qn("w:pStyle"))) is not None
        }
        own = next((k for k, v in linked.items() if v == style), None)
        if own is None or not linked.get(own + 1):
            continue
        p.find(qn("w:pPr")).find(qn("w:pStyle")).set(qn("w:val"), linked[own + 1])
        ilvl = p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:ilvl')}")
        if ilvl is not None:
            ilvl.set(qn("w:val"), str(own + 1))
        return


def _take_style(p, model) -> None:
    """Paragraph `p` takes `model`'s paragraph properties (its style, numbering and indents),
    keeping its own paragraph mark and section break."""
    old = p.find(qn("w:pPr"))
    source = model.find(qn("w:pPr"))
    new = copy.deepcopy(source) if source is not None else p.makeelement(qn("w:pPr"), {})
    for tag in ("w:rPr", "w:sectPr", "w:pPrChange"):
        for x in new.findall(qn(tag)):
            new.remove(x)
    if old is None:
        p.insert(0, new)
        return
    for tag in ("w:rPr", "w:sectPr"):
        x = old.find(qn(tag))
        if x is not None:
            new.append(x)
    p.replace(old, new)


def restyle_paragraphs(
    redline_bytes: bytes, plain_bytes: bytes, ids: list[str], author: str = "Drafting engine"
) -> bytes:
    """Each paragraph given by id takes the paragraph properties it has in the plain document,
    as a tracked formatting change: the compare keeps the form's style and numbering. The first
    paragraph carrying the id is the form's; run after repair_fields has dropped any twin."""
    wanted = set(ids)
    if not wanted:
        return redline_bytes
    plain = {
        _para_id(p): p
        for p in Document(io.BytesIO(plain_bytes)).element.body.iter(qn("w:p"))
        if _para_id(p) in wanted
    }
    doc = Document(io.BytesIO(redline_bytes))
    n = 98000
    stamp = "2026-01-01T00:00:00Z"
    rows: dict = {}
    for p in list(doc.element.body.iter(qn("w:p"))):
        pid = _para_id(p)
        if pid not in wanted or pid not in plain:
            continue
        wanted.discard(pid)
        old = p.find(qn("w:pPr"))
        before = copy.deepcopy(old) if old is not None else p.makeelement(qn("w:pPr"), {})
        for tag in ("w:rPr", "w:sectPr", "w:pPrChange"):
            for x in before.findall(qn(tag)):
                before.remove(x)
        _take_style(p, plain[pid])
        n += 1
        change = p.makeelement(
            qn("w:pPrChange"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
        )
        change.append(before)
        p.find(qn("w:pPr")).append(change)
        twin = plain[pid].getparent()
        if p.getparent().tag == qn("w:tc") and twin.tag == qn("w:tc"):
            rows.setdefault(p.getparent().getparent(), twin.getparent())
    # a row restyled like another (its cells' shading, width, run formatting): the compare
    # keeps the form's cells, so the row takes the plain document's, tracked
    for tr, twin in rows.items():
        n = _track_row_look(tr, twin, n, author, stamp)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _look(el) -> bytes:
    """Properties (a cell's, a run's), comparable across documents: without revision marks."""
    from lxml import etree

    if el is None:
        return b""
    plain = copy.deepcopy(el)
    for x in plain.findall(qn("w:tcPrChange")) + plain.findall(qn("w:rPrChange")):
        plain.remove(x)
    return etree.tostring(plain, method="c14n", exclusive=True)


def _take_row_look(tr, model) -> bool:
    """Row `tr` takes `model`'s cells: each cell's properties (shading, borders, width, the
    columns it spans), its paragraphs' properties and its runs' formatting. The cells beyond
    the model's go; False, the row unchanged, where they hold words or the row has fewer."""
    cells, models = tr.findall(qn("w:tc")), model.findall(qn("w:tc"))
    if len(cells) < len(models) or any(_element_text(tc).strip() for tc in cells[len(models) :]):
        return False
    for tc in cells[len(models) :]:
        tr.remove(tc)
    for tc, m in zip(cells, models, strict=False):
        _set_child(tc, "w:tcPr", m.find(qn("w:tcPr")))
        mp = m.find(qn("w:p"))
        rpr = m.find(f".//{qn('w:r')}/{qn('w:rPr')}")
        for p in tc.findall(qn("w:p")):
            if mp is not None:
                _take_style(p, mp)
            for r in p.iter(qn("w:r")):
                _set_child(r, "w:rPr", rpr)
    return True


def _set_child(el, tag: str, source) -> None:
    """`el`'s first `tag` child made a copy of `source`, revision marks aside; removed for none."""
    old = el.find(qn(tag))
    if old is not None:
        el.remove(old)
    if source is None:
        return
    new = copy.deepcopy(source)
    for x in new.findall(qn(tag + "Change")):
        new.remove(x)
    el.insert(0, new)


def _track_row_look(tr, twin, n: int, author: str, stamp: str) -> int:
    """Row `tr` of the redline given row `twin`'s cells, where they differ: the cells beyond
    `twin`'s, holding no words, go; each cell's properties and run formatting become `twin`'s,
    as tracked formatting changes. Returns the last revision id used."""
    cells, models = tr.findall(qn("w:tc")), twin.findall(qn("w:tc"))
    if len(cells) == len(models) and all(
        _look(a.find(qn("w:tcPr"))) == _look(b.find(qn("w:tcPr"))) for a, b in zip(cells, models, strict=True)
    ):
        return n
    if len(cells) < len(models) or any(_element_text(tc).strip() for tc in cells[len(models) :]):
        return n
    for tc in cells[len(models) :]:
        tr.remove(tc)
    for tc, m in zip(cells, models, strict=False):
        for tag, el, source in [("w:tcPr", tc, m.find(qn("w:tcPr")))] + [
            ("w:rPr", r, m.find(f".//{qn('w:r')}/{qn('w:rPr')}")) for r in tc.iter(qn("w:r"))
        ]:
            if _look(el.find(qn(tag))) == _look(source):
                continue
            before = copy.deepcopy(el.find(qn(tag)))
            if before is None:
                before = el.makeelement(qn(tag), {})
            for x in before.findall(qn(tag + "Change")):
                before.remove(x)
            _set_child(el, tag, source if source is not None else el.makeelement(qn(tag), {}))
            n += 1
            change = el.makeelement(
                qn(tag + "Change"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            change.append(before)
            el.find(qn(tag)).append(change)
    return n


def _new_table_after(anchor, header: list[str], rows: list[list[str]]):
    """A new table after `anchor` (after the whole table, if the anchor is a row), styled like
    the document's first table so it looks like the form's own, followed by an empty paragraph
    because Word wants a paragraph between a table and whatever comes next."""
    host = anchor if anchor.tag == qn("w:p") else anchor.getparent()
    tbl = host.makeelement(qn("w:tbl"), {})
    template = next(iter(host.getroottree().getroot().iter(qn("w:tbl"))), None)
    template_pr = template.find(qn("w:tblPr")) if template is not None else None
    if template_pr is not None:
        tbl.append(copy.deepcopy(template_pr))
    else:
        pr = tbl.makeelement(qn("w:tblPr"), {})
        borders = pr.makeelement(qn("w:tblBorders"), {})
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(
                borders.makeelement(qn(f"w:{side}"), {qn("w:val"): "single", qn("w:sz"): "4"})
            )
        pr.append(borders)
        tbl.append(pr)
    grid = tbl.makeelement(qn("w:tblGrid"), {})
    for _ in header:
        grid.append(grid.makeelement(qn("w:gridCol"), {}))
    tbl.append(grid)
    for cells in [header, *rows]:
        tr = tbl.makeelement(qn("w:tr"), {})
        for text in cells:
            tc = tr.makeelement(qn("w:tc"), {})
            para = tc.makeelement(qn("w:p"), {})
            set_paragraph_text(para, text)
            tc.append(para)
            tr.append(tc)
        tbl.append(tr)
    host.addnext(tbl)
    trailing = tbl.makeelement(qn("w:p"), {})
    tbl.addnext(trailing)
    return trailing


def write(
    base_bytes: bytes,
    snapshots: list[dict],
    deferred_rows: list[list[str]] | None = None,
    deferred_paragraphs: list[str] | None = None,
    deferred_cells: list[tuple[str, str, str]] | None = None,
    deferred_styles: list[str] | None = None,
) -> tuple[bytes, list[str]]:
    """Apply every unit's final paragraphs to a plain copy of the document. Returns the
    document and the originals that could not be found. With `deferred_rows` given, a table
    row to be deleted stays in the document and its paragraphs' ids are appended to that list
    instead: the comparison engine misaligns a table whose rows have gone, so rows are struck
    after the compare by strike_rows, as tracked deletions. A paragraph given another's style
    and numbering has its id appended to `deferred_styles`: the compare keeps the form's
    paragraph properties, so restyle_paragraphs carries them over afterwards."""
    doc = Document(io.BytesIO(base_bytes))
    body = _Body(doc)
    numbering = _numbering_of(doc)
    missing: list[str] = []
    # the bookmarks a cross-reference field names: a paragraph holding one is never struck as empty
    referenced = {
        m.group(1)
        for it in doc.element.body.iter(qn("w:instrText"))
        for m in _REF_FIELD.finditer(it.text or "")
    }
    for snap in snapshots:
        # the pictures and empty paragraphs after a form paragraph, first, while every form
        # paragraph still stands where the form has it
        # (found here, struck once the unit's deletions are known: limbs the unit strikes
        # do not hold their empty parent)
        empties: list[tuple[str, list]] = []
        for before, rank, where in snap.get("struck_empty", []):
            el = _located(body, before, rank, where)
            empty = _empty_after(el) if el is not None else []
            if not empty:
                missing.append(f"empty paragraphs after: {before[:80]}")
            empties.append((before, empty))
        cursor = 0
        changed_at: dict[str, list[int]] = {}  # where this unit's paragraphs were found, by words
        change_ranks = list(snap.get("changed_ranks") or [])
        change_context = list(snap.get("changed_context") or [])
        change_like = list(snap.get("changed_like") or [])
        for n, (original, final) in enumerate(snap.get("changed", [])):
            like = change_like[n] if n < len(change_like) else ""
            rank = change_ranks[n] if n < len(change_ranks) else -1
            prior, k = change_context[n] if n < len(change_context) else ("", -1)
            at = _place_of(body, prior, changed_at)
            hit = body.take_following(original, prior, k)
            if hit is None and rank >= 0:
                hit = _after(body, body.take_nth(original, rank), at)
            if hit is None:
                hit = body.take(original, cursor if at is None else at + 1)
            line = _row_line(body, original, at) if hit is None else None
            if hit is None and line is None:
                hit = _taken(body, body.in_place(original, prior, k, rank))
            if hit is None and line is None and _replace_within(body, original, final):
                continue
            if hit is None:
                finals = _line_cells(final)
                if line is None or len(finals) != len(line):
                    missing.append(original[:100])
                    continue
                changed_at.setdefault(_key(original), []).append(line[-1])
                for j, new, old in zip(line, finals, _line_cells(original), strict=True):
                    cell_paragraph = body.items[j][1]
                    if _spaced(new) == _spaced(old):
                        continue
                    if (
                        deferred_cells is not None
                        and len(cell_paragraph.getparent().findall(qn("w:p"))) > 1
                        and _para_id(cell_paragraph)
                    ):
                        # a cell of several paragraphs: rewritten after the compare, as above
                        deferred_cells.append((_para_id(cell_paragraph), new.strip(), old.strip()))
                    else:
                        set_paragraph_text(cell_paragraph, new.strip(), old.strip())
                continue
            i, el = hit
            changed_at.setdefault(_key(original), []).append(i)
            cursor = min(cursor, i)  # units are drafted from one base; order within a unit holds
            if i in body.inline:
                _set_line(el, body.inline[i], final)
            elif el.tag == qn("w:tr"):
                _set_row(el, final, original, deferred_cells, deferred_paragraphs)
            elif (
                deferred_cells is not None
                and el.getparent().tag == qn("w:tc")
                and len(el.getparent().findall(qn("w:p"))) > 1
                and _para_id(el)
            ):
                # The comparison engine rewrites a cell of several paragraphs as one inserted
                # paragraph (the cover's parties block ran together). Left unchanged for the
                # compare, the paragraph is rewritten afterwards as its own tracked change.
                deferred_cells.append((_para_id(el), final, original))
            elif (
                not like
                or _spaced(final) != _spaced(original)
                or _FOOTNOTE.findall(final) != _FOOTNOTE.findall(original)
                or _drops_typed_number(el, final)
            ):
                # restyled with its words as they stand: they are left alone. A number typed
                # into the words ("5. ") is words the view showed, though _key sets it aside;
                # one the drafter's words do not carry goes, the new style numbering it. So
                # does a footnote reference whose marker the drafter's words dropped
                set_paragraph_text(el, final, original)
            if like and el.tag == qn("w:p") and i not in body.inline:
                model = _find_near(body, like, el, struck=True, besides=el)
                if model is None or model.tag != qn("w:p"):
                    missing.append(f"style like: {like[:80]}")
                    continue
                _take_style(el, model)
                if deferred_styles is not None and _para_id(el):
                    deferred_styles.append(_para_id(el))
            if like and el.tag == qn("w:tr"):
                # a row like another: the model is a row, or a one-cell row's paragraph
                model = _find_near(body, like, el)
                if model is not None and model.tag == qn("w:p") and model.getparent().tag == qn("w:tc"):
                    model = model.getparent().getparent()
                if model is None or model.tag != qn("w:tr") or model is el or not _take_row_look(el, model):
                    missing.append(f"style like: {like[:80]}")
                    continue
                if deferred_styles is not None:
                    deferred_styles.extend(x for x in map(_para_id, el.iter(qn("w:p"))) if x)
        cursor = 0
        ranks = list(snap.get("deleted_ranks") or []) or [-1] * len(snap.get("deleted", []))
        contexts = list(snap.get("deleted_context") or [])
        # words the unit struck and wrote again elsewhere: a moved paragraph
        moved = {_key(t) for entry in snap.get("inserted", []) for t in entry[1]}
        moved_from: dict[str, list] = {}  # the form's paragraphs taken out, by their words
        emptied: list = []  # rows a deleted line of the render took paragraphs out of
        struck_at: dict[str, list[int]] = {}  # where this unit's deletions were found
        unit_missing: list[str] = []  # this unit's struck lines found nowhere
        unit_struck: list = []  # the elements this unit struck, for a table among them
        struck_here: list = []  # the paragraphs this unit struck by their ids afterwards
        for n, (original, rank) in enumerate(zip(snap.get("deleted", []), ranks, strict=False)):
            prior, k = contexts[n] if n < len(contexts) else ("", -1)
            at = _place_of(body, prior, struck_at, changed_at)
            hit = body.take_following(original, prior, k)
            if hit is None and rank >= 0:
                found = body.take_nth(original, rank)
                hit = _after(body, found, at)
                if found is not None and hit is None and at is not None:
                    hit = body.take(original, at + 1)
            elif hit is None:
                hit = body.take(original, cursor if at is None else at + 1)
            line = _row_line(body, original, at) if hit is None else None
            if hit is None and line is None:
                # one paragraph of a table cell holding several, shown as a line of its own
                hit = _taken(body, body.in_place(original, prior, k, rank))
            # a rendered line inside a Word paragraph (after a line break): taken out of it
            if hit is None and line is None and _replace_within(body, original, "", cursor):
                continue
            if hit is None:
                if line is None:
                    unit_missing.append(original)
                    continue
                struck_at.setdefault(_key(original), []).append(line[-1])
                for j in line:
                    cell_paragraph = body.items[j][1]
                    if deferred_paragraphs is not None and _para_id(cell_paragraph):
                        deferred_paragraphs.append(_para_id(cell_paragraph))
                    else:
                        set_paragraph_text(cell_paragraph, "")
                    emptied.append(cell_paragraph.getparent().getparent())
                continue
            cursor = min(cursor, hit[0])
            i, el = hit
            # what stands straight after it, read before it is taken out (blanks aside)
            nxt = el.getnext()
            while nxt is not None and nxt.tag == qn("w:p") and not _key(_element_text(nxt)):
                nxt = nxt.getnext()
            unit_struck.append(nxt)
            struck_at.setdefault(_key(original), []).append(i)
            if i in body.inline:
                _set_line(el, body.inline[i], "")
                continue
            if el.tag == qn("w:tr") and deferred_rows is not None:
                deferred_rows.append([_para_id(x) for x in el.iter(qn("w:p"))])
                continue
            if (
                el.getparent().tag == qn("w:tc")
                and deferred_paragraphs is not None
                and _para_id(el)
            ):
                # A paragraph taken out of, or emptied in, a table cell makes the comparison
                # engine rewrite the cell as one merged paragraph. Left as it is for the compare
                # and struck afterwards by its id, the cell keeps its lines.
                deferred_paragraphs.append(_para_id(el))
            elif el.getparent().tag == qn("w:tc"):
                set_paragraph_text(el, "")
            elif deferred_paragraphs is not None and _para_id(el) and _key(original) not in moved:
                # A paragraph removed beside a changed one makes the comparison engine pair
                # the removed one with the changed one's new words. Left in place, the compare
                # sees it unchanged; it is struck afterwards by its id. A moved paragraph is
                # not left: the compare would pair the form's paragraph with its copy where it
                # stood and strike that one, leaving the words where the drafter took them from
                deferred_paragraphs.append(_para_id(el))
                struck_here.append(el)
            else:
                el.getparent().remove(el)
                if _key(original) in moved:
                    moved_from.setdefault(_key(original), []).append(el)
        _strike_empty_between(struck_here, doc, referenced, deferred_paragraphs)
        if unit_missing or emptied:
            # a table the render showed in fragments (a signature table: ")", "PARTY | )") that
            # this unit struck with the text around it: where every line of a table straight
            # after what the unit struck is one the unit struck, the whole table goes; the same
            # table standing elsewhere (the execution page) is not after what it struck
            words = {
                _key(piece)
                for o in snap.get("deleted", [])
                for piece in re.split(r"\s*(?:\||//)\s*", o)
                if _key(piece)
            }
            taken = _strike_fragment_tables(doc, unit_struck, words, deferred_rows)
            unit_missing = [o for o in unit_missing if not taken or not all(
                _key(x) in taken for x in re.split(r"\s*(?:\||//)\s*", o) if _key(x)
            )]
        missing.extend(o[:100] for o in unit_missing)
        gone = set(deferred_paragraphs or [])
        for before, empty in empties:
            why = _holds_structure(empty, doc, referenced, gone)
            if why:
                missing.append(f"empty paragraphs after {before[:60]} kept: {why}")
                continue
            for p in empty:
                if deferred_paragraphs is not None and _para_id(p):
                    deferred_paragraphs.append(_para_id(p))
                elif p.getparent() is not None:
                    p.getparent().remove(p)
        if deferred_rows is not None and deferred_paragraphs is not None:
            # a row the drafter struck line by line, every word of it: struck as a row
            for tr in dict.fromkeys(emptied):
                ids = [_para_id(p) for p in tr.iter(qn("w:p"))]
                worded = [x for p, x in zip(tr.iter(qn("w:p")), ids, strict=True) if _key(_element_text(p))]
                if all(ids) and set(worded) <= set(deferred_paragraphs):
                    deferred_paragraphs[:] = [x for x in deferred_paragraphs if x not in ids]
                    deferred_rows.append(ids)
        inserted = snap.get("inserted", [])
        rows_inserted = snap.get("rows_inserted", [])
        tables_inserted = snap.get("tables_inserted", [])
        insert_ranks = list(snap.get("insert_ranks") or [])
        row_ranks = list(snap.get("row_ranks") or [])
        table_ranks = list(snap.get("table_ranks") or [])

        contexts = snap.get("anchor_context") or {}

        def anchor_for(before: str, rank: int, where: list | None = None):
            el = located(before, rank, where)
            if el is not None and el.getroottree().getroot() is not doc.element:
                # a header or footer line: the render lists them after the body, so what the
                # drafter wrote after one follows the body's last paragraph. Written into the
                # header or footer it would be in no compare and in no report
                return next(reversed(doc.element.body.findall(qn("w:p"))), None)
            return el

        def located(before: str, rank: int, where: list | None = None):
            return _located(body, before, rank, where)

        sequence = snap.get("insert_order") or (
            [("p", i) for i in range(len(inserted))]
            + [("r", i) for i in range(len(rows_inserted))]
            + [("t", i) for i in range(len(tables_inserted))]
        )
        # each keyed by the form's anchor element: two empty rows read alike and are two anchors
        cell_gains: dict[str, int] = {}  # a cell paragraph's id -> its deferred_cells entry
        last: dict = {}  # anchor -> the last element added after it
        templates: dict = {}  # anchor -> the body template as the form had it
        empty_titles: dict = {}  # anchor -> the anchor is a title the form keeps empty
        headed: dict = {}  # anchor -> the last new title written after it
        written: dict = {}  # a new paragraph's words (as _key reads them) -> its element
        for kind, n in sequence:
            if kind == "p":
                entry = inserted[n]
                before, paragraphs = entry[0], entry[1]
                headings = list(entry[2]) if len(entry) > 2 else [False] * len(paragraphs)
                form_anchor = anchor_for(
                    before,
                    insert_ranks[n] if n < len(insert_ranks) else -1,
                    _nth(contexts.get("p"), n),
                )
                if form_anchor is None:
                    missing.append(f"insert after: {before[:80]}")
                    continue
                # a paragraph of a table cell (a one-column table renders its cells as lines):
                # what is written after it goes after the table, never into the cell; a cell
                # of a row of several holding several paragraphs, each its own line in the
                # view, gets the new lines where the view shows them
                cell_row = (
                    next(form_anchor.iterancestors(qn("w:tr")), None)
                    if form_anchor.tag == qn("w:p") and not _in_cell_of_several(form_anchor)
                    else None
                )
                if deferred_cells is not None and _in_cell_of_several(form_anchor) and _para_id(form_anchor):
                    # new lines in a cell of several paragraphs: the compare would rewrite the
                    # cell as one paragraph, so they are put in after it, as the cell's gain
                    words = _element_text(form_anchor)
                    pid = _para_id(form_anchor)
                    at = cell_gains.get(pid)
                    if at is None:
                        cell_gains[pid] = len(deferred_cells)
                        deferred_cells.append((pid, PARA_BREAK.join([words, *paragraphs]), words))
                    else:
                        gained = deferred_cells[at]
                        deferred_cells[at] = (pid, PARA_BREAK.join([gained[1], *paragraphs]), words)
                    continue
                anchor = last.get(form_anchor, cell_row if cell_row is not None else form_anchor)
                # the template is read from the form's own structure, before any paragraph
                # was added after this anchor: a group inserted beneath it must not make the
                # next paragraph at its own level take the deeper style
                if form_anchor not in templates:
                    templates[form_anchor] = _body_template_for(form_anchor, doc.styles.element)
                    empty_titles[form_anchor] = (
                        templates[form_anchor] is not None
                        and templates[form_anchor] is not form_anchor
                        and _level_of(templates[form_anchor]) is None
                        and not _is_heading(templates[form_anchor])
                        and _nothing_under(form_anchor)
                    )
                body_template = templates[form_anchor]
                new_heading = headed.get(form_anchor)
                sibling_body: tuple | None = None
                bodies = 0
                for k, (text, heading) in enumerate(zip(paragraphs, headings, strict=False)):
                    beneath_of = heading.get("beneath") if isinstance(heading, dict) else None
                    if isinstance(heading, dict) and beneath_of is None:
                        # restyled by the drafter as the form paragraph reading so
                        model = _find_near(
                            body,
                            heading.get("like", ""),
                            form_anchor,
                            struck=True,
                            prefer_struck=bool(heading.get("struck")),
                            struck_here=struck_at.get(_key(heading.get("like", "")), []),
                        )
                        if model is None or model.tag != qn("w:p"):
                            missing.append(f"style like: {heading.get('like', '')[:80]}")
                        anchor = _new_paragraph_after(anchor, text, like=model, restyled=True)
                    elif heading == "beneath" and empty_titles[form_anchor]:
                        # a title the form keeps empty: what is written under it is body text
                        anchor = _new_paragraph_after(anchor, text, like=body_template)
                    elif heading == "beneath" or beneath_of is not None:
                        # a sub-paragraph of the anchor: the (i), (ii) limbs of its (a); under
                        # an unnumbered body paragraph, the first list level of its sub-clause.
                        # The anchor is the paragraph the drafter put the limbs beneath, which
                        # may be one this piece added (a lead-in ending "in respect of the
                        # Services:"), not the form paragraph the group follows
                        parent = (
                            written.get(_key(beneath_of), form_anchor)
                            if beneath_of is not None
                            else form_anchor
                        )
                        level, carrier = _enclosing_level(parent)
                        anchor = _new_paragraph_after(
                            anchor, text, like=carrier if carrier is not None else parent
                        )
                        if level is not None:
                            _set_level(anchor, level + 1)
                        else:
                            _set_level_below(anchor, parent, numbering)
                    elif heading:
                        like = (
                            _find_near(body, heading, form_anchor)
                            if isinstance(heading, str)
                            else None
                        )
                        anchor = _new_paragraph_after(
                            anchor, text, heading=True, like=like, styles=doc.styles.element
                        )
                        new_heading = anchor
                        sibling_body = (
                            _unnumbered_body(like, doc.styles.element)
                            if like is not None
                            else None
                        )
                        # the paragraphs under this heading in this group, up to the next title
                        under = len(list(takewhile(lambda h: h is False, headings[k + 1 :])))
                        if sibling_body is not None and sibling_body[1] is None and under != 1:
                            sibling_body = None  # several paragraphs, no limb to lay them out by
                        bodies = 0
                    elif new_heading is not None and sibling_body is not None:
                        # the sibling sub-clause opens with unnumbered body text (a
                        # lead-in paragraph): so does this one, a lone paragraph never
                        # a lone (a), any further paragraph taking the sibling's first list limb
                        lead, limb = sibling_body
                        model = lead if bodies == 0 or limb is None else limb
                        anchor = _new_paragraph_after(anchor, text, like=model)
                        bodies += 1
                    elif new_heading is not None and _level_of(new_heading) is not None:
                        # a paragraph of the new sub-clause: one numbering level below its heading
                        new = _new_paragraph_after(anchor, text, like=new_heading)
                        _set_level(new, (_level_of(new_heading) or 1) + 1)
                        follower = _follower_of(new_heading, doc.styles.element)
                        if follower is not None and _keeps_with_next(new, doc.styles.element):
                            # that level is a title too (an appendix's Part A): the words are
                            # body text, in the style the heading's style names to follow it
                            new.getparent().remove(new)
                            new = _new_paragraph_after(anchor, text, like=follower)
                        anchor = new
                    else:
                        anchor = _new_paragraph_after(anchor, text, like=body_template)
                    if moved_from.get(_key(text)):
                        # words the piece moved unchanged: the form's paragraph content, its
                        # fields and bookmarks with it, never the words retyped without them
                        source = moved_from[_key(text)].pop(0)
                        for child in list(anchor):
                            if child.tag != qn("w:pPr"):
                                anchor.remove(child)
                        anchor.extend(
                            copy.deepcopy(c) for c in source if c.tag != qn("w:pPr")
                        )
                    written[_key(text)] = anchor
                last[form_anchor] = anchor
                if new_heading is not None and _keeps_with_next(new_heading, doc.styles.element):
                    # a title, as Word marks one: the paragraphs added next after the same form
                    # paragraph, in a later group (each move is one), are numbered under it
                    headed[form_anchor] = new_heading
            elif kind == "r":
                before, cells = rows_inserted[n]
                form_anchor = anchor_for(
                    before, row_ranks[n] if n < len(row_ranks) else -1, _nth(contexts.get("r"), n)
                )
                if form_anchor is not None and form_anchor.tag == qn("w:p"):
                    # a table of one column: its row is found by the paragraph of its one cell
                    form_anchor = next(form_anchor.iterancestors(qn("w:tr")), form_anchor)
                anchor = last.get(form_anchor, form_anchor)
                if anchor is None:
                    missing.append(f"row after: {before[:80]}")
                    continue
                if anchor.tag == qn("w:p"):
                    table = anchor.getprevious()
                    if anchor is not form_anchor and table is not None and table.tag == qn("w:tbl"):
                        # after a table this piece added: a row added after its last row
                        anchor = table.findall(qn("w:tr"))[-1]
                    elif _key(_element_text(anchor)) == _key(before):
                        # no table to join: the words are written, as a paragraph after it
                        text = " | ".join(c for c in cells if c.strip())
                        last[form_anchor] = _new_paragraph_after(anchor, text)
                        continue
                if anchor.tag != qn("w:tr"):
                    missing.append(f"row after: {before[:80]}")
                    continue
                new = copy.deepcopy(_row_model(anchor, len(cells)))
                _fresh_ids(new)
                anchor.addnext(new)
                _set_row(new, " | ".join(cells))
                last[form_anchor] = new
            else:
                before, header, table_rows = tables_inserted[n]
                form_anchor = anchor_for(
                    before,
                    table_ranks[n] if n < len(table_ranks) else -1,
                    _nth(contexts.get("t"), n),
                )
                anchor = last.get(form_anchor, form_anchor)
                if anchor is None:
                    missing.append(f"table after: {before[:80]}")
                    continue
                last[form_anchor] = _new_table_after(anchor, header, table_rows)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), missing


def _row_model(anchor, cells: int):
    """The row an added row is copied from: the one it follows, unless that row has fewer
    cells than the added row (a heading row spanning the table's columns), then the nearest
    row of the table after it, else before it, that has a cell for each."""
    if len(anchor.findall(qn("w:tc"))) >= cells:
        return anchor
    for rows in (anchor.itersiblings(qn("w:tr")), anchor.itersiblings(qn("w:tr"), preceding=True)):
        for tr in rows:
            if len(tr.findall(qn("w:tc"))) >= cells:
                return tr
    return anchor


W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


def _para_id(p) -> str:
    return p.get(f"{{{W14}}}paraId") or ""


def _as_deleted(run) -> None:
    """A run going inside w:del: its text becomes w:delText and its field code
    w:delInstrText. Word opens a file only if every field code in a deleted region is
    written the deleted way (ECMA-376 17.16.13); a live w:instrText there is "unreadable
    content"."""
    for t in run.iter(qn("w:t")):
        t.tag = qn("w:delText")
    for t in run.iter(qn("w:instrText")):
        t.tag = qn("w:delInstrText")


def _rewrite_tracked(p, final: str, original: str, n: int, author: str, stamp: str) -> int:
    """One paragraph element rewritten as one tracked change: its runs struck, the new words
    inserted, fields and footnote markers kept as set_paragraph_text keeps them. Returns the
    next free revision id."""
    fresh = copy.deepcopy(p)
    set_paragraph_text(fresh, final, original)
    n += 1
    dele = p.makeelement(
        qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
    )
    for child in list(p):
        if child.tag == qn("w:ins"):
            # words the compare wrote into it (a row it paired with another): never the form's
            p.remove(child)
        elif child.tag in (qn("w:r"), qn("w:hyperlink"), qn("w:fldSimple")):
            p.remove(child)
            _as_deleted(child)
            dele.append(child)
    if len(dele):
        p.append(dele)
    n += 1
    ins = p.makeelement(
        qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
    )
    for r in fresh.findall(qn("w:r")):
        ins.append(copy.deepcopy(r))
    if len(ins):
        p.append(ins)
    return n


def _header_paragraphs(doc) -> dict[str, object]:
    """Every paragraph in the document's headers and footers, by id."""
    out: dict[str, object] = {}
    for section in doc.sections:
        for part in (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ):
            if part.is_linked_to_previous:
                continue
            for p in part._element.iter(qn("w:p")):
                if _para_id(p):
                    out[_para_id(p)] = p
    return out


def rewrite_paragraphs(
    docx_bytes: bytes, changes: list[tuple[str, str, str]], author: str = "Drafting engine"
) -> bytes:
    """Rewrite each paragraph given by id as one tracked change. Done after the compare, on
    paragraphs the compare saw unchanged."""
    if not changes:
        return docx_bytes
    doc = Document(io.BytesIO(docx_bytes))
    # a deferred cell may sit in a header's table (a cover label) as well as in the body
    by_id = {_para_id(p): p for p in doc.element.body.iter(qn("w:p")) if _para_id(p)}
    by_id.update(_header_paragraphs(doc))
    n = 93000
    stamp = "2026-01-01T00:00:00Z"
    for pid, final, original in changes:
        p = by_id.get(pid)
        if p is None:
            continue
        first, *more = final.split(PARA_BREAK)
        like = copy.deepcopy(p)
        if first.strip() != original.strip():  # a cell paragraph that only gained lines stands
            n = _rewrite_tracked(p, first.strip(), original, n, author, stamp)
        after = p
        for piece in more:
            # a paragraph the cell gained: inserted whole, its mark tracked as inserted
            new = copy.deepcopy(like)
            _fresh_ids(new)
            set_paragraph_text(new, piece.strip())
            n += 1
            ins = new.makeelement(
                qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            for r in new.findall(qn("w:r")):
                new.remove(r)
                ins.append(r)
            new.append(ins)
            ppr = new.find(qn("w:pPr"))
            if ppr is None:
                ppr = new.makeelement(qn("w:pPr"), {})
                new.insert(0, ppr)
            rpr = ppr.find(qn("w:rPr"))
            if rpr is None:
                rpr = ppr.makeelement(qn("w:rPr"), {})
                ppr.append(rpr)
            n += 1
            mark = rpr.makeelement(
                qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            rpr.insert(0, mark)
            after.addnext(new)
            after = new
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def place_deferred_cells(
    redline_bytes: bytes, plain_bytes: bytes, changes: list[tuple[str, str, str]]
) -> list[tuple[str, str, str]]:
    """Each deferred cell paragraph re-addressed to the paragraph that holds its place in the
    compared document. The compare pairs a table's changed rows by position, so a row added
    among them hands each form row (and its paragraph ids) the words of the drafted row after
    it. The row is found by what it reads: the compared document, accepted, reads as the
    plain one, so the n-th row of the plain document reading as the cell's row is the n-th
    live row of the compared document reading so; the cell and paragraph are at the same
    place in it. Run on the compare's output, before any row is struck."""
    if not changes:
        return changes

    def rows_of(doc) -> list:
        return [
            tr
            for tr in doc.element.body.iter(qn("w:tr"))
            if tr.find(f"{qn('w:trPr')}/{qn('w:del')}") is None
        ]

    def reads(tr) -> str:
        return _key("".join(_accepted_text(p) for p in tr.iter(qn("w:p"))))

    def live_paragraphs(tc) -> list:
        return [
            p
            for p in tc.findall(qn("w:p"))
            if p.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}") is None
        ]

    plain = Document(io.BytesIO(plain_bytes))
    by_id = {_para_id(p): p for p in plain.element.body.iter(qn("w:p")) if _para_id(p)}
    plain_rows = rows_of(plain)
    plain_reads = [reads(tr) for tr in plain_rows]
    compared: dict[str, list] = {}
    for tr in rows_of(Document(io.BytesIO(redline_bytes))):
        compared.setdefault(reads(tr), []).append(tr)
    out = []
    for pid, final, original in changes:
        q = by_id.get(pid)
        tc = q.getparent() if q is not None else None
        tr = tc.getparent() if tc is not None and tc.tag == qn("w:tc") else None
        if tr is None or tr not in plain_rows:
            out.append((pid, final, original))
            continue
        i = plain_rows.index(tr)
        same = compared.get(plain_reads[i], [])
        nth = plain_reads[:i].count(plain_reads[i])
        ci, pi = tr.findall(qn("w:tc")).index(tc), live_paragraphs(tc).index(q)
        target = None
        if nth < len(same):
            cells = same[nth].findall(qn("w:tc"))
            paras = live_paragraphs(cells[ci]) if ci < len(cells) else []
            target = paras[pi] if pi < len(paras) and _para_id(paras[pi]) else None
        out.append((_para_id(target) if target is not None else pid, final, original))
    return out


def _accepted_text(p) -> str:
    """A paragraph's text with every tracked deletion dropped and every insertion kept."""
    parts = []
    for t in p.iter(qn("w:t")):
        if not any(a.tag == qn("w:del") for a in t.iterancestors(qn("w:del"), qn("w:p"))):
            parts.append(t.text or "")
    return "".join(parts)


def unpair_inserts(
    redline_bytes: bytes, original_bytes: bytes, plain_bytes: bytes, author: str = "Drafting engine"
) -> bytes:
    """A paragraph the drafter added, paired by the compare with the form paragraph beside it.

    Where a form paragraph changed and a new one was written next to it, the compare can pair the
    new paragraph's words with the form's and mark the form's own, as finally worded, as wholly
    inserted: the lawyer reads a clause rewritten that was only amended, and a new one that is
    the old. Where the form's words read closer to the inserted neighbour than to what they were
    paired with, the two are put back: the new paragraph inserted whole, and the form paragraph
    marked up word by word into its final words. Paragraphs are matched by their ids and words."""
    originals: dict[str, object] = {}
    for q in Document(io.BytesIO(original_bytes)).element.body.iter(qn("w:p")):
        originals.setdefault(_para_id(q) or "", q)
        originals.setdefault("words:" + _key(_element_text(q)), q)
    plains: dict[str, object] = {}
    for q in Document(io.BytesIO(plain_bytes)).element.body.iter(qn("w:p")):
        plains.setdefault(_key(_element_text(q)), q)
    doc = Document(io.BytesIO(redline_bytes))
    paras = list(doc.element.body.iterchildren(qn("w:p")))
    n = 95000

    def form_words(p) -> str:
        """What the compare kept of the form in `p`: its deletions read, its insertions not."""
        return "".join(
            t.text or ""
            for t in p.iter(qn("w:t"), qn("w:delText"))
            if not any(a.tag == qn("w:ins") for a in t.iterancestors(qn("w:ins"), qn("w:p")))
        )
    stamp = "2026-01-01T00:00:00Z"

    def alike(a: str, b: str) -> float:
        return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()

    done: set[int] = set()
    for a, b in zip(paras, paras[1:], strict=False):
        for mixed, added in ((a, b), (b, a)):
            if id(mixed) in done or id(added) in done or not _is_inserted_paragraph(added):
                continue
            if mixed.find(qn("w:del")) is None or mixed.find(qn("w:ins")) is None:
                continue
            old = originals.get(_para_id(mixed)) if _para_id(mixed) else None
            old = old if old is not None else originals.get("words:" + _key(form_words(mixed)))
            if old is None:
                continue
            was, took, other = _element_text(old), _accepted_text(mixed), _accepted_text(added)
            if alike(was, other) <= alike(was, took):
                continue
            took_src, other_src = plains.get(_key(took)), plains.get(_key(other))
            if took_src is None or other_src is None:
                continue
            for el, src in ((mixed, took_src), (added, other_src)):
                for child in list(el):
                    el.remove(child)
                ppr = src.find(qn("w:pPr"))
                if ppr is not None:
                    el.append(copy.deepcopy(ppr))
            # the new paragraph, inserted whole, its mark inserted too
            n += 1
            ins = mixed.makeelement(
                qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            for r in took_src.findall(qn("w:r")):
                ins.append(copy.deepcopy(r))
            mixed.append(ins)
            ppr = mixed.find(qn("w:pPr"))
            if ppr is None:
                ppr = mixed.makeelement(qn("w:pPr"), {})
                mixed.insert(0, ppr)
            rpr = ppr.find(qn("w:rPr"))
            if rpr is None:
                rpr = ppr.makeelement(qn("w:rPr"), {})
                ppr.append(rpr)
            n += 1
            rpr.append(
                rpr.makeelement(
                    qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
                )
            )
            # the form paragraph, marked up into its final words
            n = _track_word_changes(added, old, other_src, n, author, stamp)
            mixed_id, added_id = _para_id(took_src), _para_id(old)
            if mixed_id:
                mixed.set(f"{{{W14}}}paraId", mixed_id)
            if added_id:
                added.set(f"{{{W14}}}paraId", added_id)
            done |= {id(mixed), id(added)}
            break
    if not done:
        return redline_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _is_inserted_paragraph(p) -> bool:
    """A paragraph the compare engine marked as existing only in the revised document: its
    runs all sit in w:ins and its paragraph mark is inserted."""
    if p.find(qn("w:del")) is not None:
        return False
    runs = list(p.iter(qn("w:r")))
    if not runs or not all(r.getparent().tag == qn("w:ins") for r in runs):
        return False
    ppr = p.find(qn("w:pPr"))
    rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
    return rpr is not None and rpr.find(qn("w:ins")) is not None


def restore_insert_order(redlined_bytes: bytes, plain_bytes: bytes) -> bytes:
    """Put every inserted body paragraph back where the drafter placed it.

    The comparison engine pairs the form's paragraphs with the revised document's and writes
    a paragraph that exists only in the revised document after the paired paragraphs around
    it, not at its own position: inserted limbs (a), (b), (c) come out (a), (c), (b).
    The plain document holds the drafter's order, so each inserted paragraph is moved to
    follow the paragraph that precedes it there."""
    red = Document(io.BytesIO(redlined_bytes))
    plain = Document(io.BytesIO(plain_bytes))

    def key_of(el, text_of) -> str:
        # a table is a member of the sequence too, known by its first words: a paragraph the
        # drafter wrote after a table must stay after it
        if el.tag == qn("w:tbl"):
            first = next(el.iter(qn("w:p")), None)
            return "tbl:" + _key(text_of(first)) if first is not None else "tbl:"
        return _key(text_of(el))

    plain_texts = [
        key_of(el, _element_text) for el in plain.element.body.iterchildren(qn("w:p"), qn("w:tbl"))
    ]
    plain_texts = [t for t in plain_texts if t]
    pool: dict[str, list] = {}
    red_paras = list(red.element.body.iterchildren(qn("w:p"), qn("w:tbl")))
    for p in red_paras:
        k = key_of(p, _accepted_text)
        if k:
            pool.setdefault(k, []).append(p)
    mapped: list = []  # the redline element for each plain paragraph, in plain order
    for k in plain_texts:
        bucket = pool.get(k)
        mapped.append(bucket.pop(0) if bucket else None)
    moved = False
    for i in range(1, len(mapped)):
        el, before = mapped[i], mapped[i - 1]
        if el is None or before is None or el.tag != qn("w:p") or not _is_inserted_paragraph(el):
            continue
        prev = el.getprevious()
        while prev is not None and prev.tag == qn("w:p") and not _key(_accepted_text(prev)):
            prev = prev.getprevious()  # a struck paragraph between them does not count
        if prev is before:
            continue
        before.addnext(el)
        moved = True
    if not moved:
        return redlined_bytes
    out = io.BytesIO()
    red.save(out)
    return out.getvalue()


def carry_header_changes(
    redline_bytes: bytes, form_bytes: bytes, plain_bytes: bytes, author: str = "Drafting engine"
) -> bytes:
    """The comparison engine compares bodies and carries the form's headers and footers over
    unchanged, so a drafter's edit to a header paragraph (a cover label) is
    lost. Each header paragraph the plain document changed is rewritten in the redline as one
    tracked change, matched by id."""
    form = _header_paragraphs(Document(io.BytesIO(form_bytes)))
    plain = _header_paragraphs(Document(io.BytesIO(plain_bytes)))
    doc = Document(io.BytesIO(redline_bytes))
    n = 95000
    stamp = "2026-01-01T00:00:00Z"
    changed = False
    for pid, p in _header_paragraphs(doc).items():
        before, after = form.get(pid), plain.get(pid)
        if before is None or after is None:
            continue
        original, final = _element_text(before), _element_text(after)
        if original == final:
            continue
        n = _rewrite_tracked(p, final, original, n, author, stamp)
        changed = True
    if not changed:
        return redline_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


_MARK_FIRST = ("ins", "del", "moveFrom", "moveTo")


def order_paragraph_marks(docx_bytes: bytes) -> bytes:
    """Two rules of a paragraph mark's tracked change (w:ins or w:del inside w:pPr/w:rPr): it
    comes before the run properties, and a paragraph that keeps live words has no struck mark.
    The comparison engine breaks both, and Word refuses the file or glues paragraphs."""
    doc = Document(io.BytesIO(docx_bytes))
    changed = False
    for ppr in doc.element.body.iter(qn("w:pPr")):
        rpr = ppr.find(qn("w:rPr"))
        if rpr is None:
            continue
        # A paragraph that keeps live words keeps its mark: a struck mark on a paragraph the
        # compare later filled with insertions would, on acceptance, glue it to the next one.
        para = ppr.getparent()
        live = any(r.getparent().tag != qn("w:del") for r in para.iter(qn("w:r")))
        for mark_del in rpr.findall(qn("w:del")):
            if live:
                rpr.remove(mark_del)
                changed = True
        marks = [c for c in rpr if c.tag in {qn(f"w:{n}") for n in _MARK_FIRST}]
        wanted = sorted(marks, key=lambda c: _MARK_FIRST.index(c.tag.split("}")[1]))
        if not marks or list(rpr)[: len(marks)] == wanted:
            continue
        for c in marks:
            rpr.remove(c)
        for i, c in enumerate(wanted):
            rpr.insert(i, c)
        changed = True
    if not changed:
        return docx_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


_REF_FIELD = re.compile(r"\bREF\s+(_Ref\w+)((?:\s+\\[a-zA-Z*]+)*)")
_TRAILING_WORD = re.compile(r"(\s+(?:above|below))\s*$", re.I)


def refresh_reference_fields(docx_bytes: bytes) -> bytes:
    """Write each numeric REF field's saved result as the number of its bookmark,
    counting a heading whose paragraph mark is deleted as gone.

    The bookmark already opens the right paragraph. The saved result is the number
    from before that heading was struck, which is the number Word no longer shows.
    `\\w` and `\\r` take the full number, `\\n` the marker. A bare REF caches the
    heading text and is left as it is.
    """
    doc = Document(io.BytesIO(docx_bytes))
    numbers = accepted_numbers_by_paragraph(docx_bytes, doc)
    anchors: dict[str, int] = {}
    for index, paragraph in enumerate(doc.paragraphs):
        for node in paragraph._p.iter(qn("w:bookmarkStart")):
            name = node.get(qn("w:name"))
            if name and name.startswith("_Ref"):
                anchors.setdefault(name, index)
    changed = False
    # a reference to a paragraph the draft strikes whole has nothing to show once Word
    # refreshes it ("Error! Reference source not found."): it is left as the words it showed
    struck = {
        name
        for node in doc.element.body.iter(qn("w:bookmarkStart"))
        if (name := node.get(qn("w:name")) or "").startswith("_Ref") and _opens_struck(node)
    }
    # and so has a reference to a bookmark the document does not hold, the form's own
    # included: the draft asks Word to refresh every field on opening (a precedent's
    # cross-reference can name a bookmark it never had)
    held = {node.get(qn("w:name")) for node in doc.element.body.iter(qn("w:bookmarkStart"))}
    for element in doc.element.body.iter(qn("w:p")):
        if _rewrite_reference_results(element, anchors, numbers):
            changed = True
        code = "".join(x.text or "" for x in element.iter(qn("w:instrText")))
        gone = struck | {m.group(1) for m in _REF_FIELD.finditer(code) if m.group(1) not in held}
        if gone and _unwrap_references_to(element, gone):
            changed = True
    if not changed:
        return docx_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


_TOC_ENTRY = re.compile(r"TOC\d")
_PAGEREF = re.compile(r"PAGEREF\s+(_Toc\w+)")


def _field_depths(entries: list) -> dict:
    """Each run of the contents entries holding a field character or code, with the depth of
    the field it belongs to: 1 for the contents field itself, 2 for an entry's page number,
    whose shown result is at 2 too. An entry's own words are at none."""
    depth, out = 0, {}
    for e in entries:
        for r in e.iter(qn("w:r")):
            if depth > 1:
                out[r] = depth  # the page number a field shows, inside it
            for node in r:
                if node.tag == qn("w:fldChar"):
                    kind = node.get(qn("w:fldCharType"))
                    if kind == "begin":
                        depth += 1
                    out[r] = depth
                    if kind == "end":
                        depth -= 1
                elif node.tag == qn("w:instrText"):
                    out[r] = depth
    return out


def _entry_parts(entry, depths: dict) -> tuple[str, str]:
    """An entry's shown words before its page number, as (number, title): the number is what
    stands before the entry's first tab, or its first two words ("Schedule 4")."""
    words = []
    for r in entry.iter(qn("w:r")):
        if depths.get(r, 0) > 1 or _inside_deletion(r):
            break  # the page number's field, and what follows
        if depths.get(r) == 1:
            continue
        for node in r:
            if node.tag == qn("w:t"):
                words.append(node.text or "")
            elif node.tag == qn("w:tab"):
                words.append("\t")
    text = "".join(words).rstrip("\t")
    if "\t" in text:
        number, _, title = text.partition("\t")
        return number, title.strip()
    m = re.match(r"(\S+ (?:\d+|[A-Z]))\b\s*(.*)", text)
    return (m.group(1), m.group(2).strip()) if m else ("", text.strip())


def _tabbed(entry, depths: dict) -> bool:
    """Whether a tab parts the entry's number from its title, not only its title from its page."""
    shown = []
    for r in entry.iter(qn("w:r")):
        if depths.get(r, 0) > 1:
            break
        if depths.get(r) != 1 and not _inside_deletion(r):
            shown += ["\t" if n.tag == qn("w:tab") else n.text or "" for n in r if n.tag in (qn("w:t"), qn("w:tab"))]
    return "\t" in "".join(shown).strip("\t")


def _token(number: str) -> str:
    found = re.findall(r"\d+|\b[A-Z]\b", number)
    return found[-1] if found else ""


def refresh_contents(docx_bytes: bytes) -> bytes:
    """The contents list as Word would write it for the contract as it now stands, made before
    the lawyer opens it: an entry whose heading the draft strikes, a second entry for one
    heading, or a line typed with its own dot leaders that names no heading, goes; each entry
    takes its heading's number as Word now prints it, and its words where the draft changed
    them; a heading of a style the contents field lists (its \\t and \\o switches), with no
    entry, is given one, written as the other entries of its style are. A typed label repeating
    the number the heading's style prints is not part of its title. A line typed into the
    contents is left as typed, unless its heading is struck."""
    doc = Document(io.BytesIO(docx_bytes))
    paras = [p._p for p in doc.paragraphs]
    # the contents may sit in a content control, outside the body's own paragraphs
    entries = [p for p in doc.element.body.iter(qn("w:p")) if _TOC_ENTRY.fullmatch(_style_id(p))]
    if not entries:
        return docx_bytes
    numbers = accepted_numbers_by_paragraph(docx_bytes, doc)
    marks: dict[str, int] = {}
    for i, p in enumerate(paras):
        for bm in p.iter(qn("w:bookmarkStart")):
            if (bm.get(qn("w:name")) or "").startswith("_Toc"):
                marks.setdefault(bm.get(qn("w:name")), i)
    depths = _field_depths(entries)

    def words(p, struck: bool) -> str:
        nodes = p.iter(qn("w:t"), qn("w:delText")) if struck else p.iter(qn("w:t"))
        return re.sub(r"\W", "", "".join(n.text or "" for n in nodes if struck or not _inside_deletion(n)).lower())

    headings: dict[str, int] = {}  # by its words, a heading standing before one struck
    for i, p in enumerate(paras):
        key = words(p, True)
        if key and not _TOC_ENTRY.fullmatch(_style_id(p)):
            if key not in headings or (not _accepted_text(paras[headings[key]]).strip() and _accepted_text(p).strip()):
                headings[key] = i

    def target(entry) -> tuple[int | None, bool]:
        """The heading an entry names, and whether by its page number's bookmark: else, for a
        line typed into the contents, the heading reading as it does."""
        code = "".join(x.text or "" for x in entry.iter(qn("w:instrText")) if depths.get(x.getparent()) == 2)
        m = _PAGEREF.search(code)
        if m:
            return marks.get(m.group(1)), True
        return headings.get(words(entry, True)), False

    def live(i: int) -> bool:
        return bool(_accepted_text(paras[i]).strip())

    def title_of(i: int) -> str:
        text = " ".join(_accepted_text(paras[i]).split())
        number = numbers[i].marker if i in numbers else ""
        lead = number.split()[0] if number.split() else ""
        if lead and text.lower().startswith(lead.lower() + " "):
            text = re.sub(r"^\S+\s+\S+\s*", "", text)  # "SCHEDULE 5 Insurance" under "Schedule 4"
        return text

    # the styles the contents field lists: its \t pairs and its \o heading levels
    code = "".join(x.text or "" for e in entries for x in e.iter(qn("w:instrText")) if depths.get(x.getparent()) == 1)
    listed_names: set[str] = set()
    switch = re.search(r'\\t\s+"([^"]*)"', code.replace("&quot;", '"'))
    if switch:
        parts = [x.strip() for x in switch.group(1).split(",")]
        listed_names |= {parts[k] for k in range(0, len(parts) - 1, 2) if parts[k]}
    levels = re.search(r'\\o\s+"(\d+)-(\d+)"', code)
    if levels:
        listed_names |= {f"Heading {n}" for n in range(int(levels.group(1)), int(levels.group(2)) + 1)}
    toc_styles = {st.style_id for st in doc.styles if getattr(st, "name", None) in listed_names}

    kept: list[tuple[object, int, bool]] = []  # entry, heading, named by its bookmark
    carry: list = []  # the contents field's own begin, code and separator, from an entry that goes
    for e in entries:
        t, by_mark = target(e)
        own = [r for r in e.iter(qn("w:r")) if depths.get(r) == 1]
        leader = re.search(r"\.{4,}", "".join(n.text or "" for n in e.iter(qn("w:t"))))
        named = {words(paras[k], True) for _, k, mark in kept if mark}
        typed_leader = leader and not by_mark and (t is None or words(paras[t], True) in named)
        gone = typed_leader or (t is not None and (not live(t) or any(k == t for _, k, _ in kept)))
        if not gone and t is None:
            continue  # a line naming no heading, nothing typed over it: left as the form has it
        if gone:
            ends = [r for r in own if r.find(qn("w:fldChar")) is not None and r.find(qn("w:fldChar")).get(qn("w:fldCharType")) == "end"]
            carry += [r for r in own if r not in ends]
            for r in own:
                r.getparent().remove(r)
            if ends and kept:
                kept[-1][0].append(ends[0])  # type: ignore[attr-defined]
            elif ends:
                carry += ends
            e.getparent().remove(e)
            continue
        if carry:
            ppr = e.find(qn("w:pPr"))
            at = 1 if ppr is not None else 0
            for r in carry:
                e.insert(at, r)
                at += 1
            carry = []
        kept.append((e, t, by_mark))  # type: ignore[arg-type]
    if not kept:
        return docx_bytes
    def fused(e) -> bool:  # a typed contents line run into the entry, its dot leaders typed
        return bool(re.search(r"\.{4,}", "".join(n.text or "" for n in e.iter(qn("w:t")) if depths.get(n.getparent(), 0) == 0)))

    shapes: dict[str, tuple[object, str, bool]] = {}  # heading style -> (an entry, its number, tab after it)
    lowered: set[str] = set()  # words the entries write in lower case: "and", "of"
    for e, t, by_mark in kept:
        if by_mark and not fused(e):
            number, title = _entry_parts(e, depths)
            shapes.setdefault(_style_id(paras[t]), (e, number, _tabbed(e, depths)))
            lowered |= {w for w in title.split() if w.islower()}
    for e, t, by_mark in kept:
        if not by_mark:
            continue  # a typed line stays as the form typed it
        if fused(e) and _style_id(paras[t]) in shapes:
            _, number, tabbed = shapes[_style_id(paras[t])]
            token = _token(number)
            shown = re.sub(rf"\b{re.escape(token)}\b", _token(numbers[t].marker), number, count=1) if token and t in numbers else number
            _rewrite_entry_words(e, depths, shown, tabbed, title_of(t))
            continue
        new_token = _token(numbers[t].marker) if t in numbers else ""
        old_token = _token(number)
        if new_token and old_token and new_token != old_token:
            for node in e.iter(qn("w:t")):
                if old_token in (node.text or "") and depths.get(node.getparent(), 0) == 0:
                    node.text = re.sub(rf"\b{re.escape(old_token)}\b", new_token, node.text, count=1)
                    break
        heading = paras[t]
        if heading.find(f".//{qn('w:ins')}") is not None or heading.find(f".//{qn('w:del')}") is not None:
            _retitle_entry(e, depths, title_of(t))

    def cased(text: str) -> str:
        if text != text.upper() or not lowered:
            return text
        return " ".join(
            w.lower() if k and w.lower() in lowered else w.capitalize() for k, w in enumerate(text.split())
        )

    listed = {t for _, t, _ in kept}
    used = [int(b.get(qn("w:id"))) for b in doc.element.body.iter(qn("w:bookmarkStart")) if (b.get(qn("w:id")) or "").isdigit()]
    next_id = max(used, default=0) + 1
    for i, p in enumerate(paras):
        if i in listed or _style_id(p) not in toc_styles or not live(i) or not title_of(i) or i not in numbers or not shapes:
            continue
        before = [(t, e) for e, t, _ in kept if t < i]
        anchor = max(before, key=lambda x: x[0])[1] if before else None
        model, number, tabbed = shapes.get(_style_id(p)) or shapes[next(iter(shapes))]
        new = copy.deepcopy(model)
        for child in list(new):
            if child.tag != qn("w:pPr"):
                new.remove(child)
        name = f"_TocDR{next_id:06d}"
        start = p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): str(next_id), qn("w:name"): name})
        end = p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): str(next_id)})
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ppr.addnext(start)
        else:
            p.insert(0, start)
        p.append(end)
        next_id += 1
        token = _token(number)
        shown = re.sub(rf"\b{re.escape(token)}\b", _token(numbers[i].marker), number, count=1) if token else numbers[i].marker
        _append_entry_runs(new, shown, tabbed, cased(title_of(i)), name)
        if anchor is not None:
            anchor.addnext(new)  # type: ignore[attr-defined]
        else:
            kept[0][0].addprevious(new)  # type: ignore[attr-defined]
        kept.append((new, i, True))
        listed.add(i)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _rewrite_entry_words(entry, depths: dict, number: str, tabbed: bool, title: str) -> None:
    """An entry's own words, all of them, written again before its page number: its number,
    a tab or a space, its title, and the tab before the page."""
    page = next((r for r in entry.iter(qn("w:r")) if depths.get(r, 0) > 1), None)
    if page is None:
        return
    for r in list(entry.iter(qn("w:r"))):
        if depths.get(r, 0) == 0 and not _inside_deletion(r) and any(n.tag in (qn("w:t"), qn("w:tab")) for n in r):
            r.getparent().remove(r)
    held = entry.makeelement(qn("w:p"), {})
    _append_entry_runs(held, number, tabbed, title, "")
    for r in [r for r in held if r.find(qn("w:fldChar")) is None and r.find(qn("w:instrText")) is None]:
        page.addprevious(r)


def _retitle_entry(entry, depths: dict, title: str) -> None:
    """An entry's title, the words between its number and its page number, rewritten."""
    shown = []  # the entry's own words and tabs, before its page number
    for r in entry.iter(qn("w:r")):
        if depths.get(r, 0) > 1:
            break
        if depths.get(r) == 1 or _inside_deletion(r):
            continue
        shown += [n for n in r if n.tag in (qn("w:t"), qn("w:tab"))]
    tab = next((k for k, n in enumerate(shown) if n.tag == qn("w:tab")), None)
    words = [n for n in shown[tab + 1 if tab is not None else 0 :] if n.tag == qn("w:t")]
    if not words:
        return
    if tab is None:  # "Schedule 4" and its title in the same words: the number stays
        m = re.match(r"(\S+ (?:\d+|[A-Z]))\b", "".join(n.text or "" for n in words))
        title = f"{m.group(1)} {title}" if m else title
    words[0].text = title
    for n in words[1:]:
        n.text = ""


def _append_entry_runs(entry, number: str, tabbed: bool, title: str, mark: str) -> None:
    """A contents entry's runs: its number, a tab or a space, its title, a tab and the page
    number as a field to the heading's bookmark (empty until Word counts the pages)."""

    def run(*nodes) -> None:
        r = entry.makeelement(qn("w:r"), {})
        for n in nodes:
            r.append(n)
        entry.append(r)

    def text(s: str):
        t = entry.makeelement(qn("w:t"), {})
        t.text = s
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return t

    def node(tag: str, **attrs):
        return entry.makeelement(qn(tag), {qn(k): v for k, v in attrs.items()})

    if tabbed:
        run(text(number))
        run(node("w:tab"))
        run(text(title))
    else:
        run(text(f"{number} {title}"))
    run(node("w:tab"))
    run(node("w:fldChar", **{"w:fldCharType": "begin"}))
    instr = node("w:instrText")
    instr.text = f" PAGEREF {mark} \\h "
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run(instr)
    run(node("w:fldChar", **{"w:fldCharType": "separate"}))
    run(node("w:fldChar", **{"w:fldCharType": "end"}))


def _opens_struck(bookmark) -> bool:
    """Whether a bookmark opens a paragraph the draft strikes whole: its paragraph mark is
    deleted and none of its words is left."""
    host = next(bookmark.iterancestors(qn("w:p")), None)
    if host is None:
        host = bookmark.getnext()  # a bookmark opened just before its paragraph
    if host is None or host.tag != qn("w:p"):
        return False
    mark = host.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}")
    return mark is not None and not _accepted_text(host).strip()


def _unwrap_references_to(paragraph, struck: set[str]) -> bool:
    """Each live REF field of the paragraph to a bookmark in `struck` made plain words: its
    field characters and code go, the result it showed stays."""
    changed = False
    begin = None
    parts: list = []
    for node in list(paragraph.iter(qn("w:fldChar"), qn("w:instrText"))):
        kind = node.get(qn("w:fldCharType")) if node.tag == qn("w:fldChar") else "code"
        if kind == "begin":
            begin, parts = node, [node]
        elif begin is not None and kind in ("code", "separate"):
            parts.append(node)
        elif begin is not None and kind == "end":
            parts.append(node)
            code = "".join(x.text or "" for x in parts if x.tag == qn("w:instrText"))
            match = _REF_FIELD.search(code)
            if match and match.group(1) in struck and not _inside_deletion(begin):
                for x in parts:
                    x.getparent().remove(x)
                changed = True
            begin = None
    return changed


def clause_numbers(docx_bytes: bytes, render: str) -> dict[tuple[str, int], str]:
    """The number Word shows on each numbered paragraph, as a lawyer cites it ("12.1",
    "Schedule 3, 1.3(c)"), keyed by the line of `render`, the document's render, that shows the
    paragraph: the line's key and which occurrence of that key among the render's lines it is.
    The body's paragraphs and the render's lines are aligned in order, as the two list the same
    paragraphs in the same order; a paragraph with line breaks is its first line."""
    doc = Document(io.BytesIO(docx_bytes))
    numbers = accepted_numbers_by_paragraph(docx_bytes, doc)
    paragraphs = []
    for i, paragraph in enumerate(doc.paragraphs):
        lines = _lines_of(paragraph._p)
        key = _key(lines[0]) if len(lines) > 1 else _key(_element_text(paragraph._p))
        if key:
            paragraphs.append((key, i))
    keys: list[str] = []
    ranks: list[int] = []
    seen: dict[str, int] = {}
    for line in render.split("\n"):
        key = _key(document_text(line))
        ranks.append(seen.get(key, 0))
        seen[key] = ranks[-1] + 1
        keys.append(key)
    matcher = difflib.SequenceMatcher(None, [k for k, _ in paragraphs], keys, autojunk=False)
    out: dict[tuple[str, int], str] = {}
    for a, b, size in matcher.get_matching_blocks():
        for n in range(size):
            key, i = paragraphs[a + n]
            if i in numbers:
                out[(key, ranks[b + n])] = numbers[i].cited
    return out


def _rewrite_reference_results(
    paragraph, anchors: dict[str, int], numbers: dict[int, ClauseNumber]
) -> bool:
    changed = False
    instruction = ""
    result_nodes: list = []
    state: str | None = None
    for node in paragraph.iter():
        tag = node.tag
        if tag == qn("w:fldChar"):
            state = node.get(qn("w:fldCharType"))
            if state == "begin":
                instruction = ""
                result_nodes = []
            elif state == "end":
                if _replace_reference_result(instruction, result_nodes, anchors, numbers):
                    changed = True
                state = None
        elif tag == qn("w:instrText") and state == "begin":
            instruction += node.text or ""
        elif tag == qn("w:t") and state == "separate" and not _inside_deletion(node):
            result_nodes.append(node)
    return changed


def _replace_reference_result(
    instruction: str,
    result_nodes: list,
    anchors: dict[str, int],
    numbers: dict[int, ClauseNumber],
) -> bool:
    if "PAGEREF" in instruction or not result_nodes:
        return False
    match = _REF_FIELD.search(instruction)
    if not match:
        return False
    switch = next((s for s in re.findall(r"\\([a-z])", match.group(2)) if s in "nrw"), None)
    if switch is None:
        return False
    anchor = match.group(1)
    index = anchors.get(anchor)
    if index is None:
        return False
    number = numbers.get(index)
    if number is None:
        return False
    old = "".join(node.text or "" for node in result_nodes).replace("\xa0", " ")
    suffix = _TRAILING_WORD.search(old)
    cached = old[: suffix.start()] if suffix else old
    new = _same_shape_as(cached, number, switch)
    if not new:
        return False
    rendered = new + (suffix.group(1) if suffix else "")
    if old.strip() == rendered.strip():
        return False
    result_nodes[0].text = rendered
    for extra in result_nodes[1:]:
        extra.text = ""
    return True


def _number_shape(text: str) -> str:
    """ "(d)", "14.1(k)" and "11" by shape: letters in brackets, digits, the rest kept."""
    shape = re.sub(r"\s+", "", text)
    shape = re.sub(r"\((?:[ivx]+|[IVX]+)\)", "(r)", shape)
    shape = re.sub(r"\([A-Za-z]+\)", "(a)", shape)
    return re.sub(r"\d+", "#", shape)


def _same_shape_as(cached: str, number: ClauseNumber, switch: str) -> str | None:
    """The rendering of the new number that has the shape of the result Word cached. A
    `\\r` field inside Clause 5.4 shows "(d)", not "5.4(d)"; a paragraph reference inside a
    schedule shows "11", not "Schedule 4Part A11". Word's own cached result is the only
    record of which shape the field shows, so the new number is cut to it: the full number,
    its marker, or the tail of the full number from any component. None where no rendering
    has that shape: the cached result stays, the bookmark still opens the right paragraph."""
    want = _number_shape(cached)
    # the number within its own Clause or Schedule comes first: a `\\r` field in Schedule 4
    # shows "6.1(a)", and the tail of the full number "Schedule 46.1(a)" would be "46.1(a)"
    local = number.local or ""
    candidates = [number.marker] if switch == "n" else [local, number.full, number.marker]
    for whole in (local, number.full or ""):
        for i, ch in enumerate(whole):
            if (ch.isdigit() or ch == "(") and (i == 0 or not whole[i - 1].isdigit()):
                candidates.append(whole[i:])
    for candidate in candidates:
        if candidate and _number_shape(candidate) == want:
            return candidate
    return None


def _inside_deletion(node) -> bool:
    parent = node.getparent()
    while parent is not None:
        if parent.tag == qn("w:del"):
            return True
        parent = parent.getparent()
    return False


def unique_bookmarks(docx_bytes: bytes) -> bytes:
    """A bookmark name appears once. A paragraph copied with its bookmarks carries a second
    start and end under the same name, which Word treats as unreadable content; the later
    pair goes."""
    doc = Document(io.BytesIO(docx_bytes))
    seen: set[str] = set()
    for start in list(doc.element.body.iter(qn("w:bookmarkStart"))):
        name = start.get(qn("w:name")) or ""
        if name not in seen:
            seen.add(name)
            continue
        bid = start.get(qn("w:id"))
        for end in list(doc.element.body.iter(qn("w:bookmarkEnd"))):
            if end.get(qn("w:id")) == bid:
                end.getparent().remove(end)
        start.getparent().remove(start)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def strike_paragraphs(docx_bytes: bytes, ids: list[str], author: str = "Drafting engine") -> bytes:
    """Mark every body paragraph with one of the ids given as a tracked deletion: each run and
    the paragraph mark. Done after the compare, on paragraphs the compare saw unchanged."""
    wanted = set(ids)
    if not wanted:
        return docx_bytes
    doc = Document(io.BytesIO(docx_bytes))
    n = 92000
    stamp = "2026-01-01T00:00:00Z"
    # a struck paragraph may sit in a header or footer (a cover label)
    paragraphs = [*doc.element.body.iter(qn("w:p")), *_header_paragraphs(doc).values()]
    for p in paragraphs:
        if _para_id(p) not in wanted:
            continue
        runs = list(p.iter(qn("w:r")))
        if runs and all(r.getparent().tag == qn("w:ins") for r in runs):
            # the compare's inserted twin, carrying the struck paragraph's id: not the form's
            continue
        if any(r.getparent().tag in (qn("w:ins"), qn("w:del")) for r in runs):
            # the compare paired the form's paragraph with drafted words (a definition the
            # drafter moved and changed) and put the copy left to be struck in again, inserted
            # whole: that copy goes, and the pairing stands as the one tracked change
            twin = _inserted_twin(paragraphs, p)
            if twin is not None:
                twin.getparent().remove(twin)
                continue
        for r in runs:
            parent = r.getparent()
            if parent.tag == qn("w:del"):
                continue
            _as_deleted(r)
            n += 1
            wrap = parent.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            parent.replace(r, wrap)
            wrap.append(r)
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = p.makeelement(qn("w:pPr"), {})
            p.insert(0, ppr)
        rpr = ppr.find(qn("w:rPr"))
        if rpr is None:
            rpr = ppr.makeelement(qn("w:rPr"), {})
            ppr.append(rpr)
        n += 1
        rpr.insert(
            0,
            rpr.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            ),
        )
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _inserted_twin(paragraphs: list, p):
    """The paragraph inserted whole that reads as `p` read in the form, nearest to it."""
    form_words = _key(
        "".join(
            t.text or ""
            for t in p.iter(qn("w:t"), qn("w:delText"))
            if t.tag == qn("w:delText") or next(t.iterancestors(qn("w:ins")), None) is None
        )
    )
    at = paragraphs.index(p)
    twins = [
        (abs(i - at), i)
        for i, q in enumerate(paragraphs)
        if q is not p
        and q.getparent() is not None  # a twin already taken out
        and _is_inserted_paragraph(q)
        and _key(_accepted_text(q)) == form_words
    ]
    return paragraphs[min(twins)[1]] if form_words and twins else None


def strike_rows(docx_bytes: bytes, rows: list[list[str]], author: str = "Drafting engine") -> bytes:
    """Mark every table row holding one of the paragraph ids given as a tracked deletion: the
    row's mark and every run in it. Done after the compare, on rows the compare saw unchanged,
    so the tables it compared stayed aligned."""
    wanted = {pid for row in rows for pid in row if pid}
    if not wanted:
        return docx_bytes
    doc = Document(io.BytesIO(docx_bytes))
    n = 90000
    stamp = "2026-01-01T00:00:00Z"
    for tr in list(doc.element.body.iter(qn("w:tr"))):
        if not any(_para_id(p) in wanted for p in tr.iter(qn("w:p"))):
            continue
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = tr.makeelement(qn("w:trPr"), {})
            first = tr.find(qn("w:tblPrEx"))
            if first is not None:
                first.addnext(trpr)
            else:
                tr.insert(0, trpr)
        n += 1
        trpr.append(
            trpr.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
        )
        for r in list(tr.iter(qn("w:r"))):
            parent = r.getparent()
            if parent.tag == qn("w:del"):
                continue
            _as_deleted(r)
            n += 1
            wrap = parent.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            parent.replace(r, wrap)
            wrap.append(r)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _field_wrappers(p) -> list[tuple[str, str]]:
    """Each field character of a paragraph as (kind, wrapper tag), in document order, with
    instruction runs as ("instr", wrapper)."""
    out = []
    for r in p.iter(qn("w:r")):
        wrapper = r.getparent().tag
        fld = r.find(qn("w:fldChar"))
        if fld is not None:
            out.append((fld.get(qn("w:fldCharType")) or "", wrapper))
        elif (it := r.find(qn("w:instrText"))) is not None and (it.text or "").strip():
            # a live field code belongs outside a deletion
            out.append(("instr" if wrapper != qn("w:del") else "instr-wrong", wrapper))
        elif (it := r.find(qn("w:delInstrText"))) is not None and (it.text or "").strip():
            # a deleted field code belongs inside one
            out.append(("instr" if wrapper == qn("w:del") else "instr-wrong", wrapper))
        elif wrapper in (qn("w:del"), qn("w:ins")):
            # text tracked as changed: inside a field's result, the compare edited the result
            out.append(("text", wrapper))
    return out


def _fields_broken(p) -> bool:
    """A field the comparison engine split: its begin, instruction, separate and end no longer
    sit in one wrapper, or the sequence is incomplete."""
    seq = _field_wrappers(p)
    if not seq:
        return False
    depth, wrapper, instr = 0, "", ""
    for kind, w in seq:
        if kind == "instr-wrong":
            return True  # a field code written the wrong way for where it sits
        if kind == "text":
            if depth and w != wrapper:
                return True  # part of a field's result struck or put in, the field left live
            continue
        if kind == "begin":
            if depth:
                return True
            depth, wrapper, instr = 1, w, ""
        elif not depth:
            return True  # an instruction or separate/end outside any field
        elif w != wrapper:
            return True
        elif kind == "instr":
            instr += "x"
        elif kind == "end":
            if not instr:
                return True  # a field left without an instruction
            depth = 0
    return depth != 0


def _engine_text(p) -> str:
    """A paragraph's live text as the comparison engine reads it, field codes included."""
    from docx_redline.oxml.textmap import paragraph_text

    return " ".join(paragraph_text(p).split())


def _live_fields(p) -> int:
    """How many fields open in a paragraph outside a tracked deletion."""
    return sum(
        1
        for fc in p.iter(qn("w:fldChar"))
        if fc.get(qn("w:fldCharType")) == "begin" and not _inside_deletion(fc)
    )


def _drop_twin_insertions(doc, plains: dict) -> None:
    """The compare engine sometimes marks one changed paragraph as two: the form's paragraph
    struck whole, then a copy of the drafted paragraph inserted whole directly after it.
    Repairing the first would then put the drafted words in twice. The twin is known by what
    it says: its words are the drafted paragraph's words for the struck one's id. Its
    insertions move into the struck original, which becomes the one tracked change, and the
    twin goes."""
    for p in list(doc.element.body.iter(qn("w:p"))):
        first_runs = list(p.iter(qn("w:r")))
        struck = first_runs and all(r.getparent().tag == qn("w:del") for r in first_runs)
        if not struck:
            continue
        drafted = plains.get(_para_id(p))
        if drafted is None:
            continue
        # the twin is somewhere in the run of inserted paragraphs that follows the struck
        # one: the compare engine has been seen to put it last, after the paragraphs the
        # drafter added beneath it
        twin = p.getnext()
        while twin is not None and twin.tag == qn("w:p"):
            runs = list(twin.iter(qn("w:r")))
            if not (runs and all(r.getparent().tag == qn("w:ins") for r in runs)):
                twin = None
                break
            if _element_text(twin) == _element_text(drafted):
                break
            twin = twin.getnext()
        if twin is None or twin.tag != qn("w:p"):
            continue
        # the twin stands where the drafter placed the paragraph (restore_insert_order has
        # put it there); the struck original moves to it, so the one tracked change reads in
        # the drafter's order, not the form's
        if twin.getprevious() is not p:
            twin.addprevious(p)
        for child in list(twin):
            if child.tag != qn("w:pPr"):
                p.append(child)
        twin.getparent().remove(twin)


def _word_tokens(p) -> tuple[list[tuple[tuple[str, str], str, object]], dict[int, list]]:
    """A paragraph read as the pieces a word diff compares, in order: each word or stretch of
    spaces of its text (a word can run across runs: its parts are kept as (run, text)), each
    other thing a run holds (a tab, a footnote reference), and each field or other wrapper
    (hyperlink, simple field) whole. A piece is (key, kind, payload). The paragraph's markers
    (bookmarks and the like) are returned apart, by the index of the piece they stand before."""
    from lxml import etree

    pieces: list[tuple[tuple[str, str], str, object]] = []
    markers: dict[int, list] = {}
    field: list = []
    depth = 0
    joining = False  # whether the next text can carry on the last piece's word

    def text(run, words: str) -> None:
        nonlocal joining
        for word in re.findall(r"\s+|\S+", words):
            key = ("w", " " if word.isspace() else word)
            last = pieces[-1] if pieces and joining else None
            if last is not None and last[0][1].isspace() == key[1].isspace():
                merged = " " if key[1] == " " else last[0][1] + word
                pieces[-1] = (("w", merged), "word", [*last[2], (run, word)])
            else:
                pieces.append((key, "word", [(run, word)]))
            joining = True

    for child in p:
        if child.tag == qn("w:pPr"):
            continue
        kinds = [f.get(qn("w:fldCharType")) for f in child.iter(qn("w:fldChar"))]
        if child.tag == qn("w:r") and (depth or "begin" in kinds):
            field.append(child)
            depth += kinds.count("begin") - kinds.count("end")
            if depth <= 0:
                key = ("f", b"".join(etree.tostring(x) for x in field).decode())
                pieces.append((key, "whole", field))
                field, depth, joining = [], 0, False
        elif depth:
            field.append(child)
        elif child.tag == qn("w:r"):
            for part in child:
                if part.tag == qn("w:t"):
                    text(child, part.text or "")
                elif part.tag != qn("w:rPr"):
                    pieces.append((("c", etree.tostring(part).decode()), "child", (child, part)))
                    joining = False
        elif child.find(f".//{qn('w:r')}") is not None:
            pieces.append((("e", etree.tostring(child).decode()), "whole", [child]))
            joining = False
        else:
            markers.setdefault(len(pieces), []).append(child)
            joining = False
    if field:
        pieces.append((("f", ""), "whole", field))
    return pieces, markers


def _track_word_changes(p, old, new, n: int, author: str, stamp: str) -> int:
    """The form's paragraph `old` marked up into `p` as the words that changed to read as
    `new`: unchanged words keep the form's runs, the words put in take the formatting of the
    form's words where they stand, its bookmarks stay where they stood, and a changed field is
    struck or put in whole. Returns the next free revision id."""
    olds, markers = _word_tokens(old)
    news, _ = _word_tokens(new)
    # where the next piece goes (p itself, or the open w:del / w:ins), the run being filled and
    # the form's run it copies, and the form's run the last words came from
    first = next((x[2][0][0] for x in olds if x[1] == "word"), None)
    at: dict = {"into": None, "kind": "", "run": None, "src": None, "form": first}

    def tracked(kind: str):
        nonlocal n
        n += 1
        tag = qn("w:del") if kind == "delete" else qn("w:ins")
        return p.makeelement(tag, {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp})

    def into(kind: str):
        if at["into"] is None or at["kind"] != kind:
            at["into"] = p if kind == "equal" else tracked(kind)
            if at["into"] is not p:
                p.append(at["into"])
            at["kind"], at["run"] = kind, None
        return at["into"]

    def close() -> None:
        at["into"], at["kind"], at["run"] = None, "", None

    def run_like(kind: str, src):
        parent = into(kind)
        if at["run"] is None or at["src"] is not src:
            run = copy.deepcopy(src)
            for c in list(run):
                if c.tag != qn("w:rPr"):
                    run.remove(c)
            parent.append(run)
            at["run"], at["src"] = run, src
        return at["run"]

    def put(kind: str, piece) -> None:
        _, what, payload = piece
        if what == "child":
            run_like(kind, payload[0]).append(copy.deepcopy(payload[1]))
        elif what == "word":
            for src, words in payload:
                if kind != "insert":
                    at["form"] = src
                run = run_like(kind, at["form"] if kind == "insert" and at["form"] is not None else src)
                if len(run) and run[-1].tag == qn("w:t"):
                    run[-1].text = (run[-1].text or "") + words
                else:
                    t = run.makeelement(qn("w:t"), {})
                    t.text = words
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    run.append(t)
        else:
            for el in payload:
                el = copy.deepcopy(el)
                if el.tag == qn("w:r") or kind == "equal":
                    into(kind).append(el)
                    at["run"] = None
                    continue
                # a hyperlink or simple field: the change is marked on the runs inside it
                close()
                for r in el.findall(qn("w:r")):
                    wrap = tracked(kind)
                    r.addprevious(wrap)
                    wrap.append(r)
                p.append(el)

    def markers_upto(i: int) -> None:
        for k in sorted(k for k in markers if k <= i):
            close()
            for m in markers.pop(k):
                p.append(copy.deepcopy(m))

    keys = difflib.SequenceMatcher(None, [x[0] for x in olds], [x[0] for x in news], autojunk=False)
    for op, i1, i2, j1, j2 in keys.get_opcodes():
        for i in range(i1, i2):
            markers_upto(i)
            put("equal" if op == "equal" else "delete", olds[i])
        for j in range(j1, j2) if op != "equal" else ():
            put("insert", news[j])
    markers_upto(len(olds))
    for dele in p.iter(qn("w:del")):
        _as_deleted(dele)
    return n


def repair_fields(
    redline_bytes: bytes, original_bytes: bytes, plain_bytes: bytes, author: str = "Drafting engine"
) -> bytes:
    """Every paragraph whose fields the compare broke is marked up again from the form's
    paragraph: only the words that changed are tracked, each field struck or inserted whole,
    and the form's bookmarks and runs stay on the rest. A paragraph the form does not hold is
    inserted whole. Paragraphs are matched by their ids, which the compare keeps."""
    originals = {
        _para_id(p): p for p in Document(io.BytesIO(original_bytes)).element.body.iter(qn("w:p"))
    }
    plains = {
        _para_id(p): p for p in Document(io.BytesIO(plain_bytes)).element.body.iter(qn("w:p"))
    }
    doc = Document(io.BytesIO(redline_bytes))
    _drop_twin_insertions(doc, plains)
    # the compare reads a field's code as words: a paragraph it pairs with a drafted paragraph
    # holding a field can come out reading that paragraph's words with the code typed in as text
    with_fields: dict[str, object] = {}
    for q in plains.values():
        if q.find(f".//{qn('w:fldChar')}") is not None:
            with_fields.setdefault(_engine_text(q), q)
    n = 97000
    stamp = "2026-01-01T00:00:00Z"
    for p in list(doc.element.body.iter(qn("w:p"))):
        pid = _para_id(p)
        if _fields_broken(p):
            new = plains.get(pid)
        else:
            twin = with_fields.get(_engine_text(p)) if p.find(qn("w:ins")) is not None else None
            new = twin if twin is not None and _live_fields(p) < _live_fields(twin) else None
        if new is None:
            continue
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)
        old = originals.get(pid)
        if old is not None:
            n = _track_word_changes(p, old, new, n, author, stamp)
            # the word diff reads spacing as no change: the drafter's spacing stands
            n = _respace(p, _spacing_text(new), n, author)
            continue
        n += 1
        ins = p.makeelement(
            qn("w:ins"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
        )
        for r in new.findall(qn("w:r")):
            ins.append(copy.deepcopy(r))
        if len(ins):
            p.append(ins)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def accepted_view(
    plain_bytes: bytes,
    deferred_rows: list[list[str]],
    deferred_paragraphs: list[str],
    deferred_cells: list[tuple[str, str, str]],
) -> bytes:
    """The draft as it will read once every deferred change is made and accepted: the rows and
    paragraphs the drafters removed are taken out, the cell paragraphs they changed carry their
    new words. For reading, never for delivery: it has no tracked changes."""
    doc = Document(io.BytesIO(plain_bytes))
    body = doc.element.body
    by_id = {_para_id(p): p for p in body.iter(qn("w:p")) if _para_id(p)}
    by_id.update(_header_paragraphs(doc))
    gone: set[str] = set(deferred_paragraphs)
    for row_ids in deferred_rows:
        first = by_id.get(row_ids[0]) if row_ids else None
        if first is None:
            continue
        tr = next((a for a in first.iterancestors() if a.tag == qn("w:tr")), None)
        if tr is not None and tr.getparent() is not None:  # two strikes may name one row
            tr.getparent().remove(tr)
    for pid in gone:
        p = by_id.get(pid)
        if p is not None and p.getparent() is not None:
            p.getparent().remove(p)
    for pid, final, original in deferred_cells:
        p = by_id.get(pid)
        if p is not None and p.getparent() is not None:
            first, *more = final.split(PARA_BREAK)
            like = copy.deepcopy(p)
            set_paragraph_text(p, first.strip(), original)
            for piece in reversed(more):  # the paragraphs the cell gained, in order after it
                new = copy.deepcopy(like)
                _fresh_ids(new)
                set_paragraph_text(new, piece.strip())
                p.addnext(new)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def inserts_not_placed(docx_bytes: bytes, snapshots: list[dict]) -> list[str]:
    """The words of each paragraph, row or table the snapshots add that the finished draft,
    every change accepted, does not hold: whatever step dropped it, it is named. Numbers are
    set aside: the writer renumbers a cross-reference after the drafter wrote it ("Schedule 4"
    printed "Schedule 3"), and the question is whether the paragraph is there."""

    def key(text: str) -> str:
        return re.sub(r"\d+", "#", _key(text))

    doc = Document(io.BytesIO(accept_all_changes(docx_bytes)))
    held: set[str] = set()
    for el in doc.element.body.iter(qn("w:p"), qn("w:tr")):
        if el.tag == qn("w:tr"):
            held.add(key(_row_text(el)))
            continue
        # a tab read as a space, as the view reads it: "7.<tab>Notices" is "7. Notices"
        held.add(key(_element_text(el, tabs=True)))
        held.add(key(_element_text(el)))
        held.update(key(line) for line in _lines_of(el))
    added: list[str] = []
    for snap in snapshots:
        added += [t for entry in snap.get("inserted", []) for t in entry[1]]
        added += [" | ".join(cells) for _, cells in snap.get("rows_inserted", [])]
        for _, header, rows in snap.get("tables_inserted", []):
            added += [" | ".join(r) for r in [header, *rows]]
    return [
        t
        for t in added
        if key(t) not in held
        # a row with no table to join is written as a paragraph of its filled cells
        and key(" | ".join(c for c in t.split(" | ") if c.strip())) not in held
    ]


def strike_dropped_footnotes(
    redline_bytes: bytes, plain_bytes: bytes, author: str = "Drafting engine"
) -> bytes:
    """The comparison engine keeps a footnote reference the plain document dropped, because a
    reference has no text to compare. Every reference in the redline that its paragraph no
    longer has in the plain document is struck as a tracked deletion."""
    plain = Document(io.BytesIO(plain_bytes))
    kept: dict[str, list[str]] = {}
    for p in plain.element.body.iter(qn("w:p")):
        refs = [x.get(qn("w:id")) for x in p.iter(qn("w:footnoteReference"))]
        if refs:
            kept[_para_id(p)] = refs
    doc = Document(io.BytesIO(redline_bytes))
    n = 95000
    stamp = "2026-01-01T00:00:00Z"
    for p in doc.element.body.iter(qn("w:p")):
        allowed = list(kept.get(_para_id(p), []))
        for ref in list(p.iter(qn("w:footnoteReference"))):
            run = ref.getparent()
            if run.getparent().tag == qn("w:del"):
                continue
            if ref.get(qn("w:id")) in allowed:
                allowed.remove(ref.get(qn("w:id")))
                continue
            n += 1
            parent = run.getparent()
            wrap = parent.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            parent.replace(run, wrap)
            wrap.append(run)
    _strike_note_text(doc, author, stamp)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _strike_note_text(doc: object, author: str, stamp: str) -> None:
    """A footnote whose every reference is struck has its note's text struck too, as Word does
    when a reference is deleted; otherwise the note prints live at the foot of the page."""
    from lxml import etree

    live: set[str] = set()
    struck: set[str] = set()
    for ref in doc.element.body.iter(qn("w:footnoteReference")):  # type: ignore[attr-defined]
        gone = any(a.tag == qn("w:del") for a in ref.iterancestors())
        (struck if gone else live).add(ref.get(qn("w:id")) or "")
    struck -= live
    rels = doc.part.rels.values()  # type: ignore[attr-defined]
    part = next((r.target_part for r in rels if r.reltype.endswith("/footnotes")), None)
    if not struck or part is None:
        return
    root = getattr(part, "_element", None)
    tree = root if root is not None else etree.fromstring(part.blob)
    n = 96000
    for note in tree.iter(qn("w:footnote")):
        if note.get(qn("w:id")) not in struck:
            continue
        for run in list(note.iter(qn("w:r"))):
            if run.getparent().tag == qn("w:del"):
                continue
            _as_deleted(run)
            n += 1
            wrap = run.makeelement(
                qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp}
            )
            run.getparent().replace(run, wrap)
            wrap.append(run)
    if root is None:
        part._blob = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _replace_within(body: _Body, original: str, final: str, after: int = 0) -> bool:
    """A rendered paragraph that is one line of a Word paragraph (a line break inside it): the
    line's words are replaced within that paragraph. With several candidate hosts, the first
    from `after` in document order is taken, as the changes come in document order."""
    want = _key(original)
    hosts = [
        (i, el)
        for i, (k, el) in enumerate(body.items)
        if i >= after and el.tag == qn("w:p") and want and want in k and k != want
    ]
    if not hosts or (after == 0 and len(hosts) != 1):
        return False
    el = hosts[0][1]
    text = _element_text(el)
    plain_original = " ".join(_FOOTNOTE.sub("", original).split())
    if plain_original not in text:
        # spacing may differ; fall back to a whitespace-tolerant search
        pattern = r"\s*".join(re.escape(w) for w in plain_original.split())
        m = re.search(pattern, text)
        if not m:
            return False
        text = text[: m.start()] + " ".join(_FOOTNOTE.sub("", final).split()) + text[m.end() :]
    else:
        text = text.replace(plain_original, " ".join(_FOOTNOTE.sub("", final).split()), 1)
    set_paragraph_text(el, text)
    return True


def _after(body: _Body, hit: tuple[int, object] | None, at: int | None):
    """A paragraph found by its rank, unless it lies before `at`, where the paragraph before
    it in its unit was found: then it is another paragraph reading alike, and is given back. A
    row spanning several lines of the render matches no line of the units before it, so its
    rank counts from the top of the document and names the first such row there. A paragraph
    no other reads alike is the one: `at` can be a later twin of the paragraph before it."""
    if hit is None or at is None or hit[0] > at:
        return hit
    want = body.items[hit[0]][0]
    if sum(k == want for k, _ in body.items) == 1:
        return hit
    body.used.discard(hit[0])
    return None


def _place_of(
    body: _Body, prior: str, here: dict[str, list[int]], earlier: dict | None = None
) -> int | None:
    """Where the paragraph before one was found: the last place found for its words in this
    pass over the unit (`here`, made in document order), else the one place found for them in
    the unit's `earlier` pass, else the one paragraph of the document reading so; None where
    that is not known. A header or footer line is no place: the body is listed before them."""
    want = _key(prior)
    if not want:
        return None
    if want in here:
        places = here[want][-1:]
    elif earlier and want in earlier:
        places = earlier[want]
    else:
        places = [i for i, (k, _) in enumerate(body.items) if k == want]
    if len(places) != 1 or body.items[places[0]][1].getroottree().getroot() is not body.doc.element:
        return None
    return places[0]


def _line_cells(line: str) -> list[str]:
    """A rendered row line's cells: an empty first or last cell leaves its separator at the
    line's edge with no space beside it (") |")."""
    return [x.strip() for x in re.split(r"(?:^|\s)\|(?:\s|$)", line)]


def _row_line(body: _Body, original: str, after: int | None) -> list[int] | None:
    """The cell paragraphs one rendered line of a table row reads as, where the view named the
    line on its own. The render puts a row's cell paragraphs on lines of their own (a first
    cell's earlier ones before the row's line), and a view that did not join them back into
    the row names them line by line. A line is one paragraph of a cell or, across " | ", the
    last paragraph with words of one cell and the first paragraph of each cell after it. The
    first fit after `after`, where the paragraph before the line was found; with no such
    place, only a line that fits one place in the document. Never inside a row already taken."""
    pieces = [_key(x) for x in _line_cells(original)]
    if not pieces[0]:
        return None
    taken_rows = {body.items[i][1] for i in body.used if body.items[i][1].tag == qn("w:tr")}
    where = {el: i for i, (_, el) in enumerate(body.items) if i not in body.inline}
    fits = []
    for i in range(after or 0, len(body.items)):
        key, el = body.items[i]
        if key != pieces[0] or i in body.used or i in body.inline or el.tag != qn("w:p"):
            continue
        cell = el.getparent()
        if cell.tag != qn("w:tc") or cell.getparent() in taken_rows:
            continue
        cells = cell.getparent().findall(qn("w:tc"))
        worded = [p for p in cell.findall(qn("w:p")) if _key(_element_text(p))]
        if len(pieces) > 1 and worded[-1] is not el:
            continue
        found = [i]
        for later, want in zip(cells[cells.index(cell) + 1 :], pieces[1:], strict=False):
            first = later.find(qn("w:p"))
            j = where.get(first) if first is not None else None
            if j is None or j in body.used or body.items[j][0] != want:
                break
            found.append(j)
        if len(found) == len(pieces):
            fits.append(found)
            if after is not None:
                break
    if len(fits) != 1:
        return None
    body.used.update(fits[0])
    return fits[0]


def _taken(body: _Body, i: int | None) -> tuple[int, object] | None:
    """Item i taken, if it is one and not taken already."""
    if i is None or i in body.used:
        return None
    body.used.add(i)
    return i, body.items[i][1]


def _located(body: _Body, before: str, rank: int, where: list | None = None):
    """The live element an addition's anchor text names: by what it follows, else by its
    rank, else the first reading so. A line of a table cell holding several paragraphs is that
    paragraph, by what it follows; for a line across a row's cells, its first cell's."""
    if not _key(before):
        return None
    if where:
        i = body.following(before, where[0], where[1])
        if i is not None:
            return body.items[i][1]
        i = body.in_place(_line_cells(before)[0], where[0], where[1], rank)
        if i is not None and _in_cell_of_several(body.items[i][1]):
            return body.items[i][1]
    if rank >= 0:
        el = _nth_live(body, before, rank)
        if el is not None:
            return el
        if sum(1 for k, _ in body.items if k == _key(before)) > 1:
            # one of several reading alike, and the counted one is gone: the first reading so
            # is another place, so none is named and the addition is reported, never misplaced
            return None
    return _find_live(body, before)


_TYPED_NUMBER = re.compile(r"\s*\d+(?:\.\d+)*\.?\s+")


def _drops_typed_number(p, final: str) -> bool:
    """Whether `p`'s words open with a number typed into them ("2. ") that `final` does not."""
    typed = _TYPED_NUMBER.match(_element_text(p))
    return typed is not None and not final.lstrip().startswith(typed.group().strip())


def _strike_empty_between(struck: list, doc, referenced: set[str], gone: list[str] | None) -> None:
    """Paragraphs without words lying between two paragraphs a piece struck go with them:
    the drafter never sees an empty paragraph, and left standing a numbered one prints as an
    empty "(a)" (a precedent's empty limb heading sub-limbs that are both struck). One
    holding a bookmark a cross-reference names stays."""
    if gone is None:
        return
    for el in struck:
        empty, q = [], el.getnext()
        while q is not None and q.tag == qn("w:p") and not _key(_element_text(q)):
            empty.append(q)
            q = q.getnext()
        ids = [_para_id(p) for p in empty]
        if not empty or q is None or _para_id(q) not in gone or not all(ids):
            continue
        if any({b.get(qn("w:name")) for b in p.iter(qn("w:bookmarkStart"))} & referenced for p in empty):
            continue
        gone.extend(x for x in ids if x not in gone)


def _holds_structure(empty: list, doc, referenced: set[str], gone: set[str] | None = None) -> str:
    """Why the paragraphs without words in `empty` are not struck, "" where nothing holds
    them: one is numbered and the paragraph with words after them is a deeper level of its
    list (an empty "(a)" heading its "(i)", "(ii)": striking it renumbers the clause), or one
    holds a bookmark a cross-reference field names. Paragraphs being struck (`gone`, by id)
    are passed over: limbs the draft strikes hold no parent."""
    from docx.text.paragraph import Paragraph

    from drafting_engine.word.numbering import _paragraph_numbering

    def level(p) -> int | None:
        num_id, ilvl = _paragraph_numbering(Paragraph(p, doc._body))
        return ilvl if num_id not in (None, "0") else None

    after = empty[-1].getnext() if empty else None
    while after is not None and (
        after.tag != qn("w:p") or (gone and _para_id(after) in gone) or not _key(_element_text(after))
    ):
        after = after.getnext()
    below = level(after) if after is not None else None
    for p in empty:
        names = {b.get(qn("w:name")) for b in p.iter(qn("w:bookmarkStart"))}
        if names & referenced:
            return f"it holds {', '.join(sorted(names & referenced))}, which a cross-reference names"
        mine = level(p)
        if mine is not None and below is not None and below > mine:
            return "it is a numbered paragraph the paragraphs after it sit a deeper level under"
    return ""


def _in_cell_of_several(p) -> bool:
    """Whether `p` is one of several paragraphs of a cell in a row of several cells."""
    cell = p.getparent()
    if p.tag != qn("w:p") or cell is None or cell.tag != qn("w:tc"):
        return False
    row = cell.getparent()
    return len(cell.findall(qn("w:p"))) > 1 and row is not None and len(row.findall(qn("w:tc"))) > 1


def _empty_after(el) -> list:
    """The paragraphs without words directly after `el`, up to the next paragraph with words
    or a table: pictures, page breaks, empty numbered paragraphs. A paragraph holding a section
    break ends them, kept: striking its mark would join two sections."""
    out = []
    nxt = el.getnext()
    while nxt is not None and nxt.tag != qn("w:tbl"):
        if nxt.tag == qn("w:p"):
            ppr = nxt.find(qn("w:pPr"))
            if _key(_element_text(nxt)) or (ppr is not None and ppr.find(qn("w:sectPr")) is not None):
                break
            out.append(nxt)
        elif nxt.tag not in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            break
        nxt = nxt.getnext()
    return out


def _nth_live(body: _Body, original: str, rank: int):
    """The rank-th element (0-based, in document order, counting as the render counts) whose
    text is `original`, whether or not it has been taken; None where that one is no longer in
    the document. The rank was counted on the form, so one taken out still counts: skipping it
    would put the addition after the next paragraph reading alike (a schedule's conditions
    under another schedule's heading of the same name)."""
    want = _key(original)
    seen = 0
    for i, (k, el) in enumerate(body.items):
        if k != want or not body._counts(i):
            continue
        if seen == rank:
            return el if el.getparent() is not None else None
        seen += 1
    return None


def _find_live(body: _Body, original: str):
    """The element for an original, whether or not it has been taken (an insert follows a
    paragraph that may itself have been changed)."""
    want = _key(original)
    for k, el in body.items:
        if k == want and el.getparent() is not None:
            return el
    return None


def _strike_fragment_tables(doc, followers: list, words: set[str], deferred_rows) -> set[str]:
    """Each table whose every paragraph reads as one of `words`, or none at all, and that stood
    straight after something the unit struck (`followers`, read as each was struck, or a table
    struck here) is struck, its rows deferred where `deferred_rows` is given. Returns the words
    of the tables struck."""
    struck_ids = {id(el) for el in followers if el is not None}
    taken: set[str] = set()
    done: set[int] = set()
    grew = True
    while grew:  # a table standing after one struck here follows a strike too
        grew = False
        for tbl in list(doc.element.body.iter(qn("w:tbl"))):
            if id(tbl) in done or id(tbl) not in struck_ids:
                continue
            texts = [_key(_element_text(p)) for p in tbl.iter(qn("w:p"))]
            texts = [t for t in texts if t]
            if not all(t in words for t in texts):  # an empty table is covered: a spacer
                continue
            nxt = tbl.getnext()
            while nxt is not None and nxt.tag == qn("w:p") and not _key(_element_text(nxt)):
                nxt = nxt.getnext()
            if nxt is not None:
                struck_ids.add(id(nxt))
            done.add(id(tbl))
            grew = True
            for tr in tbl.findall(qn("w:tr")):
                if deferred_rows is not None:
                    deferred_rows.append([_para_id(x) for x in tr.iter(qn("w:p"))])
                else:
                    tbl.remove(tr)
            if deferred_rows is None and tbl.getparent() is not None:
                tbl.getparent().remove(tbl)
            taken.update(texts)
    return taken


def _find_near(
    body: _Body,
    original: str,
    anchor,
    struck: bool = False,
    prefer_struck: bool = False,
    struck_here: list[int] | None = None,
    besides=None,
):
    """The live element reading as `original` nearest to `anchor`, the earlier on a tie: the
    heading a new sub-clause is a sibling of sits in the piece it is written in, while a table
    cell reading "Term" can sit anywhere. With `struck`, for a style to copy, a form paragraph
    already taken out answers where none still stands: the term sheet's conditions are written in
    the style of the precedent's own, which the same piece strikes. With `prefer_struck`, the
    drafter named a paragraph it struck: one taken out answers first, never a live one reading
    alike (a schedule title and its first heading can both read "Service Levels").
    `struck_here` is where the piece's own deletions of those words were found: a paragraph
    struck by its id afterwards still stands while the piece's new paragraphs are written.
    `besides` is never the answer: a paragraph restyled like words it reads itself takes the
    other paragraph's style (one "where:" like the other)."""
    want = _key(original)
    at = next((i for i, (_, el) in enumerate(body.items) if el is anchor), len(body.items))
    if prefer_struck and struck_here:
        return body.items[min(struck_here, key=lambda i: (abs(i - at), i > at))][1]
    found = [(i, el) for i, (k, el) in enumerate(body.items) if k == want and el is not besides]
    live = [(i, el) for i, el in found if el.getparent() is not None]
    gone = [(i, el) for i, el in found if el.getparent() is None]
    hits = (gone or live) if prefer_struck else (live or (found if struck else []))
    if not hits:
        return None
    return min(hits, key=lambda h: (abs(h[0] - at), h[0] > at))[1]


def unique_paragraph_ids(docx_bytes: bytes) -> bytes:
    """Every paragraph with its own w14:paraId and w14:textId. A comparison engine that pairs a
    deleted paragraph with its replacement copies the ids across, and Word refuses the file."""
    import secrets

    doc = Document(io.BytesIO(docx_bytes))
    seen: set[str] = set()
    roots = [doc.element.body] + [
        part._element
        for s in doc.sections
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer)
        if not part.is_linked_to_previous
    ]
    for root in roots:
        for p in root.iter(qn("w:p")):
            for attr in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
                val = p.get(attr)
                if val is None:
                    continue
                key = attr + val
                if key in seen:
                    p.set(attr, f"{secrets.randbelow(0x7FFFFFFF):08X}")
                else:
                    seen.add(key)
        _unique_bookmarks(root)
        _unique_comment_ranges(root)
        _unique_revision_ids(root, seen)
        _balance_fields(root)
        for sdt_id in root.iter(qn("w:id")):
            if sdt_id.getparent().tag == qn("w:sdtPr"):
                key = "sdt" + (sdt_id.get(qn("w:val")) or "")
                if key in seen:
                    sdt_id.set(qn("w:val"), str(secrets.randbelow(0x7FFFFFFF)))
                seen.add(key)
    out = io.BytesIO()
    doc.save(out)
    return _unique_comment_ids(out.getvalue())


def _unique_comment_ids(docx_bytes: bytes) -> bytes:
    """A comparison engine that copies a commented paragraph for its deleted and inserted
    versions copies the comment too, durable id and all. Word refuses two comments with one
    durable id; the copy gets its own."""
    import secrets
    import zipfile

    from lxml import etree

    name = "word/commentsIds.xml"
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if name not in z.namelist():
            return docx_bytes
        root = etree.fromstring(z.read(name))
        seen: set[str] = set()
        changed = False
        for e in root.iter():
            for attr, val in list(e.attrib.items()):
                if not attr.endswith("}durableId"):
                    continue
                if val in seen:
                    e.set(attr, f"{secrets.randbelow(0x7FFFFFFF):08X}")
                    changed = True
                seen.add(val)
        if not changed:
            return docx_bytes
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == name:
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                zo.writestr(item, data)
        return out.getvalue()


def _unique_comment_ranges(root) -> None:
    """One range and one reference per comment. A comparison engine that keeps a commented
    passage in both its deleted and inserted forms marks the comment twice."""
    for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
        seen: set[str] = set()
        for e in list(root.iter(qn(tag))):
            cid = e.get(qn("w:id"))
            if cid in seen:
                victim = e.getparent() if tag == "w:commentReference" else e
                victim.getparent().remove(victim)
            seen.add(cid)


def _unique_revision_ids(root, seen: set[str]) -> None:
    """Every tracked change its own w:id. Each writer stage numbers its changes from its own
    fixed base, so two stages can hand out the same id."""
    tags = {qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"), qn("w:pPrChange")}
    top = max(
        (int(e.get(qn("w:id"))) for e in root.iter(*tags) if (e.get(qn("w:id")) or "").isdigit()),
        default=0,
    )
    for e in root.iter(*tags):
        key = "rev" + (e.get(qn("w:id")) or "")
        if key in seen:
            top += 1
            e.set(qn("w:id"), str(top))
            key = "rev" + str(top)
        seen.add(key)


def _unique_bookmarks(root) -> None:
    """A comparison engine that copies a row for its deleted and inserted versions copies the
    row's bookmark with it; two bookmarks with one id is another file Word refuses."""
    seen: set[str] = set()
    ends = {}
    for e in root.iter(qn("w:bookmarkEnd")):
        ends.setdefault(e.get(qn("w:id")), []).append(e)
    top = max((int(b.get(qn("w:id"))) for b in root.iter(qn("w:bookmarkStart"))), default=0)
    for b in root.iter(qn("w:bookmarkStart")):
        bid = b.get(qn("w:id"))
        if bid not in seen:
            seen.add(bid)
            continue
        top += 1
        fresh = str(top)
        seen.add(fresh)
        b.set(qn("w:id"), fresh)
        if ends.get(bid):
            ends[bid].pop().set(qn("w:id"), fresh)


def _balance_fields(root) -> None:
    """Every field must open and close. A comparison engine that marks a rewritten
    cross-reference as deleted can drop the field's opening and keep its close; the field's
    control runs are then removed from the paragraph and its visible text stands as words.

    A field may span paragraphs: the contents page is one field whose begin and instruction
    sit on its first line and whose end sits after its last. So fields are matched across the
    whole part, not paragraph by paragraph. A half field is known by what it carries: a
    separator or code with no begin before it in its paragraph belongs to a field whose begin
    was dropped, and the end after it closes that half field, not a field opened earlier."""
    fields: list[dict] = []  # each: its paragraph, its marks, whether it has its separator
    open_fields: list[dict] = []
    halves: set = set()  # paragraphs holding a half field not yet closed
    bad: set = set()
    for mark in root.iter(qn("w:fldChar"), qn("w:instrText")):
        p = next(mark.iterancestors(qn("w:p")))
        kind = mark.get(qn("w:fldCharType")) if mark.tag == qn("w:fldChar") else "code"
        top = open_fields[-1] if open_fields else None
        if kind == "begin":
            field = {"p": p, "marks": [mark], "separated": False}
            fields.append(field)
            open_fields.append(field)
        elif p in halves:
            bad.add(p)
            if kind == "end":
                halves.discard(p)
        elif top is not None and not (top["separated"] and kind != "end"):
            top["marks"].append(mark)
            if kind == "separate":
                top["separated"] = True
            elif kind == "end":
                open_fields.pop()
        else:
            bad.add(p)
            if kind != "end":
                halves.add(p)
    bad |= {f["p"] for f in open_fields}
    if not bad:
        return
    # a field closed in a later paragraph than it opened keeps its marks
    spanning = {
        m
        for f in fields
        if f not in open_fields and next(f["marks"][-1].iterancestors(qn("w:p"))) is not f["p"]
        for m in f["marks"]
    }
    for p in bad:
        for mark in list(p.iter(qn("w:fldChar"), qn("w:instrText"))):
            r = mark.getparent()
            if mark not in spanning and r.getparent() is not None:
                r.getparent().remove(r)


def without_content_controls(docx_bytes: bytes) -> bytes:
    """The document with every content control unwrapped, its contents kept in place. The
    form's cover page wraps its party cells in content controls; a comparison engine that meets
    a changed cell inside one writes a row Word refuses to open. Bisected with Word: the same
    change with the controls unwrapped opens.

    Every paragraph without a w14:paraId is given one, the same each time: the passes after
    the compare (struck rows and paragraphs, rewritten cells, restyles) find paragraphs by it,
    and a form rebuilt by hand carries none on many of its paragraphs."""
    doc = Document(io.BytesIO(docx_bytes))
    parts = [doc.element.body] + [
        part._element
        for section in doc.sections
        for part in (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        )
        if not part.is_linked_to_previous
    ]
    for root in parts:
        for sdt in list(root.iter(qn("w:sdt"))):
            content = sdt.find(qn("w:sdtContent"))
            parent = sdt.getparent()
            i = parent.index(sdt)
            for child in list(content) if content is not None else []:
                parent.insert(i, child)
                i += 1
            parent.remove(sdt)
    paragraphs = [p for root in parts for p in root.iter(qn("w:p"))]
    taken = {_para_id(p) for p in paragraphs}
    n = 0
    for p in paragraphs:
        if _para_id(p):
            continue
        while f"{n:08X}" in taken or n == 0:
            n += 1
        p.set(f"{{{W14}}}paraId", f"{n:08X}")
        taken.add(f"{n:08X}")
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def without_comments(docx_bytes: bytes) -> bytes:
    """The form with its own comments taken out, its text untouched. The form's comments are
    notes to whoever drafts from it ("Check this period with the business team"); the
    drafters read the form without them, and left in they would go out with the draft. A
    finished contract keeps few of the form's comments, each turned into its own query; the
    draft's comments are its own reasons and open points."""
    import zipfile

    from lxml import etree

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if "word/comments.xml" not in z.namelist():
            return docx_bytes
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in z.infolist():
                data = z.read(item.filename)
                name = item.filename
                if name.startswith("word/comments") and name.endswith(".xml"):
                    root = etree.fromstring(data)
                    for child in list(root):
                        root.remove(child)
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                elif name == "word/document.xml":
                    root = etree.fromstring(data)
                    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
                        for e in list(root.iter(qn(tag))):
                            e.getparent().remove(e)
                    for ref in list(root.iter(qn("w:commentReference"))):
                        run = ref.getparent()
                        run.remove(ref)
                        if run.tag == qn("w:r") and not [
                            c for c in run if c.tag != qn("w:rPr")
                        ]:
                            run.getparent().remove(run)
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                zo.writestr(item, data)
        return out.getvalue()


def no_inserted_titles(docx_bytes: bytes) -> bytes:
    """No paragraph the draft added stands at a title level, one whose number carries a word
    ("Part %1"): there it is a new Part, and takes the paragraphs below it out of the Part they
    belong to. Groups added under a form's empty Part can come out numbered as the Parts that
    follow it, outside the Part they were written under.
    Such a paragraph, and the added paragraphs that follow it, move one level down, each to the
    style the form numbers at the next level."""
    from drafting_engine.word.numbering import (
        _paragraph_numbering,
        load_numbering_definitions,
    )

    doc = Document(io.BytesIO(docx_bytes))
    numbering = load_numbering_definitions(docx_bytes)

    def numbered(style) -> tuple[str | None, int | None]:
        num_id = ilvl = None
        seen: set[str] = set()
        while style is not None and (num_id is None or ilvl is None) and style.style_id not in seen:
            seen.add(style.style_id)
            ppr = style.element.find(qn("w:pPr"))
            npr = ppr.find(qn("w:numPr")) if ppr is not None else None
            if npr is not None:
                n, lv = npr.find(qn("w:numId")), npr.find(qn("w:ilvl"))
                num_id = num_id or (n.get(qn("w:val")) if n is not None else None)
                ilvl = ilvl if ilvl is not None else (int(lv.get(qn("w:val"))) if lv is not None else None)
            style = style.base_style
        return (num_id, ilvl if ilvl is not None else 0) if num_id else (None, None)

    by_level = {}
    for style in doc.styles:
        if style.type == 1:
            key = numbered(style)
            if key[0] is not None:
                by_level.setdefault(key, style)

    def added(p) -> bool:
        ppr = p._p.find(qn("w:pPr"))
        rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
        return rpr is not None and rpr.find(qn("w:ins")) is not None

    def is_title(num_id: str | None, ilvl: int | None) -> bool:
        level = numbering.level(num_id, ilvl) if num_id is not None and ilvl is not None else None
        return level is not None and bool(re.search(r"[A-Za-z]", re.sub(r"%\d", "", level.level_text)))

    def level_of(p) -> tuple[str | None, int | None]:
        return _paragraph_numbering(p)

    def set_level(p, num_id: str, ilvl: int) -> bool:
        lower = by_level.get((num_id, ilvl))
        if lower is None or p.style.style_id == lower.style_id:
            return False
        p.style = lower
        return True

    # a new schedule or appendix (level 0) is the draft's to add; a title inside one is not
    changed = False
    mode = ""  # "groups": siblings of the first group under an empty title; "lead": lead-ins
    lead = 0
    prev: tuple[str | None, int | None, bool] = (None, None, False)  # numId, level, a title
    for p in doc.paragraphs:
        num_id, ilvl = level_of(p)
        words = "".join(t.text or "" for t in p._p.iter(qn("w:t")))  # its inserted runs included
        if not added(p):
            mode = ""
            if words.strip():
                prev = (num_id, ilvl, is_title(num_id, ilvl))
            continue
        if num_id is None or ilvl is None or not words.strip():
            continue
        if ilvl > 0 and is_title(num_id, ilvl):
            if mode == "groups" or (prev[0] == num_id and prev[1] == ilvl and prev[2]):
                # under an empty title (an empty Part): a group, one level below it
                mode = "groups"
                changed |= set_level(p, num_id, ilvl + 1)
                prev = (num_id, ilvl + 1, False)
                continue
            if prev[0] == num_id and prev[1] is not None and prev[1] > ilvl:
                # after deeper text: a lead-in to what follows, one level below that text,
                # each further one its sibling; what follows it keeps its level
                if mode != "lead":
                    mode, lead = "lead", prev[1] + 1
                changed |= set_level(p, num_id, lead)
                continue
        elif mode == "groups" and num_id == prev[0]:
            changed |= set_level(p, num_id, ilvl + 1)
            prev = (num_id, ilvl + 1, False)
            continue
        prev = (num_id, ilvl, is_title(num_id, ilvl))
    if not changed:
        return docx_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def redline(original_bytes: bytes, revised_bytes: bytes, author: str = "Drafting engine") -> bytes:
    """The revised document marked up against the original, by a comparison engine.

    docx-redline compares, and the three faults its output carries are repaired before the
    file leaves."""
    import tempfile

    from docx_redline import compare_documents

    original_bytes = without_content_controls(original_bytes)
    revised_bytes = without_content_controls(revised_bytes)
    with tempfile.TemporaryDirectory() as tmp:
        a, b, c = f"{tmp}/a.docx", f"{tmp}/b.docx", f"{tmp}/c.docx"
        pathlib.Path(a).write_bytes(original_bytes)
        pathlib.Path(b).write_bytes(revised_bytes)
        red, _ = compare_documents(a, b, author=author)
        red.save(c)
        marked = _single_spaced(pathlib.Path(c).read_bytes(), original_bytes, revised_bytes, author)
        marked = _no_added_breaks(marked, revised_bytes)
        marked = _struck_equations(marked, revised_bytes, author)
        return unique_paragraph_ids(_valid_settings(marked))


def _struck_equations(red_bytes: bytes, revised_bytes: bytes, author: str) -> bytes:
    """The compare reads no equation, so it never strikes one: Word would show the old formula
    live beside the new. An equation is struck, as Word marks it (every math run in its own
    deletion, every structure's control properties marked deleted), where the drafter took it
    out: its paragraph struck whole by the compare (the mark struck, none of its text left), or
    its paragraph, known by its id, holding no equation in the drafter's document. The compare
    then reads that paragraph as empty and keeps it, writing the new words as a paragraph of
    their own: the emptied paragraph's mark is struck too."""
    kept = {
        _para_id(p): bool(outer_equations(p))
        for p in Document(io.BytesIO(revised_bytes)).element.body.iter(qn("w:p"))
    }
    doc = Document(io.BytesIO(red_bytes))
    ids = [int(x) for x in doc.element.body.xpath("//w:ins/@w:id | //w:del/@w:id") if x.isdigit()]
    n = max(ids, default=0) + 1
    when = datetime.datetime.now(datetime.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")

    def deletion():
        nonlocal n
        d = doc.element.body.makeelement(
            qn("w:del"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): when}
        )
        n += 1
        return d

    for p in doc.element.body.iter(qn("w:p")):
        equations = outer_equations(p)
        if not equations:
            continue
        mark = p.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}")
        live = any(
            (t.text or "").strip() and not any(True for _ in t.iterancestors(qn("w:del")))
            for t in p.iter(qn("w:t"))
        )
        struck_whole = mark is not None and not live
        pid = _para_id(p)
        if not struck_whole and (not pid or kept.get(pid, True)):
            continue
        for eq in equations:
            for r in list(eq.iter(MATH + "r")):
                if not any(True for _ in r.iterancestors(qn("w:del"))):
                    d = deletion()
                    r.addprevious(d)
                    d.append(r)
            for ctrl in eq.iter(MATH + "ctrlPr"):
                rpr = ctrl.find(qn("w:rPr"))
                if ctrl.find(qn("w:del")) is None:
                    d = deletion()
                    if rpr is not None:
                        d.append(rpr)
                    ctrl.append(d)
        if mark is None and not live:
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = p.makeelement(qn("w:pPr"), {})
                p.insert(0, ppr)
            rpr = ppr.find(qn("w:rPr"))
            if rpr is None:
                rpr = ppr.makeelement(qn("w:rPr"), {})
                ppr.append(rpr)
            rpr.insert(0, deletion())
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _no_added_breaks(red_bytes: bytes, revised_bytes: bytes) -> bytes:
    """The compare, marking a changed word just before a line break, put a second break into
    its inserted run at the paragraph's end (in a table cell: a blank line under the two
    lines the drafter wrote). A break the compare inserted as the paragraph's last mark
    goes, where the drafter's paragraph of that id reads the same and does not end in one."""
    revised: dict[str, list] = {}
    for p in Document(io.BytesIO(revised_bytes)).element.body.iter(qn("w:p")):
        revised.setdefault(_para_id(p), []).append(p)
    doc = Document(io.BytesIO(red_bytes))
    for p in doc.element.body.iter(qn("w:p")):
        mine = revised.get(_para_id(p), [])
        if len(mine) != 1 or _ends_in_break(mine[0]):
            continue  # an id the compare gave, or two of the drafter's share it
        if _squeezed(_spacing_text(p, live_only=True)) != _squeezed(_spacing_text(mine[0])):
            continue
        last = _last_live_mark(p)
        if last is not None and last.tag == qn("w:br") and any(True for _ in last.iterancestors(qn("w:ins"))):
            last.getparent().remove(last)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _last_live_mark(p):
    """The last text or line break of a paragraph not struck, or None."""
    marks = [
        x
        for x in p.iter(qn("w:t"), qn("w:br"), qn("w:tab"))
        if not any(True for _ in x.iterancestors(qn("w:del")))
    ]
    return marks[-1] if marks else None


def _ends_in_break(p) -> bool:
    last = _last_live_mark(p)
    return last is not None and last.tag == qn("w:br")


def _single_spaced(red_bytes: bytes, original_bytes: bytes, revised_bytes: bytes, author: str) -> bytes:
    """The compare reads spacing as no change, so a paragraph the drafter rewrote kept the
    form's spacing where the drafter's differs: a double space ("Supplier,  Customer"), a space
    beside a bracket or a comma ("3.4(a)( v )", "goods , services"). In each paragraph whose
    words the drafter changed, the spacing is made the drafter's: a space or tab the drafter's
    text does not have is struck, one it has is put in, each as a tracked change. Paragraphs
    are matched by their ids, which the compare keeps."""
    originals = {_para_id(p): _spacing_text(p) for p in Document(io.BytesIO(original_bytes)).element.body.iter(qn("w:p"))}
    revised = {_para_id(p): _spacing_text(p) for p in Document(io.BytesIO(revised_bytes)).element.body.iter(qn("w:p"))}
    # a paragraph the compare gave an id of its own is known by its words, spacing aside,
    # where only one paragraph of the drafter's reads so
    by_words: dict[str, set[str]] = {}
    for text in revised.values():
        by_words.setdefault(_squeezed(text), set()).add(text)
    doc = Document(io.BytesIO(red_bytes))
    n = 96000
    for p in doc.element.body.iter(qn("w:p")):
        pid, words = _para_id(p), _squeezed(_spacing_text(p, live_only=True))
        new, old = revised.get(pid), originals.get(pid)
        if new is None or _squeezed(new) != words:
            # an id the compare gave it, or one another paragraph of the drafter's has
            same = by_words.get(words, set())
            new, old = (next(iter(same)) if len(same) == 1 else None), None
        if new is not None and new != old:
            n = _respace(p, new, n, author)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _spacing_text(p, live_only: bool = False) -> str:
    """A paragraph's text with its tabs, as the runs hold them: text and tab characters of its
    runs in order; with `live_only`, those of runs not struck."""
    out = []
    for r in p.iter(qn("w:r")):
        if live_only and any(a.tag == qn("w:del") for a in r.iterancestors(qn("w:del"), qn("w:p"))):
            continue
        for part in r:
            if part.tag == qn("w:t"):
                out.append(part.text or "")
            elif part.tag == qn("w:tab"):
                out.append("\t")
            elif part.tag in (qn("w:fldChar"), qn("w:instrText")):
                out.append(_FIELD_MARK)
            elif part.tag != qn("w:rPr"):
                out.append("\x00")  # a line break, a footnote: not spacing
    return "".join(out)


_FIELD_MARK = "\x01"  # a field's mark in spacing text: not spacing, and not a word


def _squeezed(text: str) -> str:
    """Spacing text's words alone: spaces, tabs and the marks of other things set aside."""
    return re.sub(r"[\s\x00\x01]", "", text)


def _is_space(c: str) -> bool:
    """A space between words: a plain or a non-breaking one, never a tab or a line end."""
    return c.isspace() and c not in "\t\n\r\x0b\x0c"


def _respace(p, target: str, n: int, author: str) -> int:
    """`p`'s live spacing made `target`'s, where the two differ in spacing only, between two
    words only and in spaces only: a space `target` does not have is struck, one it has is put
    in, as tracked changes (inside the compare's own insertion, simply dropped or typed). A
    space is a plain or a non-breaking one. A tab, a run of tabs, and any space beside a tab
    or a line break are layout and stay, as does the space touching a field's mark (spaces
    past it are spacing); so does spacing before the first word and after the last. Returns
    the next free revision id."""
    stamp = "2026-01-01T00:00:00Z"
    now = _spacing_text(p, live_only=True)
    if now == target or _squeezed(now) != _squeezed(target):
        return n

    def gaps(text: str) -> list[tuple[int, int]]:
        """Between each two neighbouring characters of the words, the stretch between them."""
        at = [i for i, c in enumerate(text) if not c.isspace() and c not in ("\x00", _FIELD_MARK)]
        return [(a + 1, b) for a, b in zip(at, at[1:], strict=False)]

    strike: set[int] = set()
    put: dict[int, str] = {}
    for (a, b), (c, d) in zip(gaps(now), gaps(target), strict=True):
        mine, theirs = now[a:b], target[c:d]
        # spaces between two words only: a tab, a run of tabs or a line break beside them is
        # layout, and stays as the form has it
        if mine == theirs or not all(_is_space(x) or x == _FIELD_MARK for x in mine + theirs):
            continue
        # a field's marks split the spaces. A space touching a mark stays as the form has it;
        # those past it are spacing
        ours, others = mine.split(_FIELD_MARK), theirs.split(_FIELD_MARK)
        if len(ours) != len(others):
            continue
        start = a
        for k, (s, t) in enumerate(zip(ours, others, strict=True)):
            left, right = int(k > 0), int(k < len(ours) - 1)
            if (
                s != t
                and min(len(s), len(t)) >= left + right
                and s[:left] == t[:left]
                and s[len(s) - right :] == t[len(t) - right :]
            ):
                ms, mt = s[left : len(s) - right], t[left : len(t) - right]
                same = next((i for i, (x, y) in enumerate(zip(ms, mt, strict=False)) if x != y), min(len(ms), len(mt)))
                strike.update(range(start + left + same, start + left + len(ms)))
                if mt[same:]:
                    put[start + left + len(ms)] = mt[same:]
            start += len(s) + 1
    if not strike and not put:
        return n

    def tracked(kind: str, run):
        nonlocal n
        n += 1
        if kind == "del":
            _as_deleted(run)
        wrap = run.makeelement(qn(f"w:{kind}"), {qn("w:id"): str(n), qn("w:author"): author, qn("w:date"): stamp})
        wrap.append(run)
        return wrap

    def run_of(r, pieces: list):
        run = r.makeelement(qn("w:r"), dict(r.attrib))
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        for piece in pieces:
            if isinstance(piece, str):
                if piece == "\t":
                    run.append(run.makeelement(qn("w:tab"), {}))
                else:
                    t = run.makeelement(qn("w:t"), {})
                    t.text = piece
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    run.append(t)
            else:
                run.append(copy.deepcopy(piece))
        return run

    at = 0  # index into `now` of the run's next live character
    for r in list(p.iter(qn("w:r"))):
        if any(a.tag == qn("w:del") for a in r.iterancestors(qn("w:del"), qn("w:p"))):
            continue
        inserted = r.getparent().tag == qn("w:ins")
        # the run as (kind, piece) in order: kept, struck or put in
        pieces: list[tuple[str, object]] = []
        for part in r:
            if part.tag == qn("w:rPr"):
                continue
            if part.tag == qn("w:t"):
                for c in part.text or "":
                    for x in put.get(at, ""):
                        pieces.append(("ins", x))
                    pieces.append(("del" if at in strike else "keep", c))
                    at += 1
            elif part.tag == qn("w:tab"):
                pieces.append(("keep", "\t"))
                at += 1
            else:
                pieces.append(("keep", part))
                at += 1
        if all(kind == "keep" for kind, _ in pieces):
            continue
        out: list = []
        for kind, piece in pieces:
            if inserted and kind == "del":
                continue  # the compare's own insertion: the space is simply not typed
            if inserted:
                kind = "keep"
            if out and out[-1][0] == kind and isinstance(piece, str) and isinstance(out[-1][1][-1], str) and piece != "\t" and out[-1][1][-1] != "\t":
                out[-1][1][-1] += piece
            elif out and out[-1][0] == kind:
                out[-1][1].append(piece)
            else:
                out.append((kind, [piece]))
        for kind, group in out:
            run = run_of(r, group)
            r.addprevious(run if kind == "keep" else tracked(kind, run))
        r.getparent().remove(r)
    return n

def _valid_settings(docx_bytes: bytes) -> bytes:
    """docx-redline switches tracking on with an element the schema does not have
    (w:trackChanges); Word wants w:trackRevisions, in its place in the settings order."""
    import zipfile

    from lxml import etree

    name = "word/settings.xml"
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if name not in z.namelist():
            return docx_bytes
        root = etree.fromstring(z.read(name))
        stray = root.findall(f"{{{w}}}trackChanges")
        if not stray:
            return docx_bytes
        for e in stray:
            root.remove(e)
        if root.find(f"{{{w}}}trackRevisions") is None:
            track = etree.Element(f"{{{w}}}trackRevisions")
            # The schema places trackRevisions early in settings, before the tab and
            # compatibility blocks every Word document carries.
            anchor = root.find(f"{{{w}}}defaultTabStop")
            anchor = anchor if anchor is not None else root.find(f"{{{w}}}compat")
            if anchor is not None:
                anchor.addprevious(track)
            else:
                root.append(track)
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == name:
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                zo.writestr(item, data)
        return out.getvalue()


def annotate(
    redline_bytes: bytes,
    notes: list,
    author: str = "Drafting engine",
    unplaced: list[tuple[str, str]] | None = None,
) -> bytes:
    """Attach each review comment, (anchor, comment) or (anchor, comment, piece), to the
    paragraph reading as `anchor` in the contract as drafted: one reading exactly so first, else
    one holding its opening words, else a struck paragraph that read so; a table row by its
    cells. Of several, the one nearest the comments of its piece placed before it (else after
    it). A comment whose paragraph the contract no longer holds (struck, or worded otherwise
    since) goes on the nearest paragraph of its own piece that it does hold, naming the words
    it was about; one with nothing of its piece to go on is appended to `unplaced`. Never on
    the contents list, a header or a footer."""
    import difflib
    import tempfile

    from docx.text.paragraph import Paragraph
    from docx_redline import Redliner

    def struck_row(el) -> bool:
        return any(
            tr.find(f"{qn('w:trPr')}/{qn('w:del')}") is not None
            for tr in [el, *el.iterancestors(qn("w:tr"))]
            if tr.tag == qn("w:tr")
        )

    with tempfile.TemporaryDirectory() as tmp:
        path = f"{tmp}/r.docx"
        pathlib.Path(path).write_bytes(redline_bytes)
        red = Redliner(path, author=author, track_changes=False)
        # every place a comment can go, in document order: (paragraph, its words as accepted,
        # its struck words); a row is read by its cells and commented at its first words
        places: list[tuple[object, str, str]] = []
        for el in red.document.element.body.iter(qn("w:p"), qn("w:tr")):
            if struck_row(el):
                continue
            if el.tag == qn("w:tr"):
                cells = [
                    PARA_BREAK.join(
                        x for x in (_accepted_text(q) for q in tc.findall(qn("w:p"))) if x.strip()
                    )
                    for tc in el.findall(qn("w:tc"))
                ]
                host = next((q for q in el.iter(qn("w:p")) if q.find(f".//{qn('w:t')}") is not None), None)
                if host is not None:
                    places.append((host, _key(" | ".join(cells)), ""))
                continue
            if _style_id(el).lower().startswith("toc") or el.find(f".//{qn('w:r')}") is None:
                continue  # the contents list repeats every heading; nothing to anchor on
            struck = "".join(t.text or "" for t in el.iter(qn("w:delText")))
            places.append((el, _key(_accepted_text(el)), _key(struck)))
        # a change to a header or footer line is shown there by its tracked change; Word gives
        # it no comment, and it never goes on a body paragraph instead
        margins = {
            _key(_accepted_text(q))
            for section in red.document.sections
            for part in (section.header, section.footer, section.first_page_header,
                         section.first_page_footer, section.even_page_header,
                         section.even_page_footer)
            for q in part._element.iter(qn("w:p"))
        } | {
            # a header table's row, as the render names it: its cells joined (" | Draft dated")
            _key(" | ".join(_accepted_text(tc).strip() for tc in tr.findall(qn("w:tc"))))
            for section in red.document.sections
            for part in (section.header, section.footer, section.first_page_header,
                         section.first_page_footer, section.even_page_header,
                         section.even_page_footer)
            for tr in part._element.iter(qn("w:tr"))
        } - {""}

        def hits(anchor: str) -> list[int]:
            want = _key(anchor)
            probe = _key(" ".join(_FOOTNOTE.sub("", anchor).split())[:60])
            if not want:
                return []
            for test in (
                lambda k, s: k == want,
                lambda k, s: probe in k,
                lambda k, s: s == want or probe in s,
            ):
                found = [i for i, (_, k, s) in enumerate(places) if test(k, s)]
                if found:
                    return found
            return []

        groups: dict = {}
        pending: dict[int, list[tuple[str, str, str]]] = {}
        for note in notes:
            groups.setdefault(note[2] if len(note) > 2 else None, []).append(note)
        def follows(i: int, prior: str, k: int) -> bool:
            """Whether place i comes k places reading as it does after one reading `prior`."""
            want, seen = places[i][1], 0
            for j in range(i - 1, -1, -1):
                if places[j][1] == want:
                    seen += 1
                elif places[j][1]:
                    return places[j][1] == _key(prior) and seen == k
            return False

        for group in groups.values():
            found = [hits(note[0]) for note in group]
            chosen: list[int | None] = [f[0] if len(f) == 1 else None for f in found]
            for k, (note, f) in enumerate(zip(group, found, strict=True)):
                after = getattr(note[0], "after", None)
                fits = [i for i in f if after is not None and follows(i, *after)] if len(f) > 1 else []
                if len(fits) == 1:
                    chosen[k] = fits[0]  # its own paragraph, by what it follows in its piece
            for k, f in enumerate(found):
                if len(f) < 2 or chosen[k] is not None:
                    continue
                prev = next((chosen[j] for j in range(k - 1, -1, -1) if chosen[j] is not None), None)
                nxt = next((chosen[j] for j in range(k + 1, len(group)) if chosen[j] is not None), None)
                if prev is not None:
                    chosen[k] = min(f, key=lambda i: (i < prev, abs(i - prev)))
                elif nxt is not None:
                    chosen[k] = min(f, key=lambda i: (i > nxt, abs(i - nxt)))
                else:
                    chosen[k] = f[0]
            for k, note in enumerate(group):
                anchor, comment = note[0], note[1]
                if not comment:
                    continue
                at, text = chosen[k], comment
                if at is None and _key(anchor) in margins:
                    continue
                if at is None:
                    # the paragraph of the piece, between the first and last places its other
                    # comments found, reading most like the words the comment was about
                    span = [i for i in chosen if i is not None]
                    if not span:
                        if unplaced is not None:
                            unplaced.append((anchor, comment))
                        continue
                    want = _key(anchor)
                    at = max(
                        range(min(span), max(span) + 1),
                        key=lambda i: max(
                            difflib.SequenceMatcher(None, want, places[i][1]).ratio(),
                            difflib.SequenceMatcher(None, want, places[i][2]).ratio(),
                        ),
                    )
                # a row is placed at its first paragraph, which is a place of its own too
                at = next(i for i, x in enumerate(places) if x[0] is places[at][0])
                pending.setdefault(at, []).append((anchor, comment, text[:1000]))
        # one comment per paragraph: what two pieces (or a point and a change) say about the
        # same paragraph is one comment, in the order they were given
        for at, said in pending.items():
            try:
                red.add_comment(
                    Paragraph(places[at][0], red.document),
                    "\n\n".join(dict.fromkeys(t for _, _, t in said)),
                )
            except Exception:
                if unplaced is not None:
                    unplaced.extend((a, c) for a, c, _ in said)
        red.save(path)
        return unique_paragraph_ids(pathlib.Path(path).read_bytes())


def numbering_as_precedent(docx_bytes: bytes, form_bytes: bytes) -> bytes:
    """Each heading of the draft that reads as one of the precedent's, in the same style and the
    same place in their order, takes the precedent's numbering (its list and level), whatever
    the drafter gave it: a Part restarts where the precedent's restarts. One schedule's
    headings can be put on another's restart list, and that schedule's Parts number on from it.
    Headings only (ten words or fewer); a struck paragraph is not counted."""
    from lxml import etree

    def key(p) -> tuple[str, str] | None:
        words = " ".join("".join(t.text or "" for t in p.iter(qn("w:t"))).split())
        if not words or len(words.split()) > 10:
            return None
        ppr = p.find(qn("w:pPr"))
        style = ppr.find(qn("w:pStyle")) if ppr is not None else None
        return (re.sub(r"\W", "", words.lower()), style.get(qn("w:val")) if style is not None else "Normal")

    def struck(p) -> bool:
        return p.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}") is not None

    def list_of(num_pr, document) -> tuple[str, str]:
        """The counter a numbering reference counts on: its list's abstract and its level."""
        nid = num_pr.find(qn("w:numId"))
        lvl = num_pr.find(qn("w:ilvl"))
        value = nid.get(qn("w:val")) if nid is not None else ""
        abstract = next(
            (
                n.find(qn("w:abstractNumId")).get(qn("w:val"))
                for n in document.part.numbering_part.element.findall(qn("w:num"))
                if n.get(qn("w:numId")) == value and n.find(qn("w:abstractNumId")) is not None
            ),
            value,
        )
        return abstract, lvl.get(qn("w:val")) if lvl is not None else "0"

    form = Document(io.BytesIO(form_bytes))
    numbering: dict[tuple[str, str], list] = {}
    for p in form.element.body.iter(qn("w:p")):
        k = key(p)
        if k is not None:
            ppr = p.find(qn("w:pPr"))
            numbering.setdefault(k, []).append(ppr.find(qn("w:numPr")) if ppr is not None else None)
    doc = Document(io.BytesIO(docx_bytes))
    seen: dict[tuple[str, str], int] = {}
    changed = False
    for p in doc.element.body.iter(qn("w:p")):
        k = key(p)
        if k is None or struck(p) or k not in numbering:
            continue
        n = seen.get(k, 0)
        seen[k] = n + 1
        if n >= len(numbering[k]) or numbering[k][n] is None:
            continue
        want = numbering[k][n]
        ppr = p.get_or_add_pPr()
        have = ppr.find(qn("w:numPr"))
        if have is not None and etree.tostring(have) == etree.tostring(want):
            continue
        # a list restarted from the precedent's (the same counter, its own start) is the
        # precedent's list: running the pass again leaves the restart standing
        if have is not None and list_of(have, doc) == list_of(want, form):
            continue
        if have is not None:
            ppr.remove(have)
        ppr.insert_element_before(
            copy.deepcopy(want),
            "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
            "w:spacing", "w:ind", "w:contextualSpacing", "w:jc", "w:outlineLvl", "w:rPr",
            "w:sectPr", "w:pPrChange",
        )
        changed = True
    # a heading that opens its list in the precedent (it prints 1) opens it in the draft: the
    # draft may put earlier paragraphs on the same list, or on another list sharing its counter
    # (one schedule's Parts numbering on from another's), so from that heading the
    # list restarts
    def num_of(p) -> tuple[str, str] | None:
        ppr = p.find(qn("w:pPr"))
        num = ppr.find(qn("w:numPr")) if ppr is not None else None
        nid = num.find(qn("w:numId")) if num is not None else None
        if nid is None:
            return None
        lvl = num.find(qn("w:ilvl"))
        return nid.get(qn("w:val")), lvl.get(qn("w:val")) if lvl is not None else "0"

    def opens(number) -> bool:
        return bool(number and number.path and number.path[-1] == 1)

    # each heading is placed by the Part it sits in: the nearest unnumbered title above it
    # that the precedent also has ("Part B Service Levels"), and its place among that
    # Part's headings reading alike; a draft adding headings elsewhere does not shift them
    def titles(document) -> set[str]:
        return {
            k[0]
            for para in document.paragraphs
            for k in [key(para._p)]
            if k is not None and num_of(para._p) is None
        }

    shared_titles = titles(form)

    def placed(document):
        scope = ""
        counted: dict[tuple[str, str, str], int] = {}
        for i, para in enumerate(document.paragraphs):
            k = key(para._p)
            if k is None or struck(para._p):
                yield i, para, None
                continue
            if num_of(para._p) is None and k[0] in shared_titles:
                scope = k[0]
            place = (scope, *k)
            occurrence = counted.get(place, 0)
            counted[place] = occurrence + 1
            yield i, para, (place, occurrence)

    form_numbers = accepted_numbers_by_paragraph(form_bytes, form)
    starters = {
        where
        for i, para, where in placed(form)
        if where is not None and num_of(para._p) is not None and opens(form_numbers.get(i))
    }
    if starters:
        staged = io.BytesIO()
        doc.save(staged)
        doc = Document(io.BytesIO(staged.getvalue()))
        draft_numbers = accepted_numbers_by_paragraph(staged.getvalue(), doc)
        numbering_el = doc.part.numbering_part.element
        # a restart carries through the rest of its own Part only
        remap: dict[tuple[str, str], str] = {}
        for i, para, where in placed(doc):
            p = para._p
            num = num_of(p)
            if num is None or struck(p):
                continue
            nid, lvl = num
            target = p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
            scope = where[0][0] if where is not None else ""
            if (scope, nid) in remap:
                target.set(qn("w:val"), remap[(scope, nid)])
                continue
            if where is None or where not in starters or opens(draft_numbers.get(i)):
                continue
            source = next(
                (n for n in numbering_el.findall(qn("w:num")) if n.get(qn("w:numId")) == nid), None
            )
            abstract = source.find(qn("w:abstractNumId")) if source is not None else None
            if abstract is None:
                continue
            fresh = str(1 + max(int(n.get(qn("w:numId"))) for n in numbering_el.findall(qn("w:num"))))
            num_el = numbering_el.makeelement(qn("w:num"), {qn("w:numId"): fresh})
            num_el.append(num_el.makeelement(qn("w:abstractNumId"), {qn("w:val"): abstract.get(qn("w:val"))}))
            override = num_el.makeelement(qn("w:lvlOverride"), {qn("w:ilvl"): lvl})
            override.append(override.makeelement(qn("w:startOverride"), {qn("w:val"): "1"}))
            num_el.append(override)
            numbering_el.append(num_el)
            remap[(scope, nid)] = fresh
            target.set(qn("w:val"), fresh)
            changed = True
    if not changed:
        return docx_bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def remove_highlighting(docx_bytes: bytes) -> bytes:
    """The precedent's highlighting taken out of the draft, untracked: a form marks what to
    check (a pro forma, what changes per deal) and the draft's own changes show in the
    redline alone. Body, headers, footers and notes."""
    doc = Document(io.BytesIO(docx_bytes))
    parts = [doc.part, *(rel.target_part for rel in doc.part.rels.values() if not rel.is_external)]
    for part in parts:
        root = getattr(part, "element", None)
        if root is None:
            continue
        for mark in list(root.iter(qn("w:highlight"))):
            mark.getparent().remove(mark)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def ask_to_update_fields(docx_bytes: bytes) -> bytes:
    """Word refreshes the contents list and every field when the lawyer opens the draft: the
    contents the form cached name schedules and appendices this contract no longer has."""
    doc = Document(io.BytesIO(docx_bytes))
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        flag = settings.makeelement(qn("w:updateFields"), {qn("w:val"): "true"})
        # its place in the schema's order: before the first of the elements that follow it
        after = {
            "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
            "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
            "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
            "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
            "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
        }
        from lxml import etree

        later = next((c for c in settings if etree.QName(c).localname in after), None)
        if later is None:
            settings.append(flag)
        else:
            later.addprevious(flag)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def stamp_properties(docx_bytes: bytes, author: str, day: datetime.datetime) -> bytes:
    """The file's properties name the house and the day of drafting, not the form's authors."""
    doc = Document(io.BytesIO(docx_bytes))
    props = doc.core_properties
    props.author = author
    props.last_modified_by = author
    props.created = day
    props.modified = day
    props.revision = 1
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def accept_all_changes(docx_bytes: bytes) -> bytes:
    """The document as it reads with every tracked change accepted: inserted runs become plain
    runs, deleted runs go, revision marks on paragraph and run properties go. Word structure
    only; no word of the text is decided here."""
    doc = Document(io.BytesIO(docx_bytes))
    roots = [doc.element.body]
    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            if part is not None and part._element is not None:
                roots.append(part._element)
    # the notes are parts of their own, which the parser may hold only as bytes: a struck
    # form footnote left unaccepted comes back as a deletion in a later mark-up
    from lxml import etree

    notes = []
    for rel in doc.part.rels.values():
        if rel.is_external or not rel.reltype.endswith(("/footnotes", "/endnotes")):
            continue
        part = rel.target_part
        tree = getattr(part, "_element", None)
        notes.append((part, tree is None, tree if tree is not None else etree.fromstring(part.blob)))
        roots.append(notes[-1][2])
    for root in roots:
        # a paragraph struck whole (its mark deleted and no live words) goes, and so does a
        # struck row: left as empty numbered paragraphs, every later clause and schedule
        # number in the accepted document would count them
        struck = [
            p
            for p in root.iter(qn("w:p"))
            if p.find(f"{qn('w:pPr')}/{qn('w:rPr')}/{qn('w:del')}") is not None
            and not any(
                not any(a.tag == qn("w:del") for a in t.iterancestors()) and (t.text or "").strip()
                for t in p.iter(qn("w:t"))
            )
        ]
        rows = [
            tr for tr in root.iter(qn("w:tr")) if tr.find(f"{qn('w:trPr')}/{qn('w:del')}") is not None
        ]
        for el in [*struck, *rows]:
            parent = el.getparent()
            # a cell or a note keeps one paragraph: Word will not open one without
            keeps = (qn("w:tc"), qn("w:footnote"), qn("w:endnote"))
            if parent is not None and not (el.tag == qn("w:p") and parent.tag in keeps):
                parent.remove(el)
        for el in list(root.iter(qn("w:del"))):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
        for el in list(root.iter(qn("w:ins"))):
            parent = el.getparent()
            if parent is None:
                continue
            at = parent.index(el)
            for child in list(el):
                parent.insert(at, child)
                at += 1
            parent.remove(el)
        for tag in (
            "w:rPrChange",
            "w:pPrChange",
            "w:sectPrChange",
            "w:tblPrChange",
            "w:trPrChange",
            "w:tcPrChange",
        ):
            for el in list(root.iter(qn(tag))):
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
    for part, as_bytes, tree in notes:
        if as_bytes:
            part._blob = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def repair_part_names(docx_bytes: bytes) -> bytes:
    """A document whose relationships name a part in a different letter case from the part
    itself ("customXML/item2.xml" for "customXml/item2.xml"). Word opens it; the parser does
    not. The case is put right in the relationship files and content types, nothing else; a
    document with nothing to repair comes back unchanged."""
    import zipfile

    src = zipfile.ZipFile(io.BytesIO(docx_bytes))
    if not any(
        b"customXML" in src.read(n)
        for n in src.namelist()
        if n.endswith((".rels", ".xml")) and ("_rels" in n or n == "[Content_Types].xml")
    ):
        return docx_bytes
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in src.namelist():
            data = src.read(name)
            if name.endswith(".rels") or name == "[Content_Types].xml":
                data = data.replace(b"customXML", b"customXml")
            dst.writestr(name, data)
    return out.getvalue()


def normalise_hyphens(docx_bytes: bytes) -> bytes:
    """Word's non-breaking hyphen is an element, not a character, and the render drops it, so
    "AB12-3C" reads "AB123C" to every agent and nothing they quote matches the document.
    Each one becomes a hyphen character in its run. Structure only; a document with none comes
    back unchanged."""
    import zipfile

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if b"noBreakHyphen" not in z.read("word/document.xml"):
            return docx_bytes
    doc = Document(io.BytesIO(docx_bytes))
    for el in list(doc.element.body.iter(qn("w:noBreakHyphen"))):
        t = el.makeelement(qn("w:t"), {})
        t.text = "-"
        el.addprevious(t)
        el.getparent().remove(el)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
