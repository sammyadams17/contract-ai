"""Compute Word's effective clause numbers for every numbered paragraph.

Why this exists: a clause's number ("3.4(a)(iii)") is not text in the document.
Word derives it from `numbering.xml` plus each paragraph's `w:numPr`, and stores
the result only where a cross-reference field happens to have cached it. adeu
does not reconstruct it - `_get_numbering_cache` reads `numFmt` purely to choose
between a bullet and an ordered marker, and its projection emits a generic "1. "
because Markdown renumbers anyway. Mining cross-references therefore recovers
numbers for only the clauses something references, a small fraction of a template's.

This walks the document once and computes all of them, which is what lets an
agent or a reviewer address any clause by the number a lawyer would use.

Validated against Word's own answers rather than against expectation:
`cached_clause_numbers` reads the numbers Word stored in cross-reference fields,
with the switch that says which rendering each one is, so every computed number
a field cached can be checked against the number Word itself showed.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field

from docx.oxml.ns import qn

_ROMAN = [
    (1000, "m"),
    (900, "cm"),
    (500, "d"),
    (400, "cd"),
    (100, "c"),
    (90, "xc"),
    (50, "l"),
    (40, "xl"),
    (10, "x"),
    (9, "ix"),
    (5, "v"),
    (4, "iv"),
    (1, "i"),
]


@dataclass
class LevelDefinition:
    number_format: str = "decimal"
    level_text: str = "%1."
    start: int = 1
    # A numbered clause's indent usually lives here, in the level definition,
    # not in the paragraph's own pPr - so anything matching a clause's visual
    # formatting without joining its numbering needs it. In twips.
    indent_left: int | None = None


@dataclass
class NumberingInstance:
    """One `w:num`: a reference to an abstractNum plus this instance's overrides.

    `abstract_id` is the counter's identity. Several numIds routinely point at one
    abstractNum and Word counts them as a single sequence, so the numId is only how
    a paragraph names the instance - it is not what is being counted.

    `restarts` holds this instance's `startOverride` values, `{ilvl: start}`. Word
    applies one when the instance is first used at that level, which is how a contract
    restarts numbering in each Schedule.
    """

    abstract_id: str
    levels: dict[int, LevelDefinition] = field(default_factory=dict)
    restarts: dict[int, int] = field(default_factory=dict)


@dataclass
class NumberingDefinitions:
    """`{numId: NumberingInstance}`, resolved through abstractNum indirection."""

    by_num_id: dict[str, NumberingInstance] = field(default_factory=dict)

    def level(self, num_id: str, ilvl: int) -> LevelDefinition | None:
        instance = self.by_num_id.get(num_id)
        return None if instance is None else instance.levels.get(ilvl)


def load_numbering_definitions(docx_bytes: bytes) -> NumberingDefinitions:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as archive:
        if "word/numbering.xml" not in archive.namelist():
            return NumberingDefinitions()
        xml = archive.read("word/numbering.xml").decode("utf-8", errors="ignore")

    abstract: dict[str, dict[int, LevelDefinition]] = {}
    for block in re.findall(r"<w:abstractNum\b.*?</w:abstractNum>", xml, re.S):
        abstract_id = _attr(block, "w:abstractNumId")
        if abstract_id is None:
            continue
        levels: dict[int, LevelDefinition] = {}
        for level_block in re.findall(r"<w:lvl\b.*?</w:lvl>", block, re.S):
            ilvl = _attr(level_block, "w:ilvl")
            if ilvl is None or not ilvl.lstrip("-").isdigit():
                continue
            levels[int(ilvl)] = LevelDefinition(
                number_format=_child_val(level_block, "w:numFmt") or "decimal",
                level_text=_child_val(level_block, "w:lvlText") or "%1.",
                start=int(_child_val(level_block, "w:start") or 1),
                indent_left=_indent_left(level_block),
            )
        abstract[abstract_id] = levels

    definitions = NumberingDefinitions()
    for block in re.findall(r"<w:num\b(?![a-zA-Z]).*?</w:num>", xml, re.S):
        num_id = _attr(block, "w:numId")
        abstract_ref = _child_val(block, "w:abstractNumId")
        if not num_id or abstract_ref not in abstract:
            continue
        # Copy, never share: many numIds point at one abstractNum, and applying
        # this numId's overrides to a shared dict would corrupt the others.
        levels = {
            ilvl: LevelDefinition(d.number_format, d.level_text, d.start, d.indent_left)
            for ilvl, d in abstract[abstract_ref].items()
        }
        restarts: dict[int, int] = {}
        for override in re.findall(
            r"<w:lvlOverride\b.*?</w:lvlOverride>|<w:lvlOverride\b[^>]*/>", block, re.S
        ):
            ilvl_raw = _attr(override, "w:ilvl")
            start_raw = _child_val(override, "w:startOverride")
            if ilvl_raw and start_raw and ilvl_raw.lstrip("-").isdigit():
                ilvl = int(ilvl_raw)
                if ilvl in levels:
                    levels[ilvl].start = int(start_raw)
                    restarts[ilvl] = int(start_raw)
        definitions.by_num_id[num_id] = NumberingInstance(abstract_ref, levels, restarts)
    return definitions


def _attr(block: str, name: str) -> str | None:
    match = re.search(rf'{name}="([^"]*)"', block)
    return match.group(1) if match else None


def _child_val(block: str, tag: str) -> str | None:
    match = re.search(rf'<{tag}\b[^>]*w:val="([^"]*)"', block)
    return match.group(1) if match else None


def _indent_left(block: str) -> int | None:
    match = re.search(r'<w:ind\b[^>]*w:left="(-?\d+)"', block)
    return int(match.group(1)) if match else None


def format_counter(value: int, number_format: str) -> str:
    if number_format == "lowerLetter":
        return _alpha(value).lower()
    if number_format == "upperLetter":
        return _alpha(value).upper()
    if number_format == "lowerRoman":
        return _roman(value)
    if number_format == "upperRoman":
        return _roman(value).upper()
    if number_format in ("none", "bullet"):
        return ""
    return str(value)


def _alpha(value: int) -> str:
    letters = ""
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters = chr(ord("a") + remainder) + letters
    return letters or "a"


def _roman(value: int) -> str:
    out = ""
    for amount, numeral in _ROMAN:
        while value >= amount:
            out += numeral
            value -= amount
    return out


@dataclass(frozen=True)
class ClauseNumber:
    """The renderings Word caches, which are not interchangeable.

    `full` is the number in full context, what a `REF \\w` field shows and what a
    lawyer means by "Clause 3.4(a)(iii)". `marker` is the level's own text, what
    the paragraph displays and what a `REF \\n` field shows - "(iii)". A `REF \\r`
    field shows the number as seen from where the field stands - see `relative_to`.

    `path` is the counter at each level down to this one, `None` where the document
    skipped a level, and `relative_forms[k]` the number with its first `k` levels
    taken as given.
    """

    full: str
    marker: str
    list_id: str = ""
    path: tuple[int | None, ...] = ()
    relative_forms: tuple[str, ...] = ()

    @property
    def local(self) -> str:
        """The number within its own Clause or Schedule: "1.3(d)" for paragraph
        1.3(d) of Schedule 3, whose full number is "Schedule 31.3(d)"."""
        return (
            self.relative_forms[min(1, len(self.relative_forms) - 1)]
            if self.relative_forms
            else self.full
        )

    @property
    def cited(self) -> str:
        """The number as a lawyer cites it, each level Word runs together set apart:
        "Schedule 10, Part B, 3(b)" for the full number "Schedule 10Part B3.(b)". A level's
        own text is what its relative form adds to the next one's."""
        forms = self.relative_forms or (self.full,)
        out = ""
        for k, form in enumerate(forms):
            deeper = forms[k + 1] if k + 1 < len(forms) else ""
            if not form.endswith(deeper):
                return self.full
            own = form[: len(form) - len(deeper)].strip().rstrip(".").strip()
            if own:
                out += own if not out or own.startswith("(") else f", {own}"
        return out or self.full

    def relative_to(self, where: ClauseNumber | None) -> str:
        """The number seen from `where`, the latest paragraph numbered in this list
        before the reference: the levels the two share are left out, so 5.4(d) is
        "(d)" from 5.4(b) and "5.4(d)" from 5.2."""
        if where is None or where.list_id != self.list_id or not self.relative_forms:
            return self.full
        shared = 0
        while shared < len(self.path) - 1 and _at(where.path, shared) == self.path[shared]:
            shared += 1
        return self.relative_forms[shared] or self.full


def _at(path: tuple[int | None, ...], level: int) -> int | None:
    return path[level] if level < len(path) else None


class ReferencePosition:
    """Where a cross-reference stands in each list, which is what `REF \\r` is
    relative to. Pass it every paragraph's number in document order."""

    def __init__(self) -> None:
        self._latest: dict[str, ClauseNumber] = {}

    def passed(self, number: ClauseNumber | None) -> None:
        if number is not None:
            self._latest[number.list_id] = number

    def render(self, number: ClauseNumber, switch: str) -> str:
        if switch == "n":
            return number.marker
        if switch == "r":
            return number.relative_to(self._latest.get(number.list_id))
        return number.full


class ClauseNumberer:
    """Walks paragraphs in document order, maintaining Word's per-level counters.

    Counters belong to the abstractNum, not the numId: Word treats every numbering
    instance over one abstractNum as a single sequence, and a `startOverride`
    restarts that shared sequence when its instance is first used. Both halves
    matter, and they are the same mechanism seen from two ends - it is how a
    contract numbers headings and body items from one outline over two numIds,
    and how a contract restarts numbering in each Schedule.

    A level's `lvlText` may already include its ancestors ("%1.%2") or not
    ("(%3)"), so `full` joins the templates of the levels no deeper template
    names - see `_compose`.
    """

    def __init__(self, definitions: NumberingDefinitions):
        self.definitions = definitions
        self._counters: dict[str, dict[int, int]] = {}
        self._restarted: set[tuple[str, int]] = set()

    def advance(self, num_id: str | None, ilvl: int | None) -> ClauseNumber | None:
        if num_id is None or ilvl is None:
            return None
        instance = self.definitions.by_num_id.get(num_id)
        if instance is None or ilvl not in instance.levels:
            return None

        levels = instance.levels
        counters = self._counters.setdefault(instance.abstract_id, {})
        self._apply_restart(num_id, ilvl, instance, counters)
        counters[ilvl] = counters.get(ilvl, levels[ilvl].start - 1) + 1
        _clear_deeper_than(counters, ilvl)  # a new item at this level restarts its children

        rendered = {
            level: format_counter(counters.get(level, definition.start), definition.number_format)
            for level, definition in levels.items()
            if level <= ilvl
        }
        forms = tuple(
            _clean(_compose(levels, counters, rendered, first, ilvl)) for first in range(ilvl + 1)
        )
        if not forms[0]:
            return None
        marker = _clean(_substitute(levels[ilvl].level_text, rendered))
        path = tuple(counters.get(level) for level in range(ilvl + 1))
        return ClauseNumber(forms[0], marker, instance.abstract_id, path, forms)

    def _apply_restart(
        self,
        num_id: str,
        ilvl: int,
        instance: NumberingInstance,
        counters: dict[int, int],
    ) -> None:
        """A startOverride takes effect the first time its instance is used at that
        level, and only then - afterwards the instance counts on normally."""
        start = instance.restarts.get(ilvl)
        if start is None or (num_id, ilvl) in self._restarted:
            return
        self._restarted.add((num_id, ilvl))
        counters[ilvl] = start - 1
        _clear_deeper_than(counters, ilvl)


def _clean(number: str) -> str:
    """A level text carries typography as well as counters. The non-breaking space
    in "Schedule\xa01" keeps Word from splitting the label over a line; it is not
    part of the number, and it must not be what an agent cites or matches on.
    """
    return number.replace("\xa0", " ").strip().rstrip(".")


def _clear_deeper_than(counters: dict[int, int], ilvl: int) -> None:
    for deeper in [key for key in counters if key > ilvl]:
        del counters[deeper]


def _compose(
    levels: dict[int, LevelDefinition],
    counters: dict[int, int],
    rendered: dict[int, str],
    first: int,
    ilvl: int,
) -> str:
    """The number from level `first` down to `ilvl`, as Word composes it: each
    level's own text, except a level the document skipped, which Word does not
    display, and a level a deeper displayed one already names ("%1.%2" names
    level 0)."""
    named: set[int] = set()
    parts: list[str] = []
    for level in range(ilvl, first - 1, -1):
        definition = levels.get(level)
        if definition is None or level in named or (level != ilvl and level not in counters):
            continue
        parts.append(_substitute(definition.level_text, rendered))
        named |= {int(n) - 1 for n in re.findall(r"%(\d)", definition.level_text)}
    return "".join(reversed(parts))


def _substitute(level_text: str, rendered: dict[int, str]) -> str:
    def replace(match: re.Match) -> str:
        level = int(match.group(1)) - 1
        return rendered.get(level, "")

    return re.sub(r"%(\d)", replace, level_text)


@dataclass(frozen=True)
class CachedReference:
    """A clause number Word itself computed, cached in a REF field.

    `switch` decides what the number means and so what it may be compared against:
    `n` is the target's own marker, `w` is its number in full context, `r` its
    number in relative context. It is the closest thing to ground truth available
    offline - subject to being only as fresh as the last time Word updated fields.
    `paragraph_index` is the target's paragraph, `field_index` the field's own.
    """

    anchor: str
    switch: str
    number: str
    paragraph_index: int | None = None
    field_index: int | None = None


_REF_FIELD = re.compile(r"\bREF\s+(_Ref\w+)((?:\s+\\[a-zA-Z*]+)*)")
# Cross-references are authored inside a sentence ("in paragraph (ii) above"), and
# the trailing word is often typed into the field result rather than after it.
_TRAILING_DRAFTING_WORD = re.compile(r"\s+(above|below)\s*$", re.I)


def cached_clause_numbers(docx_bytes: bytes, document=None) -> list[CachedReference]:
    """Every clause number Word has cached in a numeric REF field, in document
    order. Text references and PAGEREFs are excluded: they cache a heading or a
    page number, neither of which is a clause number.
    """
    from docx import Document

    document = document or Document(io.BytesIO(docx_bytes))
    anchors = _anchor_paragraphs(document)
    return [
        CachedReference(anchor, switch, number, anchors.get(anchor), index)
        for index, paragraph in enumerate(document.paragraphs)
        for anchor, switch, number in _reference_fields(paragraph)
    ]


def _reference_fields(paragraph) -> list[tuple[str, str, str]]:
    """Walk runs, pairing each field's instruction with its cached result.

    A Word field is a flat run sequence delimited by fldChar begin/separate/end
    rather than a nested element, so it has to be reassembled.
    """
    out: list[tuple[str, str, str]] = []
    instruction = result = ""
    state: str | None = None
    for node in paragraph._p.iter():
        tag = node.tag.split("}")[-1]
        if tag == "fldChar":
            state = node.get(qn("w:fldCharType"))
            if state == "begin":
                instruction = result = ""
            elif state == "end":
                reference = _numeric_reference(instruction, result)
                if reference:
                    out.append(reference)
                state = None
        elif tag == "instrText" and state == "begin":
            instruction += node.text or ""
        elif tag == "t" and state == "separate":
            result += node.text or ""
    return out


def _numeric_reference(instruction: str, result: str) -> tuple[str, str, str] | None:
    if "PAGEREF" in instruction:
        return None
    match = _REF_FIELD.search(instruction)
    if not match:
        return None
    switch = next((s for s in re.findall(r"\\([a-z])", match.group(2)) if s in "nrw"), None)
    if switch is None:
        return None  # a bare REF caches the target's text, not its number
    number = _TRAILING_DRAFTING_WORD.sub("", result.replace("\xa0", " ")).strip()
    return (match.group(1), switch, number) if number else None


def _anchor_paragraphs(document) -> dict[str, int]:
    """Index of the paragraph each `_Ref` bookmark opens in."""
    out: dict[str, int] = {}
    for index, paragraph in enumerate(document.paragraphs):
        for node in paragraph._p.iter():
            if node.tag.endswith("}bookmarkStart"):
                name = node.get(qn("w:name"))
                if name and name.startswith("_Ref"):
                    out.setdefault(name, index)
    return out


def numbers_by_paragraph(docx_bytes: bytes, document=None) -> dict[int, ClauseNumber]:
    """Effective clause number per numbered paragraph, keyed by index in document
    order."""
    from docx import Document

    document = document or Document(io.BytesIO(docx_bytes))
    numberer = ClauseNumberer(load_numbering_definitions(docx_bytes))

    out: dict[int, ClauseNumber] = {}
    for index, paragraph in enumerate(document.paragraphs):
        number = numberer.advance(*_paragraph_numbering(paragraph))
        if number is not None:
            out[index] = number
    return out


def accepted_numbers_by_paragraph(docx_bytes: bytes, document=None) -> dict[int, ClauseNumber]:
    """Clause numbers as Word shows them once a deleted paragraph mark is accepted.

    `numbers_by_paragraph` still counts a heading whose mark is deleted. That is the
    count a contract's field cache may have been written against. This one leaves those
    headings out. A `REF` field has to cache
    this count, or it names a clause Word no longer displays.
    """
    from docx import Document

    document = document or Document(io.BytesIO(docx_bytes))
    numberer = ClauseNumberer(load_numbering_definitions(docx_bytes))

    out: dict[int, ClauseNumber] = {}
    for index, paragraph in enumerate(document.paragraphs):
        if _paragraph_mark_deleted(paragraph):
            continue
        number = numberer.advance(*_paragraph_numbering(paragraph))
        if number is not None:
            out[index] = number
    return out


def _paragraph_mark_deleted(paragraph) -> bool:
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    rPr = pPr.find(qn("w:rPr"))
    return rPr is not None and rPr.find(qn("w:del")) is not None


def _paragraph_numbering(paragraph) -> tuple[str | None, int | None]:
    """Resolve numId and ilvl for a paragraph, accumulating down the style chain.

    The two parts are commonly supplied by *different* levels of that chain, so
    they must be gathered independently rather than taken from the first `numPr`
    found. A template may define a `Level1` style with `numId=6` and no `ilvl`,
    then `Level2..4` styles with only `ilvl=1..3`, inheriting the numId from their
    base style. Stopping at the first `numPr` yields `(None, 1)` and numbers
    nothing - which is why an earlier version numbered only a fraction of the
    paragraphs.
    """
    num_id: str | None = None
    ilvl: int | None = None

    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        if pPr.find(qn("w:sectPr")) is not None and not paragraph.text.strip():
            return None, None  # Word shows a section break here, and no number
        num_id, ilvl = _read_numPr(pPr.find(qn("w:numPr")))

    style = paragraph.style
    # styles told apart by their ids in the document: python-docx makes a new object for each
    # style it hands back, and a freed object's id() is given to the next, so a walk keyed on
    # id() stopped two or three styles up as though the chain had cycled
    seen: set[str] = set()
    while style is not None and (num_id is None or ilvl is None):
        if style.style_id in seen:  # defensive: malformed style graphs can cycle
            break
        seen.add(style.style_id)
        element = getattr(style, "_element", None)
        if element is not None:
            style_pPr = element.find(qn("w:pPr"))
            if style_pPr is not None:
                style_num_id, style_ilvl = _read_numPr(style_pPr.find(qn("w:numPr")))
                num_id = num_id if num_id is not None else style_num_id
                ilvl = ilvl if ilvl is not None else style_ilvl
        style = getattr(style, "base_style", None)

    if num_id is None:
        return None, None
    return num_id, ilvl if ilvl is not None else 0


def _read_numPr(numPr) -> tuple[str | None, int | None]:
    if numPr is None:
        return None, None
    num_id_el = numPr.find(qn("w:numId"))
    ilvl_el = numPr.find(qn("w:ilvl"))
    num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else None
    ilvl_raw = ilvl_el.get(qn("w:val")) if ilvl_el is not None else None
    ilvl = int(ilvl_raw) if ilvl_raw and ilvl_raw.lstrip("-").isdigit() else None
    return num_id, ilvl
