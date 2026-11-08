"""What only the whole contract shows, read once every piece has drafted.

The drafters work at the same time, each on its own piece. A definition whose only users another
piece struck looks used to the piece that holds it; a definition seven pieces each added looks
single to each of them. Code reads the
contract as it now stands and hands each such definition back to the piece that holds it. It reads
words, never meaning: what to do is still the drafter's.
"""

from __future__ import annotations

import re

_DEFINES = re.compile(
    r'^[\W_]*["“]([^"”]+)["”]\s+(?:shall\s+)?'
    r"(?:means?|has the meaning|have the meaning|bears the meaning|includes?)\b"
)


def _term(line: str) -> str:
    m = _DEFINES.match(line.strip())
    return m.group(1).strip() if m else ""


def _body(render: str) -> list[str]:
    """The render's lines before its read-only tail, whose warnings name defined terms."""
    return render.split("<!-- READONLY_BOUNDARY_START -->")[0].split("\n")


def _uses(term: str, lines: list[str]) -> int:
    """How many times the contract uses a term outside its own definitions."""
    word = re.compile(rf"(?<![\w-]){re.escape(term)}(?![\w-])")
    return sum(len(word.findall(line)) for line in lines if _term(line) != term)


def whole_contract_faults(
    form_text: str,
    contract_text: str,
    pieces: dict[str, list[tuple[str, bool]]],
    renamed: dict[str, str] | None = None,
) -> dict[str, list[str]]:
    """Faults for each piece, by title. `pieces` gives each piece's paragraphs as they now stand,
    each with whether the draft added it. A definition the form used elsewhere and the contract no
    longer uses anywhere is raised with the piece holding it; a term defined twice is raised with
    each piece that added a copy, the form's own definition, or the first, standing; a term
    `renamed` (old to new, `renamed_terms`) is raised with each piece still using its old name."""
    form = _body(form_text)
    contract = _body(contract_text)
    holders: dict[str, list[tuple[str, bool]]] = {}  # term -> (piece, added) for each definition
    for title, paras in pieces.items():
        for text, added in paras:
            term = _term(text)
            if term:
                holders.setdefault(term, []).append((title, added))
    faults: dict[str, list[str]] = {}
    for term, held in holders.items():
        # a term of one letter is a formula's symbol, used in an equation the render may not show
        if len(term) > 1 and _uses(term, contract) == 0 and _uses(term, form) > 0:
            for title, _ in held:
                faults.setdefault(title, []).append(
                    f'"{term}" is defined in this piece and used nowhere in the contract as it now'
                    " stands: the provisions that used it have been struck. Strike its definition,"
                    " unless a rule for this piece keeps it."
                )
        if len(held) > 1:
            keep = next((h for h in held if not h[1]), held[0])
            for title, added in held:
                if added and (title, added) != keep:
                    faults.setdefault(title, []).append(
                        f'"{term}" is defined twice: {keep[0]} holds its definition. Strike the'
                        " definition this piece added."
                    )
    # a term a piece renamed in its definition, still used under its old name (a Start Date
    # that became a Commencement Date); not where the contract still
    # defines the old name, in a definition or in a clause's own quotation marks
    quoted = {m.group(1).strip() for line in contract for m in re.finditer(r'["“]([^"”]+)["”]', line)}
    for old, new in (renamed or {}).items():
        if old in quoted:
            continue
        for title, paras in pieces.items():
            if _uses(old, [text for text, _ in paras]):
                faults.setdefault(title, []).append(
                    f'"{old}" is used in this piece, but the contract now defines it as "{new}":'
                    f' write "{new}" for it.'
                )
    return faults


def renamed_terms(originals: list[tuple[str, str]]) -> dict[str, str]:
    """Each definition a piece renamed, old term to new, from its paragraphs as (original,
    now): a paragraph that defined one term and now defines another. A term made singular or
    plural is not renamed: a contract uses both numbers of a defined term."""
    out = {}
    for original, now in originals:
        old, new = _term(original), _term(now)
        if old and new and old.rstrip("s") != new.rstrip("s"):
            out[old] = new
    return out
