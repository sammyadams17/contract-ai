"""References a drafter typed as words, read against the contract as it now numbers.

A Word field (`REF`) keeps itself right; a reference a drafter typed does not. Drafters type the
number the form gave a clause, or a number they worked out, while other pieces strike clauses in
parallel and the contract renumbers under them: a drafter cites "Clause 13 (Termination)" once
an earlier clause is struck and Termination has become Clause 12, or cites a "Clause
13.5(b)(i)" the contract does not have. Code only reads the contract's own numbering here: a typed
reference that names its clause by title is set to the number that title has; one that names a
number the contract does not have, or a title no clause has, is raised for the lawyer.
"""

from __future__ import annotations

import io
import re

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.word.numbering import numbers_by_paragraph

# a contract's own citation form: "Clause 14.5(b)(i)", "Schedule 11", each optionally followed by
# the title it is known by, "(Termination)"
_REFERENCE = re.compile(
    r"\b(Clause|Schedule) (\d+(?:\.\d+)*(?:\([a-z]+\))*(?:\([ivx]+\))*)(?: \(([A-Z][^()]{2,60})\))?"
)


def _words(text: str) -> str:
    return " ".join(text.split()).lower()


def _typed_runs(paragraph) -> list:
    """The paragraph's text nodes outside any field, in order: a field's result is Word's."""
    out, depth = [], 0
    for node in paragraph._p.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "fldChar":
            kind = node.get(qn("w:fldCharType"))
            depth += 1 if kind == "begin" else -1 if kind == "end" else 0
        elif tag == "t" and depth == 0:
            if not any(a.tag == qn("w:fldSimple") for a in node.iterancestors()):
                out.append(node)
    return out


def typed_references(
    docx_bytes: bytes, numbered: bytes | None = None
) -> tuple[bytes, list[tuple[str, str]]]:
    """The contract with each typed reference that names its clause by title set to that clause's
    number, and (paragraph text, comment) for each one the lawyer must check. `numbered` is the
    contract as it will number, where `docx_bytes` still holds paragraphs struck afterwards: a
    struck schedule title counted would number every schedule after it one higher."""
    doc = Document(io.BytesIO(docx_bytes))
    source = Document(io.BytesIO(numbered)) if numbered is not None else doc
    paras = source.paragraphs
    numbers = numbers_by_paragraph(numbered if numbered is not None else docx_bytes, source)
    clauses: dict[str, str] = {}  # number -> title, as the contract now numbers
    schedules: dict[str, str] = {}
    for i, n in numbers.items():
        full = n.full.strip().rstrip(".")
        # a schedule's own heading, its number all its own: paragraph 1 of Schedule 3 has the
        # full number "Schedule 31", the two run together, and is no Schedule 31
        if re.fullmatch(r"Schedule \d+", full) and n.marker.strip().rstrip(".") == full:
            schedules[full.split()[1]] = _words(paras[i].text)
        elif re.fullmatch(r"\d+(?:\.\d+)*(?:\([a-z]+\))*(?:\([ivx]+\))*", full):
            clauses.setdefault(full, _words(paras[i].text))
    # a schedule heading the contract types ("SCHEDULE 7", its title after it or in the next
    # paragraph) rather than letting Word number it: the first, past the contents page, whose
    # entries end in a tab and the page number; a form's own list of its attachments
    # ("Schedule 1: Service Levels") comes later and is not the contract's
    for i, p in enumerate(paras):
        if re.search(r"\t\d+\s*$", p.text):
            continue
        m = re.fullmatch(r"schedule\s+(\d+)\s*(.*)", " ".join(p.text.split()), re.I)
        if m and m.group(1) not in schedules:
            schedules[m.group(1)] = _words(
                m.group(2) or next((q.text for q in paras[i + 1 :] if q.text.strip()), "")
            )

    notes: list[tuple[str, str]] = []
    changed = _misdirected_fields(doc, schedules)
    for p in doc.paragraphs:
        nodes = _typed_runs(p)
        text = "".join(t.text or "" for t in nodes)
        # each match's number, mapped back to the text node it sits in
        edits: list[tuple[int, int, str]] = []
        for m in _REFERENCE.finditer(text):
            kind, number, title = m.groups()
            known = clauses if kind == "Clause" else schedules
            if number not in known:
                notes.append((p.text, f"Cross-reference to check: this contract has no {kind} {number}."))
                continue
            if not title:
                continue
            chain = [number[:j] for j in range(len(number), 0, -1) if number[:j] in known]
            # a title may be the heading's opening words, as the precedent itself shortens
            # "Representations and Warranties" to "Representations"
            if any((known[c] + " ").startswith(_words(title) + " ") for c in chain):
                continue
            depth = number.count(".") + number.count("(")
            right = [c for c, t in known.items() if t == _words(title) and c.count(".") + c.count("(") == depth]
            if len(right) == 1:
                edits.append((m.start(2), m.end(2), right[0], f"{kind} {number}", title))
            else:
                notes.append(
                    (p.text, f"Cross-reference to check: {kind} {number} of this contract is not {title}.")
                )
        for start, end, number, cited, title in reversed(edits):
            at, placed = 0, False
            for node in nodes:
                size = len(node.text or "")
                if at <= start and end <= at + size:
                    node.text = node.text[: start - at] + number + node.text[end - at :]
                    changed = placed = True
                    break
                at += size
            if not placed:  # the number runs across two runs of text: the lawyer corrects it
                notes.append((p.text, f"Cross-reference to check: {cited} of this contract is not {title}."))
    if not changed:
        return docx_bytes, notes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), notes


_REF_CODE = re.compile(r"\bREF\s+(_Ref\w+)")
_TITLE_AFTER = re.compile(r"^\s*\(([A-Z][^()]{2,60})\)")


def _misdirected_fields(doc, schedules: dict[str, str]) -> bool:
    """A Word field shown as "Schedule 9" followed by its title, "(Service Levels)", whose
    bookmark opens another schedule's heading: refreshed, Word would print the other schedule's
    number beside this title, as when a precedent's bookmark sits on the next schedule. Where exactly one schedule has the title, the field becomes typed words with its number."""
    heading_of: dict[str, str] = {}
    for p in doc.paragraphs:
        for b in p._p.iter(qn("w:bookmarkStart")):
            heading_of.setdefault(b.get(qn("w:name")) or "", _words(p.text))
    changed = False
    for p in doc.paragraphs:
        nodes = list(p._p.iter(qn("w:fldChar"), qn("w:instrText"), qn("w:t")))
        k = 0
        while k < len(nodes):
            node = nodes[k]
            if node.tag != qn("w:fldChar") or node.get(qn("w:fldCharType")) != "begin":
                k += 1
                continue
            parts, code, shown, j = [node], "", [], k + 1
            state = "code"
            while j < len(nodes):
                n = nodes[j]
                if n.tag == qn("w:fldChar"):
                    parts.append(n)
                    kind = n.get(qn("w:fldCharType"))
                    if kind == "separate":
                        state = "result"
                    elif kind == "end":
                        break
                elif n.tag == qn("w:instrText"):
                    parts.append(n)
                    code += n.text or ""
                elif state == "result":
                    shown.append(n)
                j += 1
            after = "".join(n.text or "" for n in nodes[j + 1 : j + 4] if n.tag == qn("w:t"))
            ref, title = _REF_CODE.search(code), _TITLE_AFTER.match(after)
            words = "".join(n.text or "" for n in shown)
            if ref and title and re.fullmatch(r"Schedule \d+", words.strip()):
                target = heading_of.get(ref.group(1), "")
                named = _words(title.group(1))
                bare = re.sub(r"^schedule\s+\d+\s*", "", target)
                right = [n for n, t in schedules.items() if t == named or t.startswith(named + " ")]
                if target and not bare.startswith(named) and len(right) == 1 and shown:
                    shown[0].text = f"Schedule {right[0]}"
                    for n in shown[1:]:
                        n.text = ""
                    for x in parts:
                        x.getparent().remove(x)
                    changed = True
            k = j + 1
    return changed

