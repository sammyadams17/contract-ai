"""Thin wrapper around adeu (https://github.com/dealfluence/adeu): renders a Word file to the
Markdown the agents read, without ever leaving the original DOCX.

Why not a reduced text model parsed once: it loses real structure - clause numbers rendered
by Word list styles rather than literal text, and a table of contents that looks identical to
the clauses it lists, which led agents to "edit" a contents entry instead of the clause. The
redline itself is written by comparing documents (see draft/writer.py), not by adeu's edits.
"""

from __future__ import annotations

import io
import logging
import threading

import adeu
import structlog

# adeu logs every internal step (comment part creation, edit application, ...)
# at debug/info via structlog - useful for adeu's own development, but pure
# noise for this system's CLI and agent tool-call tracing.
structlog.configure(wrapper_class=structlog.make_filtering_bound_logger(logging.WARNING))

# Every tracked change this system writes carries ONE author, deliberately.
#
# adeu refuses to modify an active insertion belonging to a different author -
# correctly, since silently rewriting someone else's pending change is how a
# redline stops being reviewable. An author per agent or per stage would make the
# engine refuse its own earlier work: a later fix could not touch a passage an
# earlier stage changed. And the author field is client-facing: a lawyer opening
# the draft should see one author, not this system's internal wiring.
TRACKED_CHANGE_AUTHOR = "Drafting engine"


def load_docx_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def render_markdown(
    docx_bytes: bytes, clean_view: bool = False, include_appendix: bool = True
) -> str:
    """Render the document as Markdown + CriticMarkup for agent consumption.

    clean_view=True simulates "Accept All Changes" (current draftable state);
    clean_view=False shows track-changes markup (insertions/deletions/comments).
    """
    return adeu.extract_text_from_stream(
        io.BytesIO(_equations_as_text(docx_bytes)),
        filename="document.docx",
        clean_view=clean_view,
        include_appendix=include_appendix,
    )


MATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_W_DEL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del"


def math_text(el) -> str:
    """A Word equation's words on one line, as the drafter reads it and the writer matches it:
    its brackets and fractions written out, struck text left out. adeu reads no equation, so
    an equation paragraph used to reach the drafter as nothing at all, and a formula the term sheet
    changed (a cost-sharing threshold) could be neither replaced nor struck."""
    tag = el.tag
    if tag == _W_DEL or (isinstance(tag, str) and tag.endswith("Pr")):
        return ""
    if tag == MATH + "t":
        return el.text or ""
    parts = [math_text(c) for c in el]
    if not any(parts):
        return ""  # struck whole: its brackets go with its words
    if tag == MATH + "d":
        pr = el.find(MATH + "dPr")
        beg = pr.find(MATH + "begChr") if pr is not None else None
        end = pr.find(MATH + "endChr") if pr is not None else None
        return (
            (beg.get(MATH + "val", "") if beg is not None else "(")
            + ", ".join(p for p in parts if p)
            + (end.get(MATH + "val", "") if end is not None else ")")
        )
    if tag == MATH + "f":
        num, den = el.find(MATH + "num"), el.find(MATH + "den")
        return f"{math_text(num) if num is not None else ''}/{math_text(den) if den is not None else ''}"
    return "".join(parts)


def outer_equations(root) -> list:
    """Each equation of the element not inside another, the display paragraph's wrapper where
    there is one."""
    return [
        m
        for m in root.iter(MATH + "oMathPara", MATH + "oMath")
        if not any(True for _ in m.iterancestors(MATH + "oMathPara", MATH + "oMath"))
    ]


def _equations_as_text(docx_bytes: bytes) -> bytes:
    """The document with each equation written as a run of its text, for the render only."""
    import zipfile

    from docx import Document
    from docx.oxml.ns import qn

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if b"oMath" not in z.read("word/document.xml"):
            return docx_bytes
    doc = Document(io.BytesIO(docx_bytes))
    for m in outer_equations(doc.element.body):
        text = math_text(m)
        if text:
            r = m.makeelement(qn("w:r"), {})
            t = r.makeelement(qn("w:t"), {})
            t.text = text
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
            m.addprevious(r)
        m.getparent().remove(m)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# CT_ParaRPr's schema order: revision marks first, then the run properties.
_REVISION_MARK_TAGS = ("w:ins", "w:del", "w:moveFrom", "w:moveTo")


def _convert_deleted_field_instructions(doc) -> None:
    """Rename `w:instrText` to `w:delInstrText` inside every tracked deletion.

    OOXML pairs run content with a deleted form - `w:t`/`w:delText`,
    `w:instrText`/`w:delInstrText` - and Word refuses the mixed form with its
    repair dialog. adeu's `track_delete_run` converts only `w:t`, so deleting a
    clause that contains a cross-reference field leaves `w:instrText` inside
    `w:del`, and renaming exactly those elements makes the file open clean. Normalised here, beside the revision-mark
    reordering, so every document this wrapper produces is Word-readable
    whatever adeu wrote.
    """
    from docx.oxml.ns import qn

    for d in doc.element.body.iter(qn("w:del")):
        if d.getparent() is not None and d.getparent().tag == qn("w:rPr"):
            continue  # a deleted paragraph mark holds no field runs
        for instr in d.iter(qn("w:instrText")):
            instr.tag = qn("w:delInstrText")


# An editor asks "is this target findable?" for every candidate it widens to -
# thousands of questions about one document per revise on a long passage. Parsing the
# document answers none of them, so it is done once per document: `validate_edits` does
# not mutate the engine, and adeu builds its own lookup maps lazily on first use. The
# lock covers that lazy build; validation itself is milliseconds once the maps exist.
_VALIDATORS_KEPT = 4
_validators: dict[bytes, adeu.RedlineEngine] = {}
_validators_lock = threading.Lock()


def _validate(docx_bytes: bytes, edits: list) -> list[str]:
    with _validators_lock:
        engine = _validators.get(docx_bytes)
        if engine is None:
            engine = adeu.RedlineEngine(io.BytesIO(docx_bytes), author="validator")
            if len(_validators) >= _VALIDATORS_KEPT:
                del _validators[next(iter(_validators))]
            _validators[docx_bytes] = engine
        return engine.validate_edits(edits)


def compare_documents(original_bytes: bytes, revised_bytes: bytes) -> str:
    """Human-readable unified diff between two DOCX files' clean text - for
    the "final draft vs approved base" comparison.
    """
    original_text = render_markdown(original_bytes, clean_view=True, include_appendix=False)
    revised_text = render_markdown(revised_bytes, clean_view=True, include_appendix=False)
    return adeu.diff.create_unified_diff(original_text, revised_text)
