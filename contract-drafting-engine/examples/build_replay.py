"""Build the click-through demo: one HTML file, no server, no model key, from a saved run.

    .venv/bin/python examples/build_replay.py            # the saved runs in RUNS -> demo/index.html
    .venv/bin/python examples/build_replay.py OUT.html

The page replays the run: the term sheet, the agents at work (the run's own log, sped up), the
redline with its comments, and the annotated term sheet, each note a link to its clause. The
Word files are inside the page, for download.
"""

from __future__ import annotations

import base64
import datetime
import html
import io
import json
import pathlib
import re
import sys
import zipfile

from docx import Document
from docx.oxml.ns import qn

from drafting_engine.word.numbering import numbers_by_paragraph

ROOT = pathlib.Path(__file__).resolve().parents[1]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _comments(docx: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        if "word/comments.xml" not in z.namelist():
            return {}
        from lxml import etree

        root = etree.fromstring(z.read("word/comments.xml"))
    return {
        c.get(qn("w:id")): " ".join(
            "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in c.iter(qn("w:p"))
        ).strip()
        for c in root.iter(qn("w:comment"))
    }


def _segments(p) -> list[tuple[str, str]]:
    """The paragraph's text in order, each run marked same, ins or del."""
    out: list[tuple[str, str]] = []
    for child in p.iterchildren():
        tag = child.tag
        if tag == qn("w:r"):
            out.append(("same", "".join(t.text or "" for t in child.iter(qn("w:t")))))
        elif tag == qn("w:ins"):
            out.append(("ins", "".join(t.text or "" for t in child.iter(qn("w:t")))))
        elif tag == qn("w:del"):
            out.append(("del", "".join(t.text or "" for t in child.iter(qn("w:delText")))))
    return [(k, t) for k, t in out if t]


def redline(docx: bytes) -> list[dict]:
    doc = Document(io.BytesIO(docx))
    numbers = numbers_by_paragraph(docx, doc)
    notes = _comments(docx)
    out = []
    for i, para in enumerate(doc.paragraphs):
        p = para._p
        segments = _segments(p)
        if not segments:
            continue
        rpr = p.find(f"{W}pPr/{W}rPr")
        mark = (
            "ins"
            if rpr is not None and rpr.find(qn("w:ins")) is not None
            else "del"
            if rpr is not None and rpr.find(qn("w:del")) is not None
            else ""
        )
        if mark == "ins" or all(k == "ins" for k, _ in segments):
            mark = "ins"
        if mark == "del" or all(k == "del" for k, _ in segments):
            mark = "del"
        style = para.style.name if para.style is not None else ""
        number = numbers.get(i)
        out.append(
            {
                "n": number.marker if number else "",
                "full": number.full if number else "",
                "kind": "title" if style == "Title" else "heading" if style.startswith("Heading") else "body",
                "level": len(number.path) if number else 0,
                "mark": mark,
                "segments": segments,
                "comments": [notes[c.get(qn("w:id"))] for c in p.iter(qn("w:commentRangeStart")) if c.get(qn("w:id")) in notes],
            }
        )
    return out


_NOTE = re.compile(r"\s*\[([^\[\]]+)\]\s*$")


def annotated(docx: bytes) -> list[dict]:
    """The annotated term sheet's paragraphs, each with its note taken off the end."""
    out = []
    for p in Document(io.BytesIO(docx)).paragraphs:
        text = p.text.strip()
        if not text:
            continue
        m = _NOTE.search(text)
        out.append(
            {
                "text": text[: m.start()] if m else text,
                "note": m.group(1) if m else "",
                "heading": bool(re.match(r"^\d+\.\s", text)) or p.style.name.startswith(("Heading", "Title")),
            }
        )
    return out


def timeline(log: str) -> list[dict]:
    rows = []
    for line in log.splitlines():
        m = re.match(r"(\d\d):(\d\d):(\d\d)  (.*)", line)
        if m:
            h, mi, s, text = m.groups()
            rows.append({"t": int(h) * 3600 + int(mi) * 60 + int(s), "text": text})
    if rows:
        start = rows[0]["t"]
        for r in rows:
            r["t"] -= start
    return rows


def one_run(run: pathlib.Path, title: str, label: str, base: str, source: str) -> dict:
    """One saved run, as the page shows it: label names the switch, base what was drafted from."""
    draft = (run / "draft.docx").read_bytes()
    notes = (run / "annotated-termsheet.docx").read_bytes()
    record = json.loads((run / "record.json").read_text())
    edits = json.loads((run / "edits.json").read_text())
    plan = json.loads((run / "plan.json").read_text())
    return {
        "label": label,
        "base": base,
        "from": source,
        "title": title,
        "redline": redline(draft),
        "termsheet": annotated(notes),
        "timeline": timeline((run / "run.log").read_text()),
        "stats": {
            "sections": len(record["sections"]),
            "terms": len(record["facts"]),
            "pieces": len(plan),
            "edits": sum(len(r.get("edits", [])) for r in edits["results"]),
            "open_points": sum(len(r.get("open_points", [])) for r in edits["results"]),
        },
        "files": {
            "draft": base64.b64encode(draft).decode(),
            "annotated": base64.b64encode(notes).decode(),
        },
    }


def build(runs: list[dict], out: pathlib.Path, title: str) -> None:
    data = {
        "order": [r["key"] for r in runs],
        "runs": {r["key"]: one_run(r["run"], r["title"], r["label"], r["base"], r["from"]) for r in runs},
        "built": datetime.date.today().isoformat(),
    }
    page = (ROOT / "examples" / "replay.html").read_text()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page.replace("/*DATA*/null", payload).replace("{{TITLE}}", html.escape(title)))
    print(out, f"{out.stat().st_size // 1024} KB")


DEAL = "Mutual NDA: Northwind Analytics and Bluefield Robotics"
RUNS = [
    {"key": "template", "run": ROOT / "examples" / "nda" / "sample-run", "title": DEAL,
     "label": "From the template", "base": "template",
     "from": "The engine drafts from the client's own template."},
    {"key": "precedent", "run": ROOT / "examples" / "nda" / "precedent-run",
     "title": DEAL + ", drafted from the client's last NDA, with Saltmarsh Data",
     "label": "From a previous deal", "base": "previous deal",
     "from": "The engine drafts from the client's last signed deal, replacing that deal's parties and terms with this one's."},
]


if __name__ == "__main__":
    out = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "demo" / "index.html"
    build([r for r in RUNS if (r["run"] / "draft.docx").exists()], out, DEAL)
