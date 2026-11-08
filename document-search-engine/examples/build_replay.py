"""Build the click-through demo: one HTML file, no server, no model key, from a saved run.

    .venv/bin/python examples/build_replay.py      # examples/sample-run -> demo/index.html

It replays saved answers: each question's steps, the answer with its citations, and the cited
page shown when its link is clicked. The pages' text is inside the page.
"""

from __future__ import annotations

import datetime
import json
import pathlib
import sys

from document_search.library import Library

ROOT = pathlib.Path(__file__).resolve().parents[1]


def build(run: pathlib.Path, library_root: pathlib.Path, out: pathlib.Path) -> None:
    library = Library(run / "library.db")
    answers = json.loads((run / "answers.json").read_text())
    # each cited page's text, shown in the page itself: a host that runs pages in a locked-down
    # frame blocks the pop-up a PDF would open in
    from document_search.library import READABLE
    from document_search.read import pages_of

    files = {
        str(p.relative_to(library_root)): pages_of(p)
        for p in sorted(library_root.rglob("*"))
        if p.suffix.lower() in READABLE
    }
    data = {
        "answers": [{k: a[k] for k in ("id", "question", "answer", "steps", "seconds", "right")} for a in answers],
        "folders": [{"folder": f} for f, _ in library.folders()],
        "files": files,
        "built": datetime.date.today().isoformat(),
    }
    page = (ROOT / "examples" / "replay.html").read_text()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page.replace("/*DATA*/null", payload))
    print(out, f"{out.stat().st_size // 1024} KB")


if __name__ == "__main__":
    run = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "examples" / "sample-run"
    out = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "demo" / "index.html"
    build(run, ROOT / "examples" / "library", out)
