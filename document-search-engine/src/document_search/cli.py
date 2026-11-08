"""The command line.

    document-search ingest LIBRARY --db library.db     read a folder tree once
    document-search ask "question" --db library.db     one cited answer, its steps as it goes
    document-search eval QUESTIONS.json --db library.db --out results.json
                                                        every question against its known answer
"""

from __future__ import annotations

import argparse
import json
import pathlib
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor

from document_search.agent import answer
from document_search.library import Library


def _say(note: str) -> None:
    print(f"  {time.strftime('%H:%M:%S')}  {note}", flush=True)


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9£$€%.]+", " ", text.casefold().replace(",", "")).strip()


def score(expected: dict, got) -> dict:
    """Right where the answer states every expected value (in any of its listed wordings) and cites
    the expected document page.
    A count of things that can be looked at, never a judge's opinion: each miss says what is missing."""
    text = _norm(got.text)
    def says(v: str) -> bool:
        return re.search(rf"(?<![a-z0-9]){re.escape(_norm(v))}(?![a-z0-9])", text) is not None

    # each expected value is one wording or a list of the wordings that say it
    values = [v for v in expected.get("values", []) if not any(says(w) for w in ([v] if isinstance(v, str) else v))]
    wanted = expected.get("cites", [])
    cited = {(c["path"], c["page"]) for c in got.citations}
    pages = [w for w in wanted if (w["path"], w["page"]) not in cited]
    absent = [v for v in expected.get("not", []) if says(v)]
    return {
        "right": not values and not pages and not absent,
        "missing_values": values,
        "missing_citations": pages,
        "should_not_say": absent,
    }


def cmd_ingest(args) -> None:
    counts = Library(pathlib.Path(args.db)).ingest(pathlib.Path(args.library), say=_say)
    print(json.dumps(counts))


def cmd_ask(args) -> None:
    got = answer(args.question, Library(pathlib.Path(args.db)), say=_say)
    print("\n" + got.text)


def cmd_eval(args) -> None:
    library = Library(pathlib.Path(args.db))
    questions = json.loads(pathlib.Path(args.questions).read_text())

    def one(q: dict) -> dict:
        t0 = time.monotonic()
        got = answer(q["question"], library)
        return {"id": q["id"], "question": q["question"], "answer": got.text, "steps": got.steps,
                "citations": got.citations, "seconds": round(time.monotonic() - t0)} | score(q, got)

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        results = list(pool.map(one, questions))
    right = sum(r["right"] for r in results)
    for r in results:
        mark = "RIGHT" if r["right"] else "WRONG"
        print(f"{mark}  {r['id']}: {r['question']}")
        for k in ("missing_values", "missing_citations", "should_not_say"):
            if r[k]:
                print(f"        {k}: {r[k]}")
    print(f"\n{right} of {len(results)} right")
    if args.out:
        pathlib.Path(args.out).write_text(json.dumps(results, indent=1, ensure_ascii=False))


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(prog="document-search", description=__doc__.split("\n\n")[0])
    sub = ap.add_subparsers(dest="command", required=True)
    i = sub.add_parser("ingest", help="read a folder tree into the library")
    i.add_argument("library")
    i.add_argument("--db", required=True)
    i.set_defaults(func=cmd_ingest)
    a = sub.add_parser("ask", help="one question, answered with citations")
    a.add_argument("question")
    a.add_argument("--db", required=True)
    a.set_defaults(func=cmd_ask)
    e = sub.add_parser("eval", help="every question in a file against its known answer")
    e.add_argument("questions")
    e.add_argument("--db", required=True)
    e.add_argument("--out", default="")
    e.add_argument("--workers", type=int, default=3)
    e.set_defaults(func=cmd_eval)
    args = ap.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())
