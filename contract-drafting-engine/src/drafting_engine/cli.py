"""The command line: read a term sheet, draft a contract, annotate the term sheet, compare versions.

    drafting-engine read     TERMSHEET  --out RUN               the record (record.json)
    drafting-engine draft    RUN --base FORM.docx --playbook P  the redline (draft.docx)
    drafting-engine annotate RUN --termsheet TERMSHEET          the annotated term sheet
    drafting-engine compare  BEFORE.docx AFTER.docx              every difference, as text

A run folder holds one deal's work: record.json, plan.json (reused by a rerun), draft.docx and
edits.json (every piece's edits, open points and the run's failures).
"""

from __future__ import annotations

import argparse
import dataclasses
import io
import json
import os
import pathlib
import sys
import time

from docx import Document

from drafting_engine.draft.annotate import (
    annotate,
    record_items,
    render_in_place,
    render_rebuilt,
    termsheet_items,
)
from drafting_engine.draft.compare import compare_versions
from drafting_engine.draft.run import draft_contract
from drafting_engine.draft.writer import normalise_hyphens, repair_part_names
from drafting_engine.llm.client import build_chat_model
from drafting_engine.termsheet.document import render_termsheet
from drafting_engine.termsheet.extract import build_record
from drafting_engine.termsheet.record import load_record, save_record


def _say(note: str) -> None:
    print(f"{time.strftime('%H:%M:%S')}  {note}", flush=True)


def _model_name() -> str:
    return os.environ.get("DRAFTING_AZURE_OPENAI_DEPLOYMENT") or os.environ.get("DRAFTING_MODEL", "")


def _vocabulary(playbook: pathlib.Path | None) -> str:
    """The playbook folder's vocabulary.md, where it has one."""
    if playbook is None:
        return ""
    path = playbook.parent / "vocabulary.md"
    return path.read_text() if path.exists() else ""


def cmd_read(args: argparse.Namespace) -> None:
    run = pathlib.Path(args.out)
    text = render_termsheet(args.termsheet)
    _say(f"read {args.termsheet}: {len(text):,} characters; reading each section")
    record = build_record(
        args.deal or run.name,
        text,
        build_chat_model(),
        _model_name(),
        workers=args.workers,
        vocabulary=_vocabulary(pathlib.Path(args.playbook) if args.playbook else None),
    )
    save_record(record, run / "record.json")
    unverified = sum(not f.quote_verified for f in record.facts)
    _say(
        f"record: {len(record.sections)} sections, {len(record.facts)} terms"
        f" ({unverified} quotes not found verbatim), {len(record.gaps)} gaps -> {run / 'record.json'}"
    )
    for note in record.notes:
        print(f"  note: {note}")


def cmd_draft(args: argparse.Namespace) -> None:
    run = pathlib.Path(args.run)
    record = load_record(run / "record.json")
    if args.deal_facts:
        # facts a person states once for the whole deal, which decide the form's alternatives
        facts = json.loads(pathlib.Path(args.deal_facts).read_text())
        record.structure = {k: {"value": str(v), "quote": ""} for k, v in facts.items()}
    playbook = pathlib.Path(args.playbook)
    base = normalise_hyphens(repair_part_names(pathlib.Path(args.base).read_bytes()))
    template = repair_part_names(pathlib.Path(args.template).read_bytes()) if args.template else None
    details = json.loads(pathlib.Path(args.client_details).read_text()) if args.client_details else {}
    drafted = draft_contract(
        record,
        base,
        playbook,
        build_chat_model(),
        on_event=_say,
        precedent=args.precedent,
        template=template,
        client_details=details,
        author=args.author,
        workers=args.workers,
        plan_workers=args.plan_workers,
        only=set(args.only) if args.only else None,
        plan_path=run / "plan.json",
    )
    (run / "draft.docx").write_bytes(drafted.docx)
    (run / "edits.json").write_text(
        json.dumps(
            {
                "results": [{k: v for k, v in r.items() if k != "trace"} for r in drafted.results],
                "failed": drafted.failed,
            },
            indent=1,
            ensure_ascii=False,
            default=str,
        )
    )
    points = sum(len(r.get("open_points", [])) for r in drafted.results)
    edits = sum(len(r.get("edits", [])) for r in drafted.results)
    _say(f"draft: {len(drafted.results)} pieces, {edits} edits, {points} open points -> {run / 'draft.docx'}")
    for f in drafted.failed:
        print(f"  FAILED {f['kind']}: {f['title']}: {f['error'][:160]}")


def cmd_annotate(args: argparse.Namespace) -> None:
    run = pathlib.Path(args.run)
    record = json.loads((run / "record.json").read_text())
    edits = json.loads((run / "edits.json").read_text())
    contract = (run / "draft.docx").read_bytes()
    source_path = pathlib.Path(args.termsheet)
    if source_path.suffix.lower() == ".docx":
        source = source_path.read_bytes()
        document = Document(io.BytesIO(source))
        items = termsheet_items(document)
    else:
        document = None
        items = record_items(record["sections"])
    notes = annotate(items, contract, edits, record["deal"], build_chat_model(), workers=args.workers)
    if document is not None:
        data = render_in_place(document, source, items, notes)
    else:
        data = render_rebuilt(f"Annotated term sheet – {record['deal']}", items, notes)
    (run / "annotated-termsheet.docx").write_bytes(data)
    _say(f"annotated: {len(items)} items, {len(notes)} noted -> {run / 'annotated-termsheet.docx'}")


def cmd_compare(args: argparse.Namespace) -> None:
    changes = compare_versions(
        pathlib.Path(args.before).read_bytes(), pathlib.Path(args.after).read_bytes()
    )
    for n, c in enumerate(changes, 1):
        print(json.dumps({"n": n, **dataclasses.asdict(c)}, ensure_ascii=False))


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(prog="drafting-engine", description=__doc__.split("\n\n")[0])
    sub = ap.add_subparsers(dest="command", required=True)

    read = sub.add_parser("read", help="read a term sheet into a record")
    read.add_argument("termsheet", help="the term sheet, .docx or .pdf")
    read.add_argument("--out", required=True, help="the run folder to write record.json into")
    read.add_argument("--deal", default="", help="the deal's name; default the run folder's")
    read.add_argument("--playbook", default="", help="a playbook file; its folder's vocabulary.md is read")
    read.add_argument("--workers", type=int, default=6)
    read.set_defaults(func=cmd_read)

    draft = sub.add_parser("draft", help="draft the contract from a run's record")
    draft.add_argument("run", help="the run folder holding record.json")
    draft.add_argument("--base", required=True, help="the base document: the template or a precedent")
    draft.add_argument("--playbook", required=True, help="the playbook file for this contract type")
    draft.add_argument("--precedent", action="store_true", help="the base is a precedent, not the template")
    draft.add_argument("--template", default="", help="on a precedent run, the client's template")
    draft.add_argument("--client-details", default="", help="a JSON file of the client's own particulars")
    draft.add_argument("--deal-facts", default="", help="a JSON file of facts for the whole deal")
    draft.add_argument("--author", default="", help="written into the file's properties")
    draft.add_argument("--only", nargs="*", help="draft only these pieces (titles, or patterns with *)")
    draft.add_argument("--workers", type=int, default=12)
    draft.add_argument("--plan-workers", type=int, default=8)
    draft.set_defaults(func=cmd_draft)

    note = sub.add_parser("annotate", help="the term sheet with a note naming where each item landed")
    note.add_argument("run", help="the run folder holding record.json, draft.docx and edits.json")
    note.add_argument("--termsheet", required=True, help="the term sheet the record was read from")
    note.add_argument("--workers", type=int, default=8)
    note.set_defaults(func=cmd_annotate)

    compare = sub.add_parser("compare", help="every difference between two versions of a contract")
    compare.add_argument("before")
    compare.add_argument("after")
    compare.set_defaults(func=cmd_compare)

    args = ap.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())
