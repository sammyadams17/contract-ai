"""Draft a contract from a term sheet's record and a base document: plan, draft, check, write.

The command line and any hosting service both call this. It returns the redline.
"""

from __future__ import annotations

import datetime
import json
import re
import time
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

from drafting_engine.draft.agent import (
    ClauseAgent,
    draft_clause,
    fix_clause,
    footnotes_of,
)
from drafting_engine.draft.contract_checks import renamed_terms, whole_contract_faults
from drafting_engine.draft.planner import Unit, plan_units, refresh_plan
from drafting_engine.draft.references import typed_references
from drafting_engine.draft.writer import (
    accepted_view,
    annotate,
    ask_to_update_fields,
    carry_header_changes,
    clause_numbers,
    inserts_not_placed,
    no_inserted_titles,
    numbering_as_precedent,
    order_paragraph_marks,
    piece_notes,
    place_deferred_cells,
    redline,
    refresh_contents,
    refresh_reference_fields,
    remove_highlighting,
    repair_fields,
    restore_insert_order,
    restyle_paragraphs,
    rewrite_paragraphs,
    snapshot,
    stamp_properties,
    strike_dropped_footnotes,
    strike_paragraphs,
    strike_rows,
    unique_bookmarks,
    unique_paragraph_ids,
    unpair_inserts,
    without_comments,
    without_content_controls,
    write,
)
from drafting_engine.llm.agent import AgentDidNotCompleteError, patient
from drafting_engine.termsheet.record import Record
from drafting_engine.termsheet.record import plain as plain_text
from drafting_engine.word.adeu_engine import render_markdown


def _longest_first(units: list[tuple[str, str]]) -> list[int]:
    """The order to start drafting the pieces in: the longest form text first, ties in document
    order. A draft ends when its last piece does; started in document order, a long schedule
    near the end waited for every piece before it.
    Nothing a piece is drafted from depends on another piece, so only the timing changes."""
    return sorted(range(len(units)), key=lambda i: -len(units[i][1]))


class DraftCancelled(Exception):
    """The caller asked the draft to stop, at a point between pieces."""


@dataclass
class Drafted:
    """The redline, and the per-piece results."""

    docx: bytes
    results: list[dict] = field(default_factory=list)
    failed: list[dict] = field(default_factory=list)


def plan_form(
    base_clean: str,
    record: Record,
    playbook: Path,
    llm,
    plan_path: Path | None = None,
    say: Callable[[str], None] = lambda _note: None,
    precedent: bool = False,
    workers: int = 8,
) -> list[Unit]:
    """The form's pieces: the plan at ``plan_path`` when it already holds one, else a new
    plan, written there once made."""
    if plan_path is not None and plan_path.exists():
        planned = refresh_plan([Unit(**u) for u in json.loads(plan_path.read_text())], playbook)
        # written back, so the run folder shows the rules its drafters were actually given
        plan_path.write_text(json.dumps([u.__dict__ for u in planned], indent=1))
        say(f"planning the form — {len(planned)} pieces, from the saved plan")
        return planned
    say("planning the form")
    planned = plan_units(
        base_clean,
        playbook,
        llm,
        termsheet=[(s.index, s.heading.strip(), plain_text(s.text).strip()) for s in record.sections],
        workers=workers,
        precedent=precedent,
    )
    say(f"planning the form — {len(planned)} pieces")
    if plan_path is not None:
        plan_path.parent.mkdir(parents=True, exist_ok=True)
        plan_path.write_text(json.dumps([u.__dict__ for u in planned], indent=1))
    return planned


def draft_contract(
    record: Record,
    base_bytes: bytes,
    playbook: Path,
    llm,
    on_event: Callable[[str], None] | None = None,
    cancelled: Callable[[], bool] | None = None,
    *,
    precedent: bool = False,
    template: bytes | None = None,
    client_details: dict[str, str] | None = None,
    author: str = "",
    workers: int = 12,
    plan_workers: int = 8,
    only: set[str] | None = None,
    plan_path: Path | None = None,
) -> Drafted:
    """Plan the form, draft each piece, check the whole contract, and return the tracked-change
    document.

    ``playbook`` is the client's playbook file for this kind of contract. ``precedent`` says
    the base document is a signed contract from an earlier deal rather than the client's
    template; ``template``, on such a run, lets a drafter take the template's definition of a
    term the precedent does not define. ``client_details`` are the client's own particulars
    (registered office, notice details), which no term sheet states. ``author`` is written
    into the file's properties.

    ``plan_path``, when it already holds a plan, is used instead of planning again, which is
    how a rerun skips that stage. A new plan is written there once made. ``only`` drafts the
    named pieces alone, for trying a playbook change on a few.
    """
    say = on_event or (lambda _note: None)
    stop = cancelled or (lambda: False)
    details = {**(client_details or {}), "Drafting date": datetime.date.today().strftime("%-d %B %Y")}

    def halt(where: str) -> None:
        if stop():
            raise DraftCancelled(where)

    form_bytes = without_comments(without_content_controls(base_bytes))
    template_text = (
        render_markdown(without_content_controls(template), clean_view=True) if template else ""
    )
    base_clean = render_markdown(form_bytes, clean_view=True)
    base_anchored = render_markdown(form_bytes, clean_view=False)
    footnotes = footnotes_of(base_clean)
    form_numbers = clause_numbers(form_bytes, base_clean)

    halt("before planning")
    planned = plan_form(base_clean, record, playbook, llm, plan_path, say, precedent, plan_workers)
    unit_of = {u.title: u for u in planned}

    all_units = [(u.title, u.text) for u in planned]
    before_text: dict[str, str] = {}
    seen: list[str] = []
    for title, text in all_units:
        before_text[title] = "\n".join(seen)
        seen.append(text)
    units = select_units(all_units, only) if only else all_units

    agents: dict[str, ClauseAgent] = {}

    def result_of(title: str, agent, note: str) -> dict:
        work = agent.work
        return {
            "title": title,
            "edits": work.edits,
            "paragraphs": snapshot(work, title, before_text=before_text[title]),
            "open_points": agent.open_points,
            "refusals": [],
            "note": note,
            "rules_accounted": getattr(agent, "rules_accounted", []),
            "fix_accounted": getattr(agent, "fix_accounted", []),
            "rule_check": getattr(agent, "rule_check", []),
            "trace": agent.trace,
            "check": agent.check(),
            "state": getattr(agent, "state", dict)(),
        }

    @patient
    def one(unit):
        halt(unit[0])
        title, text = unit
        t0 = time.monotonic()
        say(f"drafting {title}")
        try:
            agent, note = draft_clause(
                heading=title,
                clause_text=_strip_anchors(text),
                anchored_text=base_anchored,
                contract_text=base_clean,
                contract_units=all_units,
                record=record,
                footnotes=footnotes,
                rules=unit_of[title].rules,
                termsheet_sections=unit_of[title].termsheet_sections,
                house_facts=details,
                llm=llm,
                precedent=precedent,
                numbers=form_numbers,
                template_text=template_text,
                term_homes=unit_of[title].term_homes,
                form_bytes=form_bytes,
                before_text=before_text[title],
            )
            agents[title] = agent
            result = result_of(title, agent, note)
        except DraftCancelled:
            raise
        except Exception as error:
            result = {
                "title": title,
                "edits": [],
                "open_points": [],
                "refusals": [],
                "note": f"worker failed: {type(error).__name__}: {str(error)[:200]}",
            }
        result["seconds"] = round(time.monotonic() - t0)
        say(f"drafting {title} — {len(result['edits'])} edits")
        return result

    order = _longest_first(units)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        by_place = dict(zip(order, pool.map(one, [units[i] for i in order]), strict=True))
    results = [by_place[i] for i in range(len(units))]
    # Every later step reads the pieces' agents in document order, as when they started so.
    in_order = {title: agents[title] for title, _ in all_units if title in agents}
    agents.clear()
    agents.update(in_order)

    def assemble():
        deferred_rows: list[list[str]] = []
        deferred_paragraphs: list[str] = []
        deferred_cells: list[tuple[str, str, str]] = []
        deferred_styles: list[str] = []
        plain, missing = write(
            form_bytes,
            [r.get("paragraphs", {}) for r in results],
            deferred_rows,
            deferred_paragraphs,
            deferred_cells,
            deferred_styles,
        )
        return plain, missing, deferred_rows, deferred_paragraphs, deferred_cells, deferred_styles

    plain, missing, deferred_rows, deferred_paragraphs, deferred_cells, deferred_styles = (
        assemble()
    )
    index = {title: i for i, (title, _) in enumerate(units)}

    # what only the whole contract shows, once every piece has drafted: a definition whose users
    # other pieces struck, a term several pieces each defined, and a piece whose own account
    # still says its work is outstanding, each handed back once to the piece that holds it
    accepted = accepted_view(plain, deferred_rows, deferred_paragraphs, deferred_cells)
    held = {
        r["title"]: [(p["text"], bool(p.get("inserted"))) for p in r["state"]["paras"]]
        for r in results
        if r.get("state")
    }
    renamed = renamed_terms(
        [
            (p.get("original") or "", p["text"])
            for r in results
            if r.get("state")
            for p in r["state"]["paras"]
        ]
    )
    last = whole_contract_faults(
        base_clean, render_markdown(accepted, clean_view=True), held, renamed
    )
    for r in results:
        left = [a for a in r.get("rules_accounted") or [] if a.get("status") == "outstanding"]
        if left:
            last.setdefault(r["title"], []).append(
                "Your account of this piece still says this work is outstanding: "
                + "; ".join(f"{a.get('rule', '')[:160]} ({a.get('where', '')[:120]})" for a in left)
                + ". Do it now. Where it belongs to another piece, change nothing for it and"
                " account it not_applicable."
            )
    for title, agent in agents.items():
        bare = agent.check().get("struck_without_reason", [])
        if bare:
            last.setdefault(title, []).append(
                "These form paragraphs were struck alone, with no reason and none of the term"
                " sheet's words: "
                + "; ".join(f"{b['id']} \"{b['text'][:100]}\"" for b in bare)
                + ". The form's words stand unless the term sheet or a rule changes them: restore"
                " each, or, where the term sheet or a rule does strike it, give that reason and the"
                " term sheet's words with note."
            )
    titles = [t for t in last if t in agents]
    if titles:
        halt("before the whole-contract check")
        say(
            f"the whole contract: putting right {sum(len(last[t]) for t in titles)} points"
            f" in {len(titles)} pieces"
        )

        @patient
        def settle(title: str) -> str:
            try:
                return fix_clause(agents[title], last[title], llm)
            except AgentDidNotCompleteError:
                return "the whole-contract fix did not finish"

        with ThreadPoolExecutor(max_workers=workers) as pool:
            settled = list(pool.map(settle, titles))
        for title, note in zip(titles, settled, strict=True):
            agent = agents[title]
            if getattr(agent, "fix_accounted", None):
                agent.rules_accounted = agent.fix_accounted
            kept = results[index[title]]
            results[index[title]] = {
                **result_of(title, agent, kept.get("note", "")),
                "fix_note": kept.get("fix_note", ""),
                "contract_check_note": note,
                "seconds": kept.get("seconds", 0),
            }
        plain, missing, deferred_rows, deferred_paragraphs, deferred_cells, deferred_styles = (
            assemble()
        )

    failed = [
        {"title": "writer", "error": f"not found: {m}", "old_text": m, "kind": "write"}
        for m in missing
    ]
    # a piece still accounting work as outstanding after it was sent back to finish is not
    # finished, and the run says so
    failed += [
        {
            "title": r["title"],
            "error": f"outstanding: {a.get('rule', '')[:120]} ({a.get('where', '')[:120]})",
            "old_text": "",
            "kind": "unfinished",
        }
        for r in results
        for a in r.get("rules_accounted") or []
        if a.get("status") == "outstanding"
    ]
    say("writing the document")
    # typed references read against the contract as it now numbers, before the redline marks
    # each correction as a change
    plain, reference_notes = typed_references(
        plain, accepted_view(plain, deferred_rows, deferred_paragraphs, deferred_cells)
    )
    docx = unpair_inserts(restore_insert_order(redline(form_bytes, plain), plain), form_bytes, plain)
    deferred_cells = place_deferred_cells(docx, plain, deferred_cells)
    docx = carry_header_changes(docx, form_bytes, plain)
    docx = strike_rows(docx, deferred_rows)
    docx = strike_paragraphs(docx, deferred_paragraphs)
    docx = rewrite_paragraphs(docx, deferred_cells)
    docx = repair_fields(docx, form_bytes, plain)
    docx = no_inserted_titles(restyle_paragraphs(docx, plain, deferred_styles))
    docx = strike_dropped_footnotes(docx, plain)
    docx = unique_bookmarks(unique_paragraph_ids(order_paragraph_marks(docx)))
    # a part restarts its numbering where the base document's does, whatever style a drafter chose
    docx = numbering_as_precedent(docx, form_bytes)
    docx = ask_to_update_fields(remove_highlighting(refresh_contents(refresh_reference_fields(docx))))
    # each comment with its piece: one whose paragraph is gone goes on its own piece's nearest
    notes: list = [(a, c, i) for i, r in enumerate(results) for a, c in piece_notes(r)]
    notes += reference_notes
    unplaced: list[tuple[str, str]] = []
    try:
        docx = annotate(docx, notes, unplaced=unplaced)
    except Exception as error:
        unplaced = [(a, c) for a, c, _ in notes]
        say(f"the comments could not be written: {type(error).__name__}")
    failed += [
        {"title": "comment", "error": f"not placed: {c[:200]}", "old_text": a[:100], "kind": "comment"}
        for a, c in unplaced
    ]
    failed += [
        {"title": "writer", "error": f"insert not placed: {t[:200]}", "old_text": t[:100], "kind": "write"}
        for t in inserts_not_placed(docx, [r.get("paragraphs", {}) for r in results])
    ]
    if author:
        docx = stamp_properties(docx, author, datetime.datetime.now(datetime.UTC))
    return Drafted(
        docx=docx,
        results=results,
        failed=failed,
    )


def _strip_anchors(text: str) -> str:
    return re.sub(r"\{#[^}]*\}", "", text)


def select_units(all_units: list, only: set[str]) -> list:
    """The planned units named in ``only``.

    A name is a title, or the start of exactly one title: the planner names the same piece
    "Description of Services" on one run and "Description of Services (Schedule 3)" on the
    next. A name with a star is a pattern for every title it fits ("Definitions:*",
    "*Insurance*"). A name matching no title, or several, is an error.
    """
    from fnmatch import fnmatchcase

    titles = [u[0] for u in all_units]
    chosen: set[str] = set()
    for name in only:
        if "*" in name:  # every piece whose title fits the pattern
            hits = [t for t in titles if fnmatchcase(t, name)]
            if not hits:
                raise ValueError(f"{name!r} names 0 pieces; the plan has {sorted(titles)}")
            chosen.update(hits)
            continue
        hits = [name] if name in titles else [t for t in titles if t.startswith(name)]
        if len(hits) != 1:
            raise ValueError(
                f"{name!r} names {len(hits)} pieces {hits}; the plan has {sorted(titles)}"
            )
        chosen.add(hits[0])
    return [u for u in all_units if u[0] in chosen]
