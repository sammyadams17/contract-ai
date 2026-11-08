"""A stand-in chat model that plays each agent's part from a script, so the whole pipeline runs
with no model at all. It recognises the agent by its brief and answers with tool calls."""

from __future__ import annotations

import json
import re
from itertools import count

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage

PURPOSE = "evaluating a possible joint development of warehouse picking software"


class ScriptedLLM:
    def __init__(self) -> None:
        self.calls: list[str] = []  # which agent answered, in order
        self._ids = count()

    def bind_tools(self, tools):
        return _Bound(self, tools)


class _Bound:
    def __init__(self, owner: ScriptedLLM, tools) -> None:
        self.owner = owner
        self.tools = {t.name: t for t in tools}

    def _call(self, name: str, **args) -> AIMessage:
        return AIMessage(
            content="", tool_calls=[{"name": name, "args": args, "id": f"c{next(self.owner._ids)}"}]
        )

    def invoke(self, messages) -> AIMessage:
        system = next(m.content for m in messages if isinstance(m, SystemMessage))
        task = next(m.content for m in messages if isinstance(m, HumanMessage))
        results = [json.loads(m.content) for m in messages if isinstance(m, ToolMessage)]
        if system.startswith("# The planner's brief"):
            self.owner.calls.append("planner")
            n = int(re.search(r"THE FORM has (\d+) lines", task).group(1))
            return self._call(
                "submit_outcome", pieces=[{"title": "Agreement", "start": 0, "end": n - 1}]
            )
        if system.startswith("# The instructions brief"):
            self.owner.calls.append("instructions")
            sections = [int(s) for s in re.findall(r"^\s*\[(\d+)\]", task, re.M)]
            return self._call("submit_outcome", termsheet_sections=sections)
        if task.startswith("A DRAFTER WAS GIVEN THIS TASK"):
            self.owner.calls.append("rule check")
            return self._call("submit_outcome", verdicts=[])
        if "found these faults" in task:
            self.owner.calls.append("fix")
            return self._call("submit_outcome", note="Nothing to change.", account="none")
        if system.startswith("# The drafter's brief"):
            return self._draft(task, results)
        raise AssertionError(f"no script for an agent briefed {system[:60]!r}")

    def _draft(self, task: str, results: list) -> AIMessage:
        step = len(results)
        if step == 0:
            self.owner.calls.append("drafter")
            return self._call("outline")
        paras = results[0] if isinstance(results[0], list) else results[0].get("paragraphs", [])
        by_text = {p["text"]: p["id"] for p in paras}
        purpose = next(i for t, i in by_text.items() if t.startswith('"Purpose" means'))
        dated = next(i for t, i in by_text.items() if t.startswith("THIS AGREEMENT is dated"))
        if step == 1:
            return self._call(
                "replace",
                id=purpose,
                new_text=f'"Purpose" means {PURPOSE}; and',
                reason="",
                termsheet_quote="Evaluating a possible joint development of warehouse picking software.",
            )
        if step == 2:
            return self._call(
                "open_point",
                summary="The heads of terms give no date for the agreement; please confirm it.",
                at=dated,
            )
        rules = re.findall(r"^- (.{0,40})", task.split("RULES FOR THIS PIECE:")[1], re.M)
        return self._call(
            "submit_outcome",
            note="The Purpose is the heads of terms'. The date is open.",
            account="Purpose written; date left open.",
            rules_accounted=[{"rule": r, "where": "throughout", "status": "done"} for r in rules],
        )
