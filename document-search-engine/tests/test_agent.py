"""The loop, offline: a tool returns its evidence and the answer cites the page it gave."""

from __future__ import annotations

import json
from types import SimpleNamespace

import document_search.agent as agent_mod
from document_search.agent import answer


def _call(name, **args):
    return SimpleNamespace(id=f"c-{name}", function=SimpleNamespace(name=name, arguments=json.dumps(args)))


def _script(*replies):
    queue = list(replies)

    def chat(messages, tools=None, say=None, response_format=None):
        reply = queue.pop(0)
        return reply(messages) if callable(reply) else reply

    return chat


def _link_from(messages):
    """The link the search returned, as the model would copy it."""
    hits = json.loads(next(m["content"] for m in reversed(messages) if m["role"] == "tool"))
    return hits[0]["link"]


def test_an_answer_cites_the_page_its_search_returned(library, monkeypatch):
    monkeypatch.setattr(
        agent_mod.llm, "chat",
        _script(
            SimpleNamespace(content="", tool_calls=[_call("find_clauses", query="rent", folder="Harbour Street Lease")]),
            lambda m: SimpleNamespace(content=f"The rent is £56,000 a year from 1 April 2025 ([Rent Review Memorandum 2025, p1]({_link_from(m)})).", tool_calls=None),
        ),
    )
    got = answer("What is the rent now?", library)
    assert [c["page"] for c in got.citations] == [1]
    assert got.steps == ["find_clauses(query='rent', folder='Harbour Street Lease')"]
