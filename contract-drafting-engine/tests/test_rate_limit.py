"""Drafts sharing a model allowance wait instead of being refused."""

from __future__ import annotations

from drafting_engine.llm.client import _TokenBudget


def test_a_request_past_the_allowance_waits_for_the_budget_to_refill():
    now = [0.0]
    slept: list[float] = []

    def sleep(seconds: float) -> None:
        slept.append(seconds)
        now[0] += seconds

    budget = _TokenBudget(60_000, clock=lambda: now[0], sleep=sleep)  # 1,000 tokens a second
    budget.take(10_000)  # the ten-second burst is spent
    assert slept == []
    budget.take(2_000)
    assert slept == [2.0]
