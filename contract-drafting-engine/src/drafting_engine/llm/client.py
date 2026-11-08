"""The chat model every agent uses, configured from the environment.

Two providers. Set ``DRAFTING_AZURE_OPENAI_ENDPOINT`` for Azure OpenAI (with
``DRAFTING_AZURE_OPENAI_DEPLOYMENT`` and ``DRAFTING_AZURE_OPENAI_API_VERSION``; leave
``DRAFTING_AZURE_OPENAI_API_KEY`` blank to sign in with ``DefaultAzureCredential``). Otherwise
``OPENAI_API_KEY`` and ``DRAFTING_MODEL`` select a model on the OpenAI API.
"""

from __future__ import annotations

import os
import threading
import time
from collections.abc import Callable
from contextlib import nullcontext

import httpx
from dotenv import find_dotenv, load_dotenv
from langchain_core.language_models.chat_models import BaseChatModel

# usecwd=True: a command line tool's .env belongs to the directory it is run from, not to
# wherever the package is installed.
load_dotenv(find_dotenv(usecwd=True))

_COGNITIVE_SERVICES_SCOPE = "https://cognitiveservices.azure.com/.default"

_REQUEST_TIMEOUT_SECONDS = float(os.environ.get("DRAFTING_LLM_TIMEOUT_SECONDS", "120"))
_MAX_RETRIES = int(os.environ.get("DRAFTING_LLM_MAX_RETRIES", "30"))

# Every draft in this process shares two limits on what it sends the model. A draft fans out
# to many agents at once, and a deployment's tokens-per-minute allowance is soon spent: past
# it the service refuses calls, and a draft that spends its retries being refused fails. A call
# waiting here spends no retry, so drafts sharing an allowance slow down instead of failing.
#
# DRAFTING_LLM_TOKENS_PER_MINUTE is the drafts' share, counted from each request's size;
# DRAFTING_LLM_CONCURRENCY is how many calls may be in flight. Unset or 0 is no limit.
_TOKENS_PER_MINUTE = int(os.environ.get("DRAFTING_LLM_TOKENS_PER_MINUTE", "0") or "0")
_CONCURRENCY = int(os.environ.get("DRAFTING_LLM_CONCURRENCY", "0") or "0")

# Bytes of request body per prompt token, roughly. JSON escaping only makes the estimate
# high, which errs towards leaving the allowance room.
_BYTES_PER_TOKEN = 4


class _TokenBudget:
    """Admits requests at ``per_minute`` tokens. Services enforce their allowance over
    windows of seconds rather than the whole minute, so a burst is held to ten seconds' worth."""

    def __init__(
        self,
        per_minute: int,
        clock: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._rate = per_minute / 60
        self._capacity = per_minute / 6
        self._level = self._capacity
        self._clock = clock
        self._sleep = sleep
        self._stamp = clock()
        self._lock = threading.Lock()

    def take(self, tokens: float) -> None:
        tokens = min(tokens, self._capacity)
        with self._lock:
            now = self._clock()
            self._level = min(self._capacity, self._level + (now - self._stamp) * self._rate)
            self._stamp = now
            if self._level < tokens:
                self._sleep((tokens - self._level) / self._rate)
                self._level = tokens
                self._stamp = self._clock()
            self._level -= tokens


class _GatedTransport(httpx.BaseTransport):
    """Holds a place in ``gate`` for the whole of each request, response body included,
    and spends the request's size from ``budget`` before sending it."""

    def __init__(
        self,
        gate: threading.Semaphore | None,
        budget: _TokenBudget | None,
        inner: httpx.BaseTransport,
    ) -> None:
        self._gate = gate
        self._budget = budget
        self._inner = inner

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        with self._gate or nullcontext():
            if self._budget is not None:
                self._budget.take(len(request.read()) / _BYTES_PER_TOKEN)
            response = self._inner.handle_request(request)
            try:
                response.read()
            except BaseException:
                response.close()
                raise
            return response

    def close(self) -> None:
        self._inner.close()


_gated_client: httpx.Client | None = None
_gated_client_lock = threading.Lock()


def _shared_http_client() -> httpx.Client | None:
    """One client, and so one set of limits, for every chat model this process builds."""
    global _gated_client
    if _CONCURRENCY < 1 and _TOKENS_PER_MINUTE < 1:
        return None
    with _gated_client_lock:
        if _gated_client is None:
            gate = threading.BoundedSemaphore(_CONCURRENCY) if _CONCURRENCY > 0 else None
            budget = _TokenBudget(_TOKENS_PER_MINUTE) if _TOKENS_PER_MINUTE > 0 else None
            _gated_client = httpx.Client(
                transport=_GatedTransport(gate, budget, httpx.HTTPTransport()),
                timeout=_REQUEST_TIMEOUT_SECONDS,
            )
        return _gated_client


# Reasoning models reject function tools on /chat/completions unless reasoning is off, and every
# agent here calls tools. Empty omits the parameter, for a model that does not accept it at all.
_REASONING_EFFORT = os.environ.get("DRAFTING_REASONING_EFFORT", "none").strip() or None


def _azure_auth_kwargs() -> dict:
    api_key = os.environ.get("DRAFTING_AZURE_OPENAI_API_KEY", "").strip()
    if api_key:
        return {"api_key": api_key}

    from azure.identity import DefaultAzureCredential, get_bearer_token_provider

    return {
        "azure_ad_token_provider": get_bearer_token_provider(
            DefaultAzureCredential(), _COGNITIVE_SERVICES_SCOPE
        )
    }


def build_chat_model() -> BaseChatModel:
    """Chat client with an explicit request timeout and retries.

    Both are load-bearing: a draft makes hundreds of calls, and with no socket timeout one
    half-open connection blocks a run forever instead of failing and being retried.

    No temperature: current reasoning models accept only their default, so a setting here
    could only ever break a run. Drafting is therefore not deterministic, and one run proves
    little; judge a playbook change on two.
    """
    common: dict = {
        "timeout": _REQUEST_TIMEOUT_SECONDS,
        "max_retries": _MAX_RETRIES,
        "reasoning_effort": _REASONING_EFFORT,
        "http_client": _shared_http_client(),
    }
    if os.environ.get("DRAFTING_AZURE_OPENAI_ENDPOINT"):
        from langchain_openai import AzureChatOpenAI

        return AzureChatOpenAI(
            azure_endpoint=os.environ["DRAFTING_AZURE_OPENAI_ENDPOINT"],
            azure_deployment=os.environ["DRAFTING_AZURE_OPENAI_DEPLOYMENT"],
            api_version=os.environ["DRAFTING_AZURE_OPENAI_API_VERSION"],
            **common,
            **_azure_auth_kwargs(),
        )
    from langchain_openai import ChatOpenAI

    return ChatOpenAI(model=os.environ["DRAFTING_MODEL"], **common)
