"""The model, configured from the environment, called patiently.

Azure OpenAI when ``AZURE_OPENAI_ENDPOINT`` is set (with ``AZURE_OPENAI_DEPLOYMENT``,
``AZURE_OPENAI_API_VERSION`` and ``AZURE_OPENAI_API_KEY``, or Azure sign-in when the key is blank);
otherwise the OpenAI API with ``OPENAI_API_KEY`` and ``OPENAI_MODEL``. ``EMBEDDING_MODEL`` turns on
the meaning half of search; without it, search is by words alone.
"""

from __future__ import annotations

import os
import threading
import time
from collections.abc import Callable
from functools import lru_cache
from typing import Any

from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv(usecwd=True))

# Model calls in flight at once, across every answer and deep search in this process: a deep
# search fans out, and without one shared limit two readers asking at once double it.
_GATE = threading.BoundedSemaphore(int(os.environ.get("SEARCH_LLM_CONCURRENCY", "4") or "4"))
_TRIES = 6
_WAIT = 15.0
_LONGEST = 60.0


@lru_cache(maxsize=1)
def client() -> Any:
    if os.environ.get("AZURE_OPENAI_ENDPOINT"):
        from openai import AzureOpenAI

        key = os.environ.get("AZURE_OPENAI_API_KEY", "").strip()
        auth: dict[str, Any] = {"api_key": key}
        if not key:
            from azure.identity import DefaultAzureCredential, get_bearer_token_provider

            auth = {
                "azure_ad_token_provider": get_bearer_token_provider(
                    DefaultAzureCredential(), "https://cognitiveservices.azure.com/.default"
                )
            }
        return AzureOpenAI(
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
            api_version=os.environ["AZURE_OPENAI_API_VERSION"],
            timeout=120,
            max_retries=2,
            **auth,
        )
    from openai import OpenAI

    return OpenAI(timeout=120, max_retries=2)


def model() -> str:
    return os.environ.get("AZURE_OPENAI_DEPLOYMENT") or os.environ["OPENAI_MODEL"]


def embedding_model() -> str:
    return os.environ.get("EMBEDDING_MODEL", "")


def _asked_to_wait(error: BaseException) -> float | None:
    headers = getattr(getattr(error, "response", None), "headers", None) or {}
    for header, scale in (("retry-after-ms", 1000.0), ("retry-after", 1.0)):
        if headers.get(header) is not None:
            try:
                return float(headers[header]) / scale
            except (TypeError, ValueError):
                continue
    return None


def patiently[T](call: Callable[[], T], say: Callable[[str], None] = lambda _n: None) -> T:
    """``call()``, waited for and tried again after each refusal for rate, longer each time.
    A refusal turned straight into the answer reads to a lawyer as the tool being broken."""
    for attempt in range(_TRIES):
        try:
            with _GATE:
                return call()
        except Exception as error:
            if getattr(error, "status_code", None) not in (429, 500, 502, 503, 504) or attempt == _TRIES - 1:
                raise
            wait = min(_LONGEST, max(1.0, _asked_to_wait(error) or _WAIT * (attempt + 1)))
            say(f"waiting for the model's quota, trying again in {round(wait)}s")
            time.sleep(wait)
    raise AssertionError("unreachable")


def chat(
    messages: list[dict],
    tools: list[dict] | None = None,
    say: Callable[[str], None] = lambda _n: None,
    response_format: dict | None = None,
):
    """One chat completion; the message the model returned."""
    kwargs: dict[str, Any] = {"model": model(), "messages": messages}
    if tools:
        kwargs["tools"] = tools
    if response_format:
        kwargs["response_format"] = response_format
    effort = os.environ.get("REASONING_EFFORT", "none").strip()
    if effort:
        kwargs["reasoning_effort"] = effort
    return patiently(lambda: client().chat.completions.create(**kwargs), say).choices[0].message


def embed(texts: list[str]) -> list[list[float]]:
    """Vectors for ``texts``, in batches; empty when no embedding model is configured."""
    name = embedding_model()
    if not name or not texts:
        return []
    out: list[list[float]] = []
    for start in range(0, len(texts), 64):
        batch = [t[:8000] for t in texts[start : start + 64]]
        response = patiently(lambda b=batch: client().embeddings.create(model=name, input=b))
        out += [d.embedding for d in response.data]
    return out
