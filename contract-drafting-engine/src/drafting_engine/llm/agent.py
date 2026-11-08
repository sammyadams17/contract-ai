"""Generic tool-calling agent runner shared by every specialist agent.

Each agent gets a system prompt, its tools, and a `submit_outcome` tool whose
schema is that agent's output schema. The loop runs until the model calls
submit_outcome or max_iterations is reached, so an agent ends with structured
output rather than a free-written final answer.
"""

from __future__ import annotations

import json
import time

import openai
from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_core.tools import StructuredTool
from pydantic import BaseModel, ValidationError

from drafting_engine.llm.client import build_chat_model


class AgentDidNotCompleteError(Exception):
    pass


# A request the gateway cuts off or a server error is retried: the client's own retries do not
# cover a 499, and one such request can end a long run that was otherwise complete.
_RETRIES = 4
_RETRY_WAIT_SECONDS = 10.0
_RETRYABLE = {408, 409, 425, 429, 499, 500, 502, 503, 504}


def _invoke(llm_with_tools, messages):
    for attempt in range(_RETRIES + 1):
        try:
            return llm_with_tools.invoke(messages)
        except (openai.APIStatusError, openai.APIConnectionError) as error:
            status = getattr(error, "status_code", None)
            if attempt == _RETRIES or (status is not None and status not in _RETRYABLE):
                raise
            time.sleep(_RETRY_WAIT_SECONDS * (attempt + 1))
    raise AssertionError("unreachable")


def _build_submit_tool(final_output_schema: type[BaseModel]) -> StructuredTool:
    def _submit(**kwargs):
        return kwargs

    return StructuredTool.from_function(
        func=_submit,
        name="submit_outcome",
        description=(
            "Call this exactly once, as your final action, with your complete structured "
            "result. You must call this to finish - do not answer in plain text."
        ),
        args_schema=final_output_schema,
    )


def run_agent(
    system_prompt: str,
    task_message: str,
    tools: list[StructuredTool],
    final_output_schema: type[BaseModel],
    max_iterations: int = 20,
    llm=None,
) -> BaseModel:
    llm = llm or build_chat_model()
    submit_tool = _build_submit_tool(final_output_schema)
    all_tools = [*tools, submit_tool]
    tools_by_name = {t.name: t for t in all_tools}
    llm_with_tools = llm.bind_tools(all_tools)

    messages: list = [SystemMessage(content=system_prompt), HumanMessage(content=task_message)]

    for _ in range(max_iterations):
        response: AIMessage = _invoke(llm_with_tools, messages)
        messages.append(response)

        if not response.tool_calls:
            messages.append(
                HumanMessage(
                    content="You must act by calling a tool, and finish by calling submit_outcome. "
                    "Do not respond with plain text."
                )
            )
            continue

        for call in response.tool_calls:
            if call["name"] == "submit_outcome":
                try:
                    return final_output_schema.model_validate(call["args"])
                except ValidationError as e:
                    # Give the model a chance to correct its own output instead of crashing
                    # the whole agent run over a near-miss (e.g. status="complete").
                    messages.append(
                        ToolMessage(
                            content=json.dumps({"error": f"submit_outcome args were invalid: {e}"}),
                            tool_call_id=call["id"],
                        )
                    )
                    continue

            tool = tools_by_name.get(call["name"])
            if tool is None:
                result = {"error": f"unknown tool '{call['name']}'"}
            else:
                try:
                    result = tool.invoke(call["args"])
                except Exception as e:  # tool failures are observations, not crashes
                    result = {"error": str(e)}

            messages.append(
                ToolMessage(content=json.dumps(result, default=str), tool_call_id=call["id"])
            )

    raise AgentDidNotCompleteError(
        f"Agent did not call submit_outcome within {max_iterations} iterations"
    )


def patient(call):
    """Retry a call the service refuses for rate or drops mid-connection, waiting longer each
    time. When many agents run at once, the client's own few quick retries are not enough: a
    rate limit or a closed connection would end a run that only had to wait."""

    def wrapped(*args, **kwargs):
        for attempt in range(8):
            try:
                return call(*args, **kwargs)
            except Exception as error:
                text = str(error).lower()
                transient = (
                    "429" in str(error)
                    or "rate limit" in text
                    or "connection error" in text
                    or "connecterror" in text
                    or "ssl" in text
                    or "eof occurred" in text
                    or "timed out" in text
                )
                if not transient or attempt == 7:
                    raise
                time.sleep(15 * (attempt + 1))
        raise AssertionError("unreachable")

    return wrapped
