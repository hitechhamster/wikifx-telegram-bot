import json
import logging
import os
from collections.abc import Awaitable, Callable
from typing import Any

import httpx


LOGGER = logging.getLogger("wikifx-deepseek-agent")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
DEEPSEEK_API_BASE = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-flash").strip()
DEEPSEEK_THINKING = os.getenv("DEEPSEEK_THINKING", "disabled").strip().casefold()
DEEPSEEK_TIMEOUT_SECONDS = float(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "30"))
AI_MAX_TOKENS = int(os.getenv("AI_MAX_TOKENS", "700"))
AI_TOOL_ROUNDS = 4

ToolExecutor = Callable[[str, dict[str, Any]], Awaitable[dict[str, Any]]]


SYSTEM_PROMPT = """You are WikiFX Broker Check India, a concise Telegram assistant.
Reply in the same language as the user. You help users look up brokers and understand only the
facts returned by the provided tools.

Rules:
1. For every broker-specific question, always call search_brokers before answering.
2. Treat questions such as "Exness怎么样", "is XM safe?", and "tell me about Octa" as broker searches.
3. Never invent licences, regulation, scores, risk status, dates, websites, or safety conclusions.
4. If data is missing, say it is unavailable. If several brokers match, list the candidates and ask
   the user to specify one; do not merge their facts.
5. Never guarantee that a broker is safe, reliable, profitable, or scam-free. Explain that the
   information is a snapshot and not investment advice.
6. Use list_my_follows when the user asks what they follow.
7. Use set_broker_follow only when the user explicitly asks to follow or unfollow a broker. Asking
   about a broker is not permission to follow it.
8. Follow/unfollow works only in private chat. Do not claim that an action succeeded unless the tool
   result says it succeeded.
9. Prefer short, useful answers. Include the WikiFX profile URL returned by the tool when relevant.
"""


TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "search_brokers",
            "description": (
                "Search the local WikiFX broker workbook by broker name, company name, official "
                "domain, licence number, or WikiFX ID. Use this before any broker-specific answer."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "The broker-identifying text extracted from the user's message.",
                    }
                },
                "required": ["query"],
                "additionalProperties": False,
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_my_follows",
            "description": "List brokers currently followed by this Telegram user.",
            "parameters": {
                "type": "object",
                "properties": {},
                "additionalProperties": False,
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_broker_follow",
            "description": (
                "Follow or unfollow one exact broker for this Telegram user. Only call after the "
                "user explicitly requests the action and an exact WikiFX ID is known."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "broker_id": {
                        "type": "string",
                        "description": "Exact WikiFX broker ID returned by search_brokers.",
                    },
                    "action": {
                        "type": "string",
                        "enum": ["follow", "unfollow"],
                    },
                },
                "required": ["broker_id", "action"],
                "additionalProperties": False,
            },
        },
    },
]


class AIServiceError(RuntimeError):
    pass


def is_ai_enabled() -> bool:
    return bool(DEEPSEEK_API_KEY)


def _request_payload(messages: list[dict[str, Any]]) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": DEEPSEEK_MODEL,
        "messages": messages,
        "tools": TOOLS,
        "tool_choice": "auto",
        "max_tokens": AI_MAX_TOKENS,
    }
    if DEEPSEEK_THINKING in {"low", "enabled"}:
        payload["thinking"] = {"type": "enabled"}
        payload["reasoning_effort"] = "low"
    else:
        payload["thinking"] = {"type": "disabled"}
    return payload


async def ask_deepseek(
    user_text: str,
    history: list[dict[str, str]],
    tool_executor: ToolExecutor,
) -> str:
    if not DEEPSEEK_API_KEY:
        raise AIServiceError("DeepSeek is not configured.")

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": SYSTEM_PROMPT},
        *history[-8:],
        {"role": "user", "content": user_text[:2000]},
    ]
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient(timeout=DEEPSEEK_TIMEOUT_SECONDS) as client:
        for _ in range(AI_TOOL_ROUNDS):
            try:
                response = await client.post(
                    f"{DEEPSEEK_API_BASE}/chat/completions",
                    headers=headers,
                    json=_request_payload(messages),
                )
                response.raise_for_status()
                data = response.json()
            except httpx.HTTPStatusError as error:
                LOGGER.warning("DeepSeek HTTP error: %s", error.response.status_code)
                raise AIServiceError(
                    f"DeepSeek request failed with status {error.response.status_code}."
                ) from error
            except (httpx.HTTPError, ValueError) as error:
                LOGGER.warning("DeepSeek request failed: %s", type(error).__name__)
                raise AIServiceError("DeepSeek is temporarily unavailable.") from error

            choices = data.get("choices") or []
            if not choices:
                raise AIServiceError("DeepSeek returned no answer.")

            message = choices[0].get("message") or {}
            tool_calls = message.get("tool_calls") or []
            if not tool_calls:
                content = (message.get("content") or "").strip()
                if not content:
                    raise AIServiceError("DeepSeek returned an empty answer.")
                return content

            assistant_message: dict[str, Any] = {
                "role": "assistant",
                "content": message.get("content"),
                "tool_calls": tool_calls,
            }
            if message.get("reasoning_content") is not None:
                assistant_message["reasoning_content"] = message["reasoning_content"]
            messages.append(assistant_message)

            for call in tool_calls:
                function = call.get("function") or {}
                name = function.get("name") or ""
                try:
                    arguments = json.loads(function.get("arguments") or "{}")
                    if not isinstance(arguments, dict):
                        raise ValueError("Tool arguments must be an object.")
                    result = await tool_executor(name, arguments)
                except (ValueError, TypeError, json.JSONDecodeError) as error:
                    result = {"ok": False, "error": f"Invalid tool arguments: {error}"}
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": call.get("id") or name,
                        "content": json.dumps(result, ensure_ascii=False),
                    }
                )

    raise AIServiceError("DeepSeek used too many tool steps.")
