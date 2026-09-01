"""Anthropic API calls and prompt logic. Claude is the brain; Python is the hands."""
from collections.abc import Iterator

import anthropic

from config import ANTHROPIC_API_KEY, MODEL

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def send_message(user_message: str) -> str:
    """Send a plain-text message to Claude and return its text reply."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=1024,
        messages=[{"role": "user", "content": user_message}],
    )
    return next(block.text for block in response.content if block.type == "text")


def stream_message(messages: list[dict]) -> Iterator[str]:
    """Stream a reply to a conversation and yield text deltas as they arrive.

    messages: full conversation history, e.g.
        [{"role": "user", "content": "hi"}, {"role": "assistant", "content": "..."}, ...]
    Each message's "content" may be a plain string, or a list of content
    blocks (text/image/document) for messages with attachments.
    The API is stateless, so the caller is responsible for sending the full
    history each time — this function does not remember prior calls.

    This is the simple text-only streaming path (no tool use). server.py's
    /api/chat endpoint calls client.messages.stream(...) directly instead,
    since it needs to handle tool_use content blocks in the response.
    """
    with client.messages.stream(
        model=MODEL,
        max_tokens=1024,
        messages=messages,
    ) as stream:
        yield from stream.text_stream
