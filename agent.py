"""Anthropic API calls and prompt logic. Claude is the brain; Python is the hands."""
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
