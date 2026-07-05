"""API key loading and app settings."""
import os

from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
MODEL = "claude-sonnet-5"

SESSION_LOG_PATH = "session_log.txt"
