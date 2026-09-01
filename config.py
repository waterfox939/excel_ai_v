"""API key loading and app settings."""
import json
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

USER_CONFIG_DIR = Path.home() / ".excel-ai-agent"
USER_CONFIG_PATH = USER_CONFIG_DIR / "config.json"


def _load_user_config() -> dict:
    if not USER_CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(USER_CONFIG_PATH.read_text())
    except (json.JSONDecodeError, OSError):
        return {}


def save_api_key(key: str) -> None:
    """Persist the API key to the per-user config file (never inside the app bundle)."""
    USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    config = _load_user_config()
    config["anthropic_api_key"] = key
    USER_CONFIG_PATH.write_text(json.dumps(config, indent=2))
    USER_CONFIG_PATH.chmod(0o600)


# Check order: per-user config file (set by the packaged app's first-run
# prompt) first, then .env/environment variable (dev checkout convenience).
ANTHROPIC_API_KEY = _load_user_config().get("anthropic_api_key") or os.environ.get("ANTHROPIC_API_KEY")
MODEL = "claude-sonnet-5"

SESSION_LOG_PATH = "session_log.txt"
