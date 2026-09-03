"""API key loading and app settings."""
import getpass
import json
import os
import subprocess
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


def _restrict_to_current_user(path: Path) -> None:
    """Best-effort: make a file readable only by the user who owns it.

    Windows needs its own path here. os.chmod on Windows only toggles the
    read-only bit — it ignores the owner/group/other bits entirely, so
    chmod(0o600) is a silent no-op that leaves an API key readable by every
    other account on the machine. icacls is the actual equivalent:
    /inheritance:r drops the inherited ACEs, then the current user is
    granted full control alone. Failures are non-fatal — a slightly
    over-permissive config file shouldn't stop the app from starting.
    """
    if os.name == "nt":
        subprocess.run(
            ["icacls", str(path), "/inheritance:r", "/grant:r", f"{getpass.getuser()}:F"],
            capture_output=True,
            check=False,
        )
    else:
        path.chmod(0o600)


def save_api_key(key: str) -> None:
    """Persist the API key to the per-user config file (never inside the app bundle)."""
    USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    config = _load_user_config()
    config["anthropic_api_key"] = key
    USER_CONFIG_PATH.write_text(json.dumps(config, indent=2))
    _restrict_to_current_user(USER_CONFIG_PATH)


# Check order: per-user config file (set by the packaged app's first-run
# prompt) first, then .env/environment variable (dev checkout convenience).
ANTHROPIC_API_KEY = _load_user_config().get("anthropic_api_key") or os.environ.get("ANTHROPIC_API_KEY")
MODEL = "claude-sonnet-5"

SESSION_LOG_PATH = "session_log.txt"
