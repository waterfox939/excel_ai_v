"""Packaged app entry point. Run once per session (double-click the launcher).

Order: cert setup -> API key prompt if missing -> manifest sideload -> start
the server in the foreground. This is the "make install"-equivalent for a
local, no-cloud install — each step is idempotent, so re-running it (e.g.
after an update) is safe.
"""
import getpass
import shutil
import sys
from pathlib import Path

# Must run before importing server.py / config.py, since those resolve
# paths relative to this file's location — see base_dir() below.
sys.path.insert(0, str(Path(__file__).parent.parent))


def base_dir() -> Path:
    """Directory the packaged app's bundled data (addin/, etc.) lives in.

    PyInstaller sets sys._MEIPASS to wherever it actually placed bundled
    `datas` — for --onedir that's an _internal/ folder next to the exe, not
    the exe's own directory, so derive from _MEIPASS rather than
    sys.executable.
    """
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.parent


def prompt_for_api_key_if_missing() -> None:
    import config

    if config.ANTHROPIC_API_KEY:
        return
    print("No Anthropic API key found.")
    key = getpass.getpass("Paste your Anthropic API key (input hidden): ").strip()
    if not key:
        print("No key entered — exiting.")
        sys.exit(1)
    config.save_api_key(key)
    print(f"Saved to {config.USER_CONFIG_PATH}")
    # config.ANTHROPIC_API_KEY was already read at import time (now stale);
    # re-import agent.py's client after this so it picks up the new key.


def sideload_manifest() -> None:
    manifest_src = base_dir() / "addin" / "manifest.prod.xml"
    wef_dir = Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents" / "wef"
    if not manifest_src.exists():
        print(f"Warning: {manifest_src} not found, skipping sideload.")
        return
    wef_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(manifest_src, wef_dir / "manifest.xml")
    print(f"Sideloaded manifest to {wef_dir}")


def main() -> None:
    from cert_setup import ensure_cert

    print("Setting up HTTPS certificate...")
    cert = ensure_cert()
    if cert is None:
        print("Could not set up a trusted HTTPS cert — Excel likely won't be able to load the task pane.")

    prompt_for_api_key_if_missing()
    sideload_manifest()

    import uvicorn
    from server import PORT, app

    ssl_kwargs = {"ssl_keyfile": str(cert.key_path), "ssl_certfile": str(cert.cert_path)} if cert else {}
    print(f"Starting server on https://localhost:{PORT} — leave this window open while using the Add-in.")
    uvicorn.run(app, host="127.0.0.1", port=PORT, **ssl_kwargs)


if __name__ == "__main__":
    main()
