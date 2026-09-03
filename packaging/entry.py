"""Packaged app entry point. Run once per session (double-click the launcher).

Order: cert setup -> API key prompt if missing -> manifest sideload -> start
the server in the foreground. This is the "make install"-equivalent for a
local, no-cloud install — each step is idempotent, so re-running it (e.g.
after an update) is safe.
"""
import getpass
import os
import platform
import shutil
import sys
import xml.etree.ElementTree as ET
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


def _sideload_macos(manifest_src: Path) -> None:
    wef_dir = Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents" / "wef"
    wef_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(manifest_src, wef_dir / "manifest.xml")
    print(f"Sideloaded manifest to {wef_dir}")


def _manifest_id(manifest_src: Path) -> str:
    ns = "{http://schemas.microsoft.com/office/appforoffice/1.1}"
    root = ET.parse(manifest_src).getroot()
    id_el = root.find(f"{ns}Id")
    if id_el is None or not id_el.text:
        raise ValueError(f"Could not find <Id> in {manifest_src}")
    return id_el.text.strip()


def _sideload_windows(manifest_src: Path) -> None:
    # Registers the add-in under HKCU\...\WEF\Developer, keyed by the
    # manifest's <Id> GUID with the local file path as the value. This is
    # the mechanism Microsoft's own `npm start` sideload flow uses
    # internally (confirmed from office-addin-dev-settings' source) —
    # closer to macOS's "just drop a file" experience than the officially
    # documented shared-network-folder catalog, which requires an actual
    # network share and a manual Insert > My Add-ins step every time.
    # This key is NOT a publicly documented/guaranteed-stable contract
    # (unlike the TrustedCatalogs registry format), so if Excel ever stops
    # picking this up, the shared-folder-catalog route is the documented
    # fallback — see Microsoft's "network shared folder catalog" docs.
    import winreg

    addin_id = _manifest_id(manifest_src)
    # Persistent copy — don't point the registry at a path that could move
    # (e.g. the folder the user unzipped into, or a temp extraction dir).
    # Resolve via %LOCALAPPDATA% rather than assuming ~/AppData/Local, which
    # is wrong wherever the folder is redirected (roaming profiles, OneDrive
    # Known Folder Move, managed corporate machines).
    local_app_data = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    dest_dir = local_app_data / "ExcelAIAgent"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / "manifest.xml"
    shutil.copy2(manifest_src, dest)

    key_path = r"SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
    key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path)
    winreg.SetValueEx(key, addin_id, 0, winreg.REG_SZ, str(dest))
    winreg.CloseKey(key)
    print(f"Sideloaded manifest to {dest} (registered in HKCU\\{key_path})")


def sideload_manifest() -> None:
    manifest_src = base_dir() / "addin" / "manifest.xml"
    if not manifest_src.exists():
        print(f"Warning: {manifest_src} not found, skipping sideload.")
        return
    system = platform.system()
    if system == "Darwin":
        _sideload_macos(manifest_src)
    elif system == "Windows":
        _sideload_windows(manifest_src)
    else:
        print(f"Automatic sideload not implemented for {system} — see README for the manual steps.")


def main() -> None:
    # Line-buffer stdout. PyInstaller block-buffers it whenever it is not a
    # real console (piped, redirected to a log file), which would hide the
    # first-run progress messages — including the warning that a cert dialog
    # is about to appear — until the process exits.
    sys.stdout.reconfigure(line_buffering=True)

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
    scheme = "https" if cert else "http"
    print(f"Starting server on {scheme}://localhost:{PORT} — leave this window open while using the Add-in.")
    uvicorn.run(app, host="127.0.0.1", port=PORT, **ssl_kwargs)


if __name__ == "__main__":
    main()
