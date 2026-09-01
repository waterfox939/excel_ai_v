"""Self-signed HTTPS cert for the packaged local server, trusted via the macOS login keychain.

Office Add-ins require HTTPS to load the task pane, even for a purely local
install. This generates a self-signed cert (once) via the system `openssl`
CLI — no compiled Python dependency needed, which also keeps it out of the
PyInstaller bundle — and trusts it in the user's login keychain
(deliberately the login keychain, not the system keychain, so no
sudo/admin password is needed for a first-run flow).
"""
import platform
import subprocess
from dataclasses import dataclass
from pathlib import Path

CERT_DIR = Path.home() / ".excel-ai-agent" / "certs"
CERT_PATH = CERT_DIR / "localhost.crt"
KEY_PATH = CERT_DIR / "localhost.key"


@dataclass
class Cert:
    cert_path: Path
    key_path: Path


def _generate(cert_path: Path, key_path: Path) -> None:
    cert_path.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "openssl", "req", "-x509", "-newkey", "rsa:2048",
            "-keyout", str(key_path), "-out", str(cert_path),
            "-days", "825",  # macOS/browser-enforced validity cap
            "-nodes",  # no passphrase — this key must be readable non-interactively at startup
            "-subj", "/CN=localhost",
            "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
        ],
        check=True,
        capture_output=True,
    )
    key_path.chmod(0o600)


def _trust(cert_path: Path) -> None:
    if platform.system() != "Darwin":
        print(f"Cert generated at {cert_path}. Trust it manually on this OS to use HTTPS.")
        return
    login_keychain = Path.home() / "Library" / "Keychains" / "login.keychain-db"
    subprocess.run(
        [
            "security", "add-trusted-cert", "-d", "-r", "trustRoot",
            "-k", str(login_keychain), str(cert_path),
        ],
        check=True,
    )


def ensure_cert() -> Cert | None:
    """Ensure a trusted self-signed localhost cert exists; generate + trust it on first run.

    Returns None if cert generation fails, so the caller can fall back to
    plain HTTP rather than crash outright — HTTPS is required for the
    Office Add-in to actually load the task pane, but a broken cert setup
    shouldn't prevent the server from starting at all for debugging.
    """
    try:
        if not (CERT_PATH.exists() and KEY_PATH.exists()):
            _generate(CERT_PATH, KEY_PATH)
            _trust(CERT_PATH)
        return Cert(cert_path=CERT_PATH, key_path=KEY_PATH)
    except Exception as exc:
        print(f"Could not set up HTTPS cert ({exc}); falling back to plain HTTP.")
        return None
