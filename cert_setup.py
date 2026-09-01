"""Self-signed HTTPS cert for the packaged local server, trusted per-user (no admin/sudo).

Office Add-ins require HTTPS to load the task pane, even for a purely local
install. This generates a self-signed cert (once) and trusts it in the
CURRENT USER's certificate store — deliberately not the system-wide store,
so no admin password is needed for a first-run flow.

macOS: generated via the system `openssl` CLI (ships on every Mac, no pip
install needed), trusted via `security` into the login keychain — verified
silent (no dialog) on this machine.

Windows: generated via the `cryptography` pip package (installs cleanly
from a prebuilt wheel on Windows — unlike some macOS dev setups, no Rust
toolchain needed there) since PowerShell's New-SelfSignedCertificate can't
cleanly export a plain-PEM private key without real pain. Trusted via
`certutil -user -addstore Root`, targeting CurrentUser so no admin
elevation is needed — reported silent by several deployment/sysadmin
sources, but NOT verified firsthand on a real Windows machine. If it turns
out to show a one-time confirmation dialog in practice, that's expected
and acceptable (same posture as a password prompt), not a bug to chase.
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


def _generate_macos(cert_path: Path, key_path: Path) -> None:
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


def _trust_macos(cert_path: Path) -> None:
    login_keychain = Path.home() / "Library" / "Keychains" / "login.keychain-db"
    subprocess.run(
        [
            "security", "add-trusted-cert", "-d", "-r", "trustRoot",
            "-k", str(login_keychain), str(cert_path),
        ],
        check=True,
    )


def _generate_windows(cert_path: Path, key_path: Path) -> None:
    # Deliberately not shelling out to PowerShell's New-SelfSignedCertificate
    # here: it stores the key CNG-backed in the certificate store, and
    # exporting that to a plain unencrypted PEM key (what Python's ssl
    # module / uvicorn need) is genuinely painful without bundling openssl.
    # `cryptography` installs from a prebuilt wheel on Windows (no compiler
    # needed there, unlike the Rust-toolchain problem hit on some macOS
    # setups), so generate the PEM files directly with it instead.
    import datetime
    import ipaddress

    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
    san = x509.SubjectAlternativeName(
        [x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_address("127.0.0.1"))]
    )
    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + datetime.timedelta(days=825))
        .add_extension(san, critical=False)
        .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
        .sign(key, hashes.SHA256())
    )

    cert_path.parent.mkdir(parents=True, exist_ok=True)
    key_path.write_bytes(
        key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))


def _trust_windows(cert_path: Path) -> None:
    # -user targets CurrentUser\Root (no admin elevation). Reported silent
    # by several sysadmin/deployment sources, but not independently
    # verified on a real machine — see module docstring.
    subprocess.run(
        ["certutil", "-f", "-user", "-addstore", "Root", str(cert_path)],
        check=True,
        capture_output=True,
    )


def ensure_cert() -> Cert | None:
    """Ensure a trusted self-signed localhost cert exists; generate + trust it on first run.

    Returns None if cert generation fails, so the caller can fall back to
    plain HTTP rather than crash outright — HTTPS is required for the
    Office Add-in to actually load the task pane, but a broken cert setup
    shouldn't prevent the server from starting at all for debugging.
    """
    system = platform.system()
    try:
        if not (CERT_PATH.exists() and KEY_PATH.exists()):
            if system == "Darwin":
                _generate_macos(CERT_PATH, KEY_PATH)
                _trust_macos(CERT_PATH)
            elif system == "Windows":
                _generate_windows(CERT_PATH, KEY_PATH)
                _trust_windows(CERT_PATH)
            else:
                print(f"Cert generation not implemented for {system}. Trust it manually to use HTTPS.")
                return None
        return Cert(cert_path=CERT_PATH, key_path=KEY_PATH)
    except Exception as exc:
        print(f"Could not set up HTTPS cert ({exc}); falling back to plain HTTP.")
        return None
