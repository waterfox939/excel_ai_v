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
elevation is needed. Adding to the root store shows a one-time confirmation
dialog; that's expected (same posture as a password prompt), not a bug to
chase — but it means the user can decline, which is why trust is verified
independently of generation on every run. See ensure_cert().
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


def _is_trusted_macos(cert_path: Path) -> bool:
    # verify-cert only succeeds if the chain validates, which for a
    # self-signed cert means it's actually trusted as a root.
    result = subprocess.run(
        ["security", "verify-cert", "-c", str(cert_path)],
        capture_output=True,
        check=False,
    )
    return result.returncode == 0


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
    # -user targets CurrentUser\Root (no admin elevation). Windows shows a
    # one-time "you are about to install a certificate" confirmation for
    # root-store additions; that dialog is expected, and this call blocks
    # until the user answers it.
    subprocess.run(
        ["certutil", "-f", "-user", "-addstore", "Root", str(cert_path)],
        check=True,
        capture_output=True,
    )


def _thumbprint(cert_path: Path) -> str:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes

    cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
    return cert.fingerprint(hashes.SHA1()).hex().upper()


def _is_trusted_windows(cert_path: Path) -> bool:
    # Look the cert up by thumbprint in CurrentUser\Root. Checking before
    # re-adding is what keeps the confirmation dialog to genuinely once —
    # certutil -f would happily re-prompt on every launch.
    result = subprocess.run(
        ["certutil", "-user", "-verifystore", "Root", _thumbprint(cert_path)],
        capture_output=True,
        check=False,
    )
    return result.returncode == 0


def _manual_trust_hint(cert_path: Path, system: str) -> str:
    if system == "Windows":
        return f'certutil -f -user -addstore Root "{cert_path}"'
    return f'security add-trusted-cert -d -r trustRoot -k ~/Library/Keychains/login.keychain-db "{cert_path}"'


def ensure_cert() -> Cert | None:
    """Ensure a trusted self-signed localhost cert exists; generate + trust it on first run.

    Generation and trust are checked independently on every run, and that
    separation matters: the two steps can fail apart from each other. If the
    user dismisses Windows' root-certificate confirmation dialog, the PEM
    files are already on disk, so keying the trust step off "do the files
    exist?" would mean it never ran again — leaving an untrusted cert that
    silently breaks the task pane on every subsequent launch.

    Returns None only if generation itself fails, so the caller can fall back
    to plain HTTP rather than crash outright. A cert that exists but isn't
    trusted is still returned (with a loud warning): HTTPS-with-a-warning is
    closer to working than no HTTPS at all, and the message tells the user
    exactly how to fix it.
    """
    system = platform.system()
    if system not in ("Darwin", "Windows"):
        print(f"Cert setup not implemented for {system}. Generate and trust a localhost cert manually to use HTTPS.")
        return None

    try:
        if not (CERT_PATH.exists() and KEY_PATH.exists()):
            if system == "Darwin":
                _generate_macos(CERT_PATH, KEY_PATH)
            else:
                _generate_windows(CERT_PATH, KEY_PATH)
    except Exception as exc:
        print(f"Could not generate an HTTPS cert ({exc}); falling back to plain HTTP.")
        return None

    cert = Cert(cert_path=CERT_PATH, key_path=KEY_PATH)
    try:
        is_trusted = _is_trusted_macos(CERT_PATH) if system == "Darwin" else _is_trusted_windows(CERT_PATH)
        if not is_trusted:
            print("Trusting the local HTTPS certificate (you may see a one-time confirmation prompt)...")
            if system == "Darwin":
                _trust_macos(CERT_PATH)
            else:
                _trust_windows(CERT_PATH)
    except Exception as exc:
        print(
            f"\n  WARNING: could not trust the local HTTPS certificate ({exc}).\n"
            f"  Excel will refuse to load the task pane until it is trusted.\n"
            f"  This run will retry automatically; to do it by hand:\n"
            f"    {_manual_trust_hint(CERT_PATH, system)}\n"
        )
    return cert
