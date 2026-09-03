# -*- mode: python ; coding: utf-8 -*-
# Cross-platform as written — PyInstaller auto-appends .exe on Windows, and
# os.path.join handles separators correctly on either OS. Must be BUILT on
# the target OS though (PyInstaller doesn't cross-compile), which is what
# .github/workflows/release.yml uses hosted Windows and macOS runners for.
import os

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(SPEC)), ".."))

# certifi ships its CA bundle as a data file rather than code, so PyInstaller
# needs it named explicitly when it's in use. Current anthropic SDKs talk
# through httpx2, which verifies against the OS trust store and doesn't pull
# certifi in at all — so this is conditional. Importing it unconditionally
# hard-failed the build on a clean install of requirements.txt.
try:
    import certifi

    _certifi_datas = [(certifi.where(), "certifi")]
except ImportError:
    _certifi_datas = []

a = Analysis(
    ["entry.py"],
    pathex=[REPO_ROOT],
    binaries=[],
    datas=[
        *_certifi_datas,
        # The frontend ships as source — no bundler, so there is no build
        # output to collect. server.py serves these two directories directly.
        (os.path.join(REPO_ROOT, "addin", "src", "taskpane"), os.path.join("addin", "src", "taskpane")),
        (os.path.join(REPO_ROOT, "addin", "assets"), os.path.join("addin", "assets")),
        (os.path.join(REPO_ROOT, "addin", "manifest.xml"), "addin"),
    ],
    hiddenimports=[
        # uvicorn resolves these dynamically at runtime based on what's
        # installed; PyInstaller's static analysis misses them, producing a
        # ModuleNotFoundError that only surfaces when actually running the
        # built exe. Deliberately not installing uvloop/httptools in the
        # packaging venv keeps this list short (pure-asyncio+h11 uvicorn).
        "uvicorn.loops.auto",
        "uvicorn.protocols.http.auto",
        "uvicorn.protocols.websockets.auto",
        "uvicorn.lifespan.on",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Nothing imports these, but they're large and PyInstaller will happily
    # vacuum them up if a dev venv happens to have them installed. Naming
    # them keeps the download roughly a third of its former size, which is
    # the difference between a quick grab and a discouraging one.
    excludes=["pandas", "numpy", "matplotlib", "tkinter", "PIL", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="excel-ai-agent",
    debug=False,
    strip=False,
    upx=False,
    console=True,  # visible terminal — first-run needs it for the API key prompt
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="excel-ai-agent",
)
