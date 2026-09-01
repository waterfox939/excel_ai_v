# -*- mode: python ; coding: utf-8 -*-
# Cross-platform as written — PyInstaller auto-appends .exe on Windows, and
# os.path.join handles separators correctly on either OS. Must be BUILT on
# the target OS though (PyInstaller doesn't cross-compile): run this on a
# real Windows machine to get a Windows .exe, this Mac can only produce a
# macOS build.
import os

import certifi

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(SPEC)), ".."))

a = Analysis(
    ["entry.py"],
    pathex=[REPO_ROOT],
    binaries=[],
    datas=[
        (certifi.where(), "certifi"),  # certifi ships its CA bundle as a data file, not code
        (os.path.join(REPO_ROOT, "addin", "dist"), os.path.join("addin", "dist")),
        (os.path.join(REPO_ROOT, "addin", "manifest.prod.xml"), "addin"),
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
    excludes=[],
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
