"""Build the standalone packaged app: production frontend + PyInstaller onedir bundle.

Run from the repo root: python3 packaging/build.py

Wraps a workaround for a real PyInstaller bug on some macOS python.org
installs: sys._base_executable points at an empty dispatcher stub
(python3.X with no arch suffix) instead of the real interpreter binary
(python3.X-intel64 / -arm64), which crashes PyInstaller's dependency
analysis. If your Python doesn't have this problem, the workaround is a
no-op (harmless).
"""
import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent


def _real_base_executable() -> str | None:
    """Resolve sys._base_executable through one more symlink hop if it's an empty stub."""
    base = getattr(sys, "_base_executable", sys.executable)
    if os.path.exists(base) and os.path.getsize(base) > 0:
        return None  # not stubbed, no workaround needed
    real = os.path.realpath(sys.executable)
    if os.path.exists(real) and os.path.getsize(real) > 0:
        return real
    return None


def main() -> None:
    print("Building production frontend (npm run build)...")
    subprocess.run(["npm", "run", "build"], cwd=REPO_ROOT / "addin", check=True)

    fix = _real_base_executable()
    if fix:
        print(f"Working around a stubbed sys._base_executable (using {fix})")
        sys._base_executable = fix

    print("Running PyInstaller...")
    from PyInstaller.__main__ import run

    sys.argv = [
        "pyinstaller",
        str(Path(__file__).parent / "excel-ai-agent.spec"),
        "--distpath", str(Path(__file__).parent / "dist"),
        "--workpath", str(Path(__file__).parent / "build"),
        "--noconfirm",
    ]
    run()
    print(f"\nBuilt: {Path(__file__).parent / 'dist' / 'excel-ai-agent' / 'excel-ai-agent'}")


if __name__ == "__main__":
    main()
