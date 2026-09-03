"""Build the standalone packaged app: a PyInstaller onedir bundle.

Run from the repo root:  python packaging/build.py

There is no frontend build step. addin/src/taskpane/ is plain browser JS
with no imports, so it is bundled and served as-is — which is why this
project needs no Node.js, no npm, and no bundler to produce a release.

PyInstaller does not cross-compile: a Windows .exe must be built on
Windows and a macOS binary on macOS. Neither needs a developer machine —
.github/workflows/release.yml runs this on GitHub's hosted runners for
both platforms and attaches the results to a Release.

Wraps a workaround for a real PyInstaller bug on some macOS python.org
installs: sys._base_executable points at an empty dispatcher stub
(python3.X with no arch suffix) instead of the real interpreter binary
(python3.X-intel64 / -arm64), which crashes PyInstaller's dependency
analysis. On Windows, and on Pythons without this problem, it's a no-op.
"""
import os
import platform
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

    exe_name = "excel-ai-agent.exe" if platform.system() == "Windows" else "excel-ai-agent"
    print(f"\nBuilt: {Path(__file__).parent / 'dist' / 'excel-ai-agent' / exe_name}")


if __name__ == "__main__":
    main()
