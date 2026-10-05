"""
build.py — One-shot build script for Vedveda Desktop App
=========================================================
Run:  python build.py

Steps performed:
  1. Generate app icon  (static/icon.png + static/icon.ico)
  2. Run PyInstaller with vedveda.spec
  3. Report the output location

The finished distributable is in:
  dist/VedvedaDashboard/VedvedaDashboard.exe   (folder mode, all DLLs beside it)
"""

import subprocess
import sys
import os
import shutil


ROOT = os.path.abspath(os.path.dirname(__file__))
DIST = os.path.join(ROOT, "dist", "VedvedaDashboard")
EXE  = os.path.join(DIST, "VedvedaDashboard.exe")


def step(msg: str):
    bar = "─" * 60
    print(f"\n{bar}\n  {msg}\n{bar}")


def run(cmd: list[str], **kwargs):
    result = subprocess.run(cmd, **kwargs)
    if result.returncode != 0:
        print(f"[ERROR] Command failed: {' '.join(cmd)}")
        sys.exit(result.returncode)


def main():
    # ── 1. Generate icons ──────────────────────────────────────────────────
    step("Step 1/3 — Generating application icons")
    run([sys.executable, "generate_icon.py"], cwd=ROOT)

    # ── 2. Clean previous build ────────────────────────────────────────────
    step("Step 2/3 — Cleaning previous build artifacts")
    for folder in ["build", "dist"]:
        p = os.path.join(ROOT, folder)
        if os.path.exists(p):
            shutil.rmtree(p)
            print(f"  Removed: {p}")

    # ── 3. Run PyInstaller ─────────────────────────────────────────────────
    step("Step 3/3 — Building EXE with PyInstaller (this may take 2-5 min)")

    # Locate pyinstaller executable
    scripts_dir = os.path.join(
        os.path.dirname(sys.executable), "Scripts"
    )
    pyinstaller_exe = os.path.join(scripts_dir, "pyinstaller.exe")
    if not os.path.exists(pyinstaller_exe):
        # Try user local packages
        import site
        for sp in site.getusersitepackages() if hasattr(site, "getusersitepackages") else []:
            candidate = os.path.join(sp, "..", "Scripts", "pyinstaller.exe")
            if os.path.exists(candidate):
                pyinstaller_exe = candidate
                break
        else:
            # Fall back to module invocation
            pyinstaller_exe = None

    if pyinstaller_exe and os.path.exists(pyinstaller_exe):
        cmd = [pyinstaller_exe, "vedveda.spec", "--noconfirm"]
    else:
        cmd = [sys.executable, "-m", "PyInstaller", "vedveda.spec", "--noconfirm"]

    run(cmd, cwd=ROOT)

    # ── Done ───────────────────────────────────────────────────────────────
    if os.path.exists(EXE):
        size_mb = os.path.getsize(EXE) / 1_048_576
        print(f"""
╔══════════════════════════════════════════════════════════╗
║   BUILD SUCCESSFUL                                      ║
╠══════════════════════════════════════════════════════════╣
║   EXE : {EXE}
║   Size: {size_mb:.1f} MB
║                                                          ║
║   To distribute, copy the entire folder:                ║
║   dist/VedvedaDashboard/                                ║
╚══════════════════════════════════════════════════════════╝
""")
    else:
        print("[WARNING] Build finished but EXE not found. Check PyInstaller output above.")


if __name__ == "__main__":
    main()
