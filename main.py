"""
main.py — Vedveda Desktop Application Entry Point
===================================================
Boots the Flask server in a background thread, then opens a native
WebView window so the app behaves like a proper desktop application.

Run directly:  python main.py
Build to EXE:  python build.py
"""

import sys
import os
import threading
import time
import socket
import logging

# ── Suppress Flask/Werkzeug console output in packaged mode ─────────────────
logging.getLogger("werkzeug").setLevel(logging.ERROR)

# ── Path helpers (works both in dev and when frozen by PyInstaller) ──────────
def resource_path(relative: str) -> str:
    """Return the absolute path to a bundled resource."""
    if getattr(sys, "frozen", False):
        # PyInstaller unpacks to sys._MEIPASS
        base = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base, relative)


def data_path(relative: str) -> str:
    """
    Return the path for writable runtime data (DB, logs).
    When frozen, writes next to the .exe; in dev, next to main.py.
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base, relative)


# ── Patch paths BEFORE importing app so SQLite DB lands in the right place ──
os.environ["VEDVEDA_DATA_DIR"] = data_path("")
os.environ["VEDVEDA_RES_DIR"]  = resource_path("")

# ── Imports that depend on the patched environment ──────────────────────────
import webview                      # pywebview
from app import app, db, Member, Transaction

# ────────────────────────────────────────────────────────────────────────────
# Configuration
# ────────────────────────────────────────────────────────────────────────────
APP_TITLE   = "Vedveda Sales & Analytics Dashboard"
APP_WIDTH   = 1280
APP_HEIGHT  = 800
APP_HOST    = "127.0.0.1"
APP_PORT    = 5055          # fixed port; avoids conflict with dev Flask on 5000
ICON_PATH   = resource_path(os.path.join("static", "icon.png"))


# ────────────────────────────────────────────────────────────────────────────
# Flask server thread
# ────────────────────────────────────────────────────────────────────────────

def _find_free_port(preferred: int) -> int:
    """Return `preferred` if it's free, otherwise pick any free port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex((APP_HOST, preferred)) != 0:
            return preferred
    with socket.socket() as s:
        s.bind(("", 0))
        return s.getsockname()[1]


def _run_flask(port: int) -> None:
    """Start Flask in production mode (no reloader, no debugger)."""
    # Initialise the DB and seed if empty
    with app.app_context():
        # Override DB path so it always lives next to the .exe
        data_dir = os.environ.get("VEDVEDA_DATA_DIR", "")
        db_path  = os.path.join(data_dir, "vedveda.db")
        app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
        db.create_all()

        if Member.query.count() == 0:
            print("[INFO] First run — seeding database from VEDVEDA.xlsx …")
            try:
                from seed import seed_database
                xlsx = os.path.join(
                    os.environ.get("VEDVEDA_RES_DIR", ""), "VEDVEDA.xlsx"
                )
                # Temporarily patch EXCEL_PATH in seed module
                import seed as seed_mod
                seed_mod.EXCEL_PATH = xlsx
                seed_database(db, Member, Transaction)
                print("[INFO] Seed complete.")
            except Exception as exc:
                print(f"[WARNING] Seeding skipped: {exc}")

    app.run(host=APP_HOST, port=port, debug=False, use_reloader=False, threaded=True)


def _wait_for_server(port: int, timeout: float = 10.0) -> bool:
    """Block until Flask is accepting connections (or timeout)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((APP_HOST, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.1)
    return False


# ────────────────────────────────────────────────────────────────────────────
# System-tray icon (pystray)
# ────────────────────────────────────────────────────────────────────────────

def _build_tray_icon(window):
    """Create and return a pystray Icon object."""
    try:
        import pystray
        from PIL import Image

        # Load app icon or generate a fallback coloured square
        try:
            img = Image.open(ICON_PATH).resize((64, 64))
        except Exception:
            img = Image.new("RGB", (64, 64), color="#0ea5e9")

        def _show(icon, item):
            window.show()

        def _hide(icon, item):
            window.hide()

        def _quit(icon, item):
            icon.stop()
            window.destroy()

        menu = pystray.Menu(
            pystray.MenuItem("Show Dashboard", _show, default=True),
            pystray.MenuItem("Hide to Tray",   _hide),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit",            _quit),
        )
        return pystray.Icon("Vedveda", img, APP_TITLE, menu)
    except ImportError:
        return None          # pystray not available; tray icon skipped


# ────────────────────────────────────────────────────────────────────────────
# WebView JS API (callable from the browser page via window.pywebview.api)
# ────────────────────────────────────────────────────────────────────────────

class JsApi:
    """Exposes Python functions to JavaScript running inside the WebView."""

    def __init__(self, win):
        self._win = win

    def minimize(self):
        self._win.minimize()

    def toggle_fullscreen(self):
        self._win.toggle_fullscreen()

    def get_app_version(self):
        return "1.0.0"


# ────────────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────────────

def main():
    port = _find_free_port(APP_PORT)
    url  = f"http://{APP_HOST}:{port}"

    # 1. Start Flask in a daemon thread
    flask_thread = threading.Thread(
        target=_run_flask, args=(port,), daemon=True, name="FlaskThread"
    )
    flask_thread.start()

    # 2. Wait until Flask is ready
    ready = _wait_for_server(port)
    if not ready:
        # Show a plain error window if Flask never started
        webview.create_window(
            "Vedveda — Error",
            html="<h2 style='font-family:sans-serif;color:red'>Server failed to start.</h2>",
        )
        webview.start()
        return

    # 3. Create the WebView window
    js_api = JsApi(None)        # placeholder; patched below after window created
    window = webview.create_window(
        title        = APP_TITLE,
        url          = url,
        width        = APP_WIDTH,
        height       = APP_HEIGHT,
        min_size     = (900, 600),
        resizable    = True,
        text_select  = True,
        js_api       = js_api,
    )
    js_api._win = window        # patch back reference

    # 4. System-tray icon in its own thread
    tray = _build_tray_icon(window)
    if tray:
        tray_thread = threading.Thread(target=tray.run, daemon=True, name="TrayThread")
        tray_thread.start()

    # 5. Start the WebView event loop (blocks until window is closed)
    webview.start(debug=False)

    # 6. Clean up tray on exit
    if tray:
        try:
            tray.stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
