# vedveda.spec — PyInstaller build specification
# ================================================

import sys
from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

added_files = [
    (str(ROOT / "templates"),   "templates"),
    (str(ROOT / "static"),      "static"),
    (str(ROOT / "VEDVEDA.xlsx"),"." ),
]

hidden_imports = [
    "flask",
    "flask_sqlalchemy",
    "sqlalchemy",
    "sqlalchemy.dialects.sqlite",
    "sqlalchemy.orm",
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "pystray",
    "pystray._win32",
    "PIL",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL.ImageFont",
    "webview",
    "webview.platforms.winforms",
    "pythonnet",
    "clr",
    "seed",
    "app",
    "jinja2",
    "jinja2.ext",
]

a = Analysis(
    [str(ROOT / "main.py")],
    pathex        = [str(ROOT)],
    binaries      = [],
    datas         = added_files,
    hiddenimports = hidden_imports,
    hookspath     = [],
    hooksconfig   = {},
    runtime_hooks = [],
    excludes      = ["tkinter", "unittest", "test"],
    win_no_prefer_redirects = False,
    win_private_assemblies  = False,
    cipher        = block_cipher,
    noarchive     = False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries = True,
    name             = "VedvedaDashboard",
    debug            = False,
    bootloader_ignore_signals = False,
    strip            = False,
    upx              = True,
    console          = False,
    icon             = str(ROOT / "static" / "icon.ico"),
    version          = str(ROOT / "file_version_info.txt"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip    = False,
    upx      = True,
    upx_exclude = [],
    name     =     "VedvedaDashboard",
)