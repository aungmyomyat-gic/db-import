# Build this spec on 64-bit Windows: python -m PyInstaller --clean --noconfirm DBImport.spec
from pathlib import Path

root = Path(SPECPATH)
a = Analysis(
    [str(root / "desktop.py")],
    pathex=[str(root)],
    binaries=[],
    # Include only the interface; saved credentials and uploaded CSVs stay out.
    datas=[(str(root / "templates" / "index.html"), "templates")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "tkinter"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="db-import",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)
