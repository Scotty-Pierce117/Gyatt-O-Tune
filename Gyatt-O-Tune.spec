from pathlib import Path


project_root = Path(SPECPATH).resolve()


a = Analysis(
    [str(project_root / "src" / "scattersyncms" / "main.py")],
    pathex=[str(project_root / "src")],
    binaries=[],
    datas=[
        (
            str(project_root / "src" / "scattersyncms" / "assets" / "gyatt-o-tune.svg"),
            "scattersyncms/assets",
        ),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Gyatt-O-Tune",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)