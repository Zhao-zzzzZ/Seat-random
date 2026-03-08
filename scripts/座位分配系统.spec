# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

项目根目录 = Path(SPECPATH).parent

a = Analysis(
    [str(项目根目录 / 'src' / '座位分配.py')],
    pathex=[str(项目根目录)],
    binaries=[],
    datas=[],
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
    name='座位分配系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=str(项目根目录 / 'scripts' / 'version.txt'),
    uac_admin=True,
)
