# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['UNIMUS Kultur - arkeologi - Json til Excel v0.2.py'],
    pathex=[],
    binaries=[],
    datas=[('UiT_Segl_Bok_Bla_30px.png', '.'), ('UiT_Logo_Bok_2l_Bla_RGB.png', '.')],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='UNIMUS Kultur - arkeologi - Json til Excel v0.2',
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
    icon='D:\\favicon2.ico',
)
