# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['QN_CT_GUI_Figma.py'],
    pathex=[],
    binaries=[],
    datas=[('qn_automailer_settings.json', '.')],
    hiddenimports=['win32com.client', 'pythoncom', 'win32timezone'],
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
    name='QN_Automailer',
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
