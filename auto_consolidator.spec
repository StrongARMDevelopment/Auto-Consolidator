# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\AaronMelton\\OneDrive\\Documents\\Programs\\AutoConsolidation\\Auto Consolidate\\auto_consolidator.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl', 'pandas', 'tkinter', 'ttkthemes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'IPython', 'jupyter', 'notebook'],
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
    name='Auto_Consolidator',
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
    icon=['C:\\Users\\AaronMelton\\OneDrive\\Documents\\Programs\\AutoConsolidation\\Auto Consolidate\\auto_consolidator.ico'],
)
