# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_kontrolor.py'],
    pathex=[],
    binaries=[],
    datas=[('logo_1.png', '.'), ('logo.png', '.'), ('logo_1.ico', '.')],
    hiddenimports=['openpyxl', 'openpyxl.cell._writer', 'openpyxl.styles.stylesheet'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GabimKontrolProgram',
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
    icon='logo.ico',
    version_file='version_info.txt'
)
