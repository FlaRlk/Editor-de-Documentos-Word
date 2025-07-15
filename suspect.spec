# -*- mode: python ; coding: utf-8 -*-

import os

# Obter caminho absoluto para o ícone
icon_path = os.path.abspath('icon.ico')

a = Analysis(
    ['suspect.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.')],  # Incluir o ícone como um recurso
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
    name='suspect_word_edit',
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
    icon=icon_path,
)
