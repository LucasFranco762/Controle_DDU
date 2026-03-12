# -*- mode: python ; coding: utf-8 -*-

import os

base_dir = os.path.abspath(os.path.dirname(__file__))
icon_ico = os.path.join(base_dir, 'Icone.ico')
icon_png = os.path.join(base_dir, 'Icone.png')
icon_path = icon_ico if os.path.exists(icon_ico) else icon_png


a = Analysis(
    ['Controle_documentos.py'],
    pathex=[],
    binaries=[],
    datas=[('Icone.png', '.'), ('ESCUDO PMMG.png', '.'), ('config.json', '.')],
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
    name='Controle_documentos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
