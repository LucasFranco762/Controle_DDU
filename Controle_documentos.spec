# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os

# SPECPATH é uma variável interna do PyInstaller que aponta para a pasta deste .spec
base_dir = SPECPATH

# ── Recursos visuais empacotados dentro do executável (lidos via sys._MEIPASS) ──
datas = [
    (os.path.join(base_dir, 'Icone.png'),       '.'),
    (os.path.join(base_dir, 'ESCUDO PMMG.png'), '.'),
    (os.path.join(base_dir, 'check_blue.svg'),  '.'),
]

binaries = []

hiddenimports = [
    'sqlite3',
    'requests',
    'requests.adapters',
    'requests.auth',
    'urllib3',
    'Crypto',
    'Crypto.Cipher',
    'Crypto.Cipher.AES',
    'Crypto.Util',
    'Crypto.Util.Padding',
]

# Coleta completa de todas as dependências externas (incluindo plugins Qt, etc.)
for pkg in ['PySide6', 'plotly', 'kaleido', 'reportlab', 'Crypto']:
    tmp = collect_all(pkg)
    datas         += tmp[0]
    binaries      += tmp[1]
    hiddenimports += tmp[2]

a = Analysis(
    [os.path.join(base_dir, 'Controle_documentos.py')],
    pathex=[base_dir],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='Controle_DDU',
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
    icon=os.path.join(base_dir, 'Icone.png'),
)
