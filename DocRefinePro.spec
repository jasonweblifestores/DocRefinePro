# DocRefinePro.spec
# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

# 1. Collect all PySide6 bits
datas = []
binaries = []
hiddenimports = []
tmp_ret = collect_all('PySide6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 2. Safe Icon Logic (Fixes Windows Crash)
# If the icon file isn't in the repo, we skip it instead of crashing.
target_icon = 'resources/app_icon.ico'
if not os.path.exists(target_icon):
    print(f"WARNING: {target_icon} not found. Using default icon.")
    target_icon = None

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name='DocRefinePro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=target_icon, 
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DocRefinePro',
)

# 3. Create .app Bundle (Fixes macOS Crash)
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='DocRefinePro.app',
        icon=None,
        bundle_identifier=None,
    )