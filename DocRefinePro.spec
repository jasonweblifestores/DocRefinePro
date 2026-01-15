# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_data_files

# 1. OPTIMIZATION: Targeted PySide6 Collection
# Fixed: Removed 'include_pycache' which is unsupported in PyInstaller 6.18+
hiddenimports = ['PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets']
datas = collect_data_files('PySide6')

# 2. OPTIMIZATION: Explicit Exclusions
# Removing unused Chromium and 3D modules saves ~450MB
excluded_modules = [
    'PySide6.QtWebEngine', 'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets',
    'PySide6.Qt3D', 'PySide6.QtDesigner', 'PySide6.QtNetwork', 'PySide6.QtSql',
    'PySide6.QtQuick', 'PySide6.QtQml', 'PySide6.QtVirtualKeyboard'
]

# 3. Safe Icon Logic
target_icon = 'resources/app_icon.ico'
if not os.path.exists(target_icon):
    target_icon = None

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_modules,
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
    strip=True,     # Discards symbols to reduce size
    upx=True,       # High compression for binaries
    console=False,
    icon=target_icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=True,
    upx=True,
    name='DocRefinePro',
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='DocRefinePro.app',
        icon=None,
        bundle_identifier='com.docrefine.pro',
    )