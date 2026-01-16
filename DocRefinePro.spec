# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

# ==============================================================================
#   DOCREFINE PRO v128 BUILD SPEC (Robust Filter Edition)
#   Optimized for PySide6 bloat reduction
# ==============================================================================

# 1. Collect PySide6
pyside_datas, pyside_binaries, pyside_hidden = collect_all('PySide6')

# 2. Exclusion Filters
EXCLUSION_PATTERNS = [
    'Qt6WebEngine', 'QtWebEngine', 
    'Qt6Quick', 'QtQuick', 
    'Qt6Qml', 'QtQml', 
    'Qt63D', 'Qt3D',
    'Qt6VirtualKeyboard', 'QtVirtualKeyboard',
    'Qt6SerialBus', 'QtSerialBus',
    'Qt6Designer', 'QtDesigner',
    'Qt6Help', 'QtHelp',
    'Qt6Test', 'QtTest',
    'Qt6Sensors', 'QtSensors',
    'Qt6Charts', 'QtCharts',
    'Qt6DataVisualization', 'QtDataVisualization',
    'opengl32sw', 'd3dcompiler',
]

def filter_binaries(bin_list):
    """Returns a filtered list of binaries, robust to tuple size."""
    kept = []
    for binary in bin_list:
        # binary might be (dest, src) or (dest, src, type)
        # We iterate to find if ANY part of the tuple matches our ban list
        is_bad = False
        for item in binary:
            if isinstance(item, str) and any(bad in item for bad in EXCLUSION_PATTERNS):
                print(f"  [Spec Filter] Dropping: {os.path.basename(item)}")
                is_bad = True
                break
        
        if not is_bad:
            kept.append(binary)
    return kept

# 3. Apply Filter
filtered_binaries = filter_binaries(pyside_binaries)

# 4. Standard App Setup
block_cipher = None

target_icon = 'resources/app_icon.ico'
if not os.path.exists(target_icon):
    target_icon = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=filtered_binaries, 
    datas=pyside_datas,
    hiddenimports=[
        'docrefine.gui.app_qt', 
        'docrefine.processing',
        'docrefine.worker'
    ] + pyside_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 
        'matplotlib', 
        'scipy', 
        'notebook', 
        'pandas'
    ],
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
    strip=True,
    upx=True,
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