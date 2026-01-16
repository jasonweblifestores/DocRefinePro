# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

# ==============================================================================
#   DOCREFINE PRO v128 BUILD SPEC
#   Optimized for PySide6 bloat reduction (No WebEngine, No QML)
# ==============================================================================

# 1. Collect PySide6, but keep reference so we can filter it
#    collect_all returns: (datas, binaries, hiddenimports)
pyside_datas, pyside_binaries, pyside_hidden = collect_all('PySide6')

# 2. Define Exclusion Filters (Substrings to match against filenames)
#    If a binary contains any of these, it dies.
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
    'opengl32sw', 'd3dcompiler',  # Windows software renderers (heavy)
]

def filter_binaries(bin_list):
    """Returns a filtered list of binaries, removing unwanted Qt modules."""
    kept = []
    dropped_size = 0
    for src, dst, kind in bin_list:
        # Check exclusion
        if any(bad in src for bad in EXCLUSION_PATTERNS):
            print(f"  [Spec Filter] Dropping: {os.path.basename(src)}")
            continue
        kept.append((src, dst, kind))
    return kept

# 3. Apply Filter
filtered_binaries = filter_binaries(pyside_binaries)

# 4. Standard App Setup
block_cipher = None

# Safe Icon Logic
target_icon = 'resources/app_icon.ico'
if not os.path.exists(target_icon):
    target_icon = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=filtered_binaries, # USE FILTERED LIST
    datas=pyside_datas,         # Keep datas (translations etc are small enough)
    hiddenimports=[
        'docrefine.gui.app_qt', # Ensure our new GUI entry point is caught
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
    strip=True,     # Strip symbols
    upx=True,       # UPX Compression
    console=False,  # No Terminal
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