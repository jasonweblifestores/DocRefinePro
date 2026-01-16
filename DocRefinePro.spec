# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

# ... (Standard headers)

# 1. Collect PySide6
pyside_datas, pyside_binaries, pyside_hidden = collect_all('PySide6')

# 2. Define Exclusion List
EXCLUSION_PATTERNS = [
    'PySide6.QtWebEngine', 'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
    'QtWebEngine', 'QtWebEngineCore', 'QtWebEngineWidgets', 'QtWebEngineQuick',
    'PySide6.QtQuick', 'PySide6.QtQuickWidgets', 'PySide6.QtQuick3D', 'PySide6.QtQuickControls2',
    'QtQuick', 'QtQuickWidgets', 'QtQuick3D', 'QtQuickControls2',
    'PySide6.QtQml', 'PySide6.QtQml.WorkerScript', 'QtQml',
    'PySide6.Qt3DCore', 'PySide6.Qt3DInput', 'PySide6.Qt3DLogic', 'PySide6.Qt3DRender', 'PySide6.Qt3DAnimation', 'PySide6.Qt3DExtras',
    'Qt3DCore', 'Qt3DInput', 'Qt3DLogic', 'Qt3DRender', 'Qt3DAnimation', 'Qt3DExtras',
    'PySide6.QtVirtualKeyboard', 'QtVirtualKeyboard',
    'PySide6.QtSerialBus', 'PySide6.QtSerialPort', 'QtSerialBus', 'QtSerialPort',
    'PySide6.QtSensors', 'QtSensors',
    'PySide6.QtCharts', 'QtCharts',
    'PySide6.QtDataVisualization', 'QtDataVisualization',
    'PySide6.QtTest', 'QtTest',
    'PySide6.QtTextToSpeech', 'QtTextToSpeech',
    'PySide6.QtDesigner', 'QtDesigner',
    'PySide6.QtHelp', 'QtHelp',
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets', 'QtMultimedia', 'QtMultimediaWidgets',
    'PySide6.QtLocation', 'QtLocation',
    'PySide6.QtPositioning', 'QtPositioning',
    'PySide6.QtNetworkAuth', 'QtNetworkAuth',
    'PySide6.QtScxml', 'QtScxml',
    'PySide6.QtRemoteObjects', 'QtRemoteObjects',
    'PySide6.QtStateMachine', 'QtStateMachine',
    'opengl32sw', 'd3dcompiler'
]

# 3. Filter Binaries (CRASH FIX: Handle tuple sizes safely)
filtered_binaries = []
for binary in pyside_binaries:
    is_bad = False
    for item in binary:
        if isinstance(item, str) and any(bad in item for bad in EXCLUSION_PATTERNS):
            is_bad = True
            break
    if not is_bad:
        filtered_binaries.append(binary)

# 4. Filter Hidden Imports
filtered_hidden_imports = [
    h for h in pyside_hidden 
    if not any(bad in h for bad in EXCLUSION_PATTERNS)
]

block_cipher = None
target_icon = 'resources/app_icon.ico'
if not os.path.exists(target_icon): target_icon = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=filtered_binaries, 
    datas=pyside_datas,
    hiddenimports=[
        'docrefine.gui.app_qt', 
        'docrefine.processing', 
        'docrefine.worker'
    ] + filtered_hidden_imports, 
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUSION_PATTERNS,
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