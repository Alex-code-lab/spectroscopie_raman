# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

try:
    spec_dir = os.path.dirname(__file__)
except NameError:
    # __file__ is not set in PyInstaller's spec exec namespace.
    spec_dir = globals().get("specpath", os.getcwd())
project_dir = os.path.abspath(spec_dir)
sys.setrecursionlimit(sys.getrecursionlimit() * 5)
block_cipher = None

datas = []
binaries = []
hiddenimports = []

# --- Qt / PySide6 (inclut QtWebEngine pour QWebEngineView) ---
pkg_datas, pkg_binaries, pkg_hidden = collect_all("PySide6")
datas += pkg_datas
binaries += pkg_binaries
hiddenimports += pkg_hidden
hiddenimports += collect_submodules("PySide6.QtWebEngineWidgets")
hiddenimports += collect_submodules("PySide6.QtWebEngineCore")
hiddenimports += collect_submodules("PySide6.QtWebEngine")

# --- Dépendances scientifiques (sécuriser l'import à l'exécution) ---
hiddenimports += collect_submodules("scipy")
hiddenimports += collect_submodules("sklearn")

# --- Numpy (évite l'import depuis un environnement externe) ---
np_datas, np_binaries, np_hidden = collect_all("numpy")
datas += np_datas
binaries += np_binaries
hiddenimports += np_hidden
hiddenimports += collect_submodules("numpy")

# --- Données locales (si besoin) ---
styles_dir = os.path.join(project_dir, "styles")
if os.path.isdir(styles_dir):
    datas.append((styles_dir, "styles"))

a = Analysis(
    ["main.py"],
    pathex=[project_dir],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pyqt5",
        "PyQt5",
        "PyQt5.QtCore",
        "PyQt5.QtGui",
        "PyQt5.QtWidgets",
        "PyQt5.QtWebEngine",
        "PyQt5.QtWebEngineCore",
        "PyQt5.QtWebEngineWidgets",
        "PyQt5.sip",
        "PyQt6",
        "PyQt6.QtCore",
        "PyQt6.QtGui",
        "PyQt6.QtWidgets",
        "PySide2",
    ],
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
    name="Ramanalyze",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # pas de console sous Windows
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
