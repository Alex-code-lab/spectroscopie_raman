# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

project_dir = os.path.abspath(os.path.dirname(__file__))
block_cipher = None

datas = []
binaries = []
hiddenimports = []

# --- Qt / PySide6 (inclut QtWebEngine pour QWebEngineView) ---
for pkg in ("PySide6", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineCore"):
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

# --- Dépendances scientifiques (sécuriser l'import à l'exécution) ---
hiddenimports += collect_submodules("scipy")
hiddenimports += collect_submodules("sklearn")

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
    excludes=[],
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
