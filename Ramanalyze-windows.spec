# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_submodules

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

# --- Qt / PySide6 (inclut QtWebEngine pour QWebEngineView) -------------------
# IMPORTANT pour la taille : on N'utilise PAS collect_all("PySide6"), qui
# embarquerait TOUS les modules Qt (Qt3D, Charts, Multimedia, Quick3D,
# Designer…) et gonflerait l'exe de plusieurs centaines de Mo. Les hooks
# PySide6 de PyInstaller n'incluent que les modules réellement importés + leurs
# dépendances, et embarquent automatiquement les données QtWebEngine
# (QtWebEngineProcess, icudtl.dat, locales, resources). On se contente de
# déclarer les seuls modules Qt que l'application importe.
hiddenimports += [
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "PySide6.QtWebChannel",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
]

# --- Dépendances scientifiques : uniquement les sous-paquets utilisés --------
# scipy : optimize.curve_fit, stats.norm/spearmanr, ndimage.uniform_filter1d,
#         signal.find_peaks (on n'embarque donc pas io, integrate, interpolate,
#         spatial, fft… ni les suites de tests).
for _scipy_sub in ("scipy.optimize", "scipy.stats", "scipy.ndimage", "scipy.signal"):
    hiddenimports += collect_submodules(_scipy_sub)

# sklearn : seulement PCA (sklearn.decomposition).
hiddenimports += collect_submodules("sklearn.decomposition")
hiddenimports += collect_submodules("sklearn.utils")

# pybaselines : correction de ligne de base (Baseline).
hiddenimports += collect_submodules("pybaselines")

# plotly : on garde ses sous-modules (validateurs de figures requis au runtime).
hiddenimports += collect_submodules("plotly")

# numpy / pandas / openpyxl : pris en charge par les hooks intégrés de
# PyInstaller (hook-numpy, hook-pandas…), sans embarquer tests ni données
# superflues. fpdf n'est pas importé par l'app : on ne le collecte plus.

# --- Données locales (ressources de l'app dans ramanalyze/assets) ---
assets_dir = os.path.join(project_dir, "ramanalyze", "assets")

styles_dir = os.path.join(assets_dir, "styles")
if os.path.isdir(styles_dir):
    datas.append((styles_dir, os.path.join("assets", "styles")))

logo_dir = os.path.join(assets_dir, "Logo Ramanalyze")
if os.path.isdir(logo_dir):
    datas.append((logo_dir, os.path.join("assets", "Logo Ramanalyze")))

# --- Modèle Excel de mise en page (feuille de protocole) ---
template_xlsx = os.path.join(assets_dir, "modele_tableau.xlsx")
if os.path.isfile(template_xlsx):
    datas.append((template_xlsx, "assets"))

a = Analysis(
    [os.path.join("ramanalyze", "main.py")],
    pathex=[project_dir, os.path.join(project_dir, "ramanalyze")],
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
        # --- Modules Qt non utilisés (gros gains de taille) ---
        # ⚠ Ne PAS exclure QtQuick/QtQml/QtNetwork/QtPositioning/QtOpenGL/
        #   QtPrintSupport : ce sont des dépendances internes de QtWebEngine.
        "PySide6.Qt3DAnimation",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DExtras",
        "PySide6.Qt3DInput",
        "PySide6.Qt3DLogic",
        "PySide6.Qt3DRender",
        "PySide6.QtCharts",
        "PySide6.QtDataVisualization",
        "PySide6.QtMultimedia",
        "PySide6.QtMultimediaWidgets",
        "PySide6.QtSpatialAudio",
        "PySide6.QtQuick3D",
        "PySide6.QtBluetooth",
        "PySide6.QtNfc",
        "PySide6.QtSensors",
        "PySide6.QtSerialPort",
        "PySide6.QtSerialBus",
        "PySide6.QtSql",
        "PySide6.QtTest",
        "PySide6.QtDesigner",
        "PySide6.QtHelp",
        "PySide6.QtUiTools",
        "PySide6.QtScxml",
        "PySide6.QtStateMachine",
        "PySide6.QtTextToSpeech",
        "PySide6.QtRemoteObjects",
        "PySide6.QtPdf",
        "PySide6.QtPdfWidgets",
        "PySide6.QtSvg",
        "PySide6.QtSvgWidgets",
        "PySide6.QtNetworkAuth",
        # --- Autres bibliothèques jamais importées par l'app ---
        "matplotlib",
        "tkinter",
        "_tkinter",
        "IPython",
        "jupyter",
        "notebook",
        "pytest",
        "nose",
        "fpdf",
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
    icon=os.path.join(project_dir, "ramanalyze", "assets", "Logo Ramanalyze", "logo-ramanalyze.ico"),
)
