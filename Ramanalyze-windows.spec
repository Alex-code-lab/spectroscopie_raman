# -*- mode: python ; coding: utf-8 -*-
"""Spec PyInstaller Windows pour Ramanalyze.

Construit un executable one-file depuis l'application actuelle dans
``Ramanalyze/``. Les ressources sont embarquees sous ``assets/`` pour matcher
les chemins utilises par ``main.py`` en mode PyInstaller (``sys._MEIPASS``).
"""

from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata


APP_NAME = "Ramanalyze"
block_cipher = None


def _project_dir() -> Path:
    if "__file__" in globals():
        return Path(__file__).resolve().parent
    return Path(globals().get("SPECPATH", globals().get("specpath", "."))).resolve()


def _require(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} introuvable: {path}")
    return path


def _runtime_module(name: str) -> bool:
    ignored_parts = (".tests", ".test_", ".conftest")
    ignored_names = (
        "plotly.conftest",
        "plotly.io._sg_scraper",
        "kaleido._mocker",
    )
    return (
        not any(part in name for part in ignored_parts)
        and not name.endswith(".tests")
        and name not in ignored_names
    )


def _collect_submodules(package: str) -> list[str]:
    try:
        return collect_submodules(package, filter=_runtime_module)
    except Exception as exc:  # noqa: BLE001 - le build doit rester lisible.
        print(f"[spec] Sous-modules non collectes pour {package}: {exc}")
        return []


def _collect_datas(package: str):
    try:
        return collect_data_files(package)
    except Exception as exc:  # noqa: BLE001
        print(f"[spec] Donnees non collectees pour {package}: {exc}")
        return []


def _collect_plotly_runtime() -> list[str]:
    modules = [
        "plotly.graph_objects",
        "plotly.graph_objs",
        "plotly.graph_objs._figure",
        "plotly.graph_objs._layout",
        "plotly.graph_objs._scatter",
        "plotly.offline",
        "plotly.offline.offline",
        "plotly.shapeannotation",
        "plotly.validator_cache",
    ]
    modules += _collect_submodules("plotly.graph_objs.layout")
    modules += _collect_submodules("plotly.graph_objs.scatter")
    modules += _collect_submodules("plotly.io")
    return modules


def _copy_metadata(package: str):
    try:
        return copy_metadata(package)
    except Exception:
        return []


project_dir = _project_dir()
app_dir = _require(project_dir / "Ramanalyze", "Dossier application")
entry_point = _require(app_dir / "main.py", "Point d'entree")
assets_dir = _require(app_dir / "assets", "Dossier assets")
icon_file = _require(assets_dir / "logo.ico", "Icone Windows")

# Les imports locaux sont ecrits sans package (from store import ...).
sys.setrecursionlimit(max(sys.getrecursionlimit(), 5000))
pathex = [str(app_dir), str(project_dir)]

datas = [
    (str(assets_dir), "assets"),
]
binaries = []

# QtWebEngine est charge par PlotlyView. Les hooks PyInstaller/PySide6 gerent
# les DLL et ressources Qt; ces imports cachés rendent l'intention explicite.
hiddenimports = [
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "PySide6.QtWebChannel",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
]

# Imports dynamiques dans le code:
# - scipy.signal.find_peaks dans peak_tracker_tab.py
# - scipy.optimize.curve_fit dans titrant_utils.py
# - pybaselines.Baseline dans les corrections de ligne de base
for package in (
    "scipy.optimize",
    "scipy.signal",
    "pybaselines",
    "kaleido",
):
    hiddenimports += _collect_submodules(package)

# Plotly charge Figure/Scatter et leurs sous-proprietes par imports paresseux.
hiddenimports += _collect_plotly_runtime()

# Plotly/Kaleido peuvent consulter leurs metadata au runtime pour les exports.
datas += _copy_metadata("plotly")
datas += _copy_metadata("kaleido")
datas += _collect_datas("kaleido")

excludes = [
    # Bindings Qt concurrents.
    "PyQt5",
    "PyQt6",
    "PySide2",
    "pyqt5",
    "tkinter",
    "_tkinter",
    # Ecosysteme notebook/tests non utilise.
    "IPython",
    "jupyter",
    "notebook",
    "pytest",
    "nose",
    # Bibliotheques lourdes non utilisees par Ramanalyze-simple.
    "matplotlib",
    "plotly.express",
    "plotly.figure_factory",
    "patsy",
    "statsmodels",
    "sklearn",
    "openpyxl",
    "xlrd",
    "xlsxwriter",
    # Modules Qt volumineux non importes. Ne pas exclure QtNetwork,
    # QtPositioning, QtQml, QtQuick, QtOpenGL ni QtPrintSupport: WebEngine peut
    # en dependre selon la version PySide6/PyInstaller.
    "PySide6.Qt3DAnimation",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DExtras",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DLogic",
    "PySide6.Qt3DRender",
    "PySide6.QtBluetooth",
    "PySide6.QtCharts",
    "PySide6.QtDataVisualization",
    "PySide6.QtDesigner",
    "PySide6.QtHelp",
    "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets",
    "PySide6.QtNfc",
    "PySide6.QtPdf",
    "PySide6.QtPdfWidgets",
    "PySide6.QtQuick3D",
    "PySide6.QtRemoteObjects",
    "PySide6.QtScxml",
    "PySide6.QtSensors",
    "PySide6.QtSerialBus",
    "PySide6.QtSerialPort",
    "PySide6.QtSpatialAudio",
    "PySide6.QtSql",
    "PySide6.QtStateMachine",
    "PySide6.QtSvg",
    "PySide6.QtSvgWidgets",
    "PySide6.QtTest",
    "PySide6.QtTextToSpeech",
    "PySide6.QtUiTools",
]

a = Analysis(
    [str(entry_point)],
    pathex=pathex,
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_file),
)
