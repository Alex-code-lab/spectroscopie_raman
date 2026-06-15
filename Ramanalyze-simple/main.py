"""Ramanalyze-simple

Visualiseur minimal de spectres Raman + titration.
- Onglet « Visualiseur » : on va chercher des .txt, ils s'affichent en grand,
  le nom de chaque courbe est le nom du fichier.
- Onglet « Titration » : on saisit une concentration par échantillon, on mesure
  l'intensité d'un pic (ou un ratio de deux pics), et on trace la courbe de
  titration (intensité vs concentration), avec ajustement sigmoïde optionnel.

Lancer :  python main.py
"""

import os
import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

# Imports compatibles PyInstaller (onefile).
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    APP_DIR = sys._MEIPASS
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from store import SpectraStore
from viewer_tab import ViewerTab
from titration_tab import TitrationTab
from peak_tracker_tab import PeakTrackerTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ramanalyze · Visualiseur de spectres")
        self.resize(1380, 820)

        icon_path = os.path.join(APP_DIR, "assets", "logo.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.store = SpectraStore()

        tabs = QTabWidget(self)
        tabs.addTab(ViewerTab(self.store, self), "Visualiseur")
        tabs.addTab(TitrationTab(self.store, self), "Titration")
        tabs.addTab(PeakTrackerTab(self.store, self), "Suivi de pic")
        self.setCentralWidget(tabs)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Ramanalyze-simple")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
