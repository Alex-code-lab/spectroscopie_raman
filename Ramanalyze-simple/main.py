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
import traceback

from PySide6.QtGui import QIcon, QPalette, QColor
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QMessageBox

# Imports compatibles PyInstaller (onefile).
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    APP_DIR = sys._MEIPASS
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from store import SpectraStore
from presentation_tab import PresentationTab
from viewer_tab import ViewerTab
from titration_tab import TitrationTab
from peak_tracker_tab import PeakTrackerTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ramanalyze")
        self.resize(1380, 820)

        logo_path = os.path.join(APP_DIR, "assets", "logo.png")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        self.store = SpectraStore()

        tabs = QTabWidget(self)
        # onglets avec contour, l'onglet actif en bleu
        tabs.setStyleSheet(
            "QTabWidget::pane { border: 1px solid #555; top: -1px; }"
            "QTabBar::tab { padding: 6px 14px; margin-right: 3px; "
            "  background: #3a3a3a; color: #cfcfcf; "
            "  border: 1px solid #555; border-bottom: none; "
            "  border-top-left-radius: 5px; border-top-right-radius: 5px; }"
            "QTabBar::tab:selected { background: #0057b8; color: white; "
            "  border-color: #2f7bd6; }"
            "QTabBar::tab:hover:!selected { background: #474747; }"
        )
        tabs.addTab(PresentationTab(self), "Présentation")
        tabs.addTab(ViewerTab(self.store, self), "Visualiseur de spectres")
        tabs.addTab(PeakTrackerTab(self.store, self), "Titration par suivi de pic")
        tabs.addTab(TitrationTab(self.store, self), "Titration par rapport de pics")
        self.setCentralWidget(tabs)


def _install_excepthook():
    """Affiche toute erreur non gérée dans une popup (sinon elle reste invisible)."""
    def hook(exc_type, exc, tb):
        msg = "".join(traceback.format_exception(exc_type, exc, tb))
        print(msg, file=sys.stderr)
        try:
            box = QMessageBox()
            box.setIcon(QMessageBox.Critical)
            box.setWindowTitle("Erreur inattendue")
            box.setText("Une erreur s'est produite :")
            box.setDetailedText(msg)
            box.exec()
        except Exception:
            pass
    sys.excepthook = hook


def _dark_palette() -> QPalette:
    """Palette sombre fixe, identique sur Windows / macOS / Linux."""
    p = QPalette()
    p.setColor(QPalette.Window, QColor("#2b2b2b"))
    p.setColor(QPalette.WindowText, QColor("#e6e6e6"))
    p.setColor(QPalette.Base, QColor("#1e1e1e"))
    p.setColor(QPalette.AlternateBase, QColor("#2a2a2a"))
    p.setColor(QPalette.Text, QColor("#e6e6e6"))
    p.setColor(QPalette.Button, QColor("#3a3a3a"))
    p.setColor(QPalette.ButtonText, QColor("#e6e6e6"))
    p.setColor(QPalette.ToolTipBase, QColor("#2b2b2b"))
    p.setColor(QPalette.ToolTipText, QColor("#e6e6e6"))
    p.setColor(QPalette.PlaceholderText, QColor("#8a8a8a"))
    p.setColor(QPalette.Highlight, QColor("#0057b8"))
    p.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    p.setColor(QPalette.Link, QColor("#4ea1ff"))
    p.setColor(QPalette.Disabled, QPalette.Text, QColor("#777777"))
    p.setColor(QPalette.Disabled, QPalette.WindowText, QColor("#777777"))
    p.setColor(QPalette.Disabled, QPalette.ButtonText, QColor("#777777"))
    return p


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Ramanalyze-simple")
    # apparence identique quel que soit l'OS et le thème (toujours sombre)
    app.setStyle("Fusion")
    app.setPalette(_dark_palette())
    _install_excepthook()
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
