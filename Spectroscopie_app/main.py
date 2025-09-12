import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabWidget,
    QWidget,
    QVBoxLayout,
    QStatusBar,
)

# S'assurer que les imports de modules frères fonctionnent quel que soit l'endroit
# d'où on lance le script
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from file_picker import FilePickerWidget
from spectra_plot import SpectraTab
from analysis_tab import AnalysisTab
from metadata_picker import MetadataPickerWidget



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spectroscopie Raman")
        self.resize(1200, 800)

        # --- Status bar ---
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Prêt")

        # --- Onglets ---
        tabs = QTabWidget(self)
        self.setCentralWidget(tabs)

        # Onglet Fichiers
        self.file_tab = QWidget(self)
        file_layout = QVBoxLayout(self.file_tab)
        self.file_picker = FilePickerWidget(self.file_tab)
        file_layout.addWidget(self.file_picker)
        tabs.addTab(self.file_tab, "Fichiers")

        # Onglet Métadonnées
        self.metadata_tab = QWidget(self)
        metadata_layout = QVBoxLayout(self.metadata_tab)
        self.metadata_picker = MetadataPickerWidget(self.metadata_tab)
        metadata_layout.addWidget(self.metadata_picker)
        tabs.addTab(self.metadata_tab, "Métadonnées")

        # Onglet Spectres
        self.spectra_tab = SpectraTab(self.file_picker, self)
        tabs.addTab(self.spectra_tab, "Spectres")

        # Onglet Analyse
        self.analysis_tab = AnalysisTab(self)
        tabs.addTab(self.analysis_tab, "Analyse")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())