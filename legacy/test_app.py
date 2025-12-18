# NOTE: Legacy — prototype d'app, non utilisé par `main.py`.
import sys
import os
import re
import numpy as np
import pandas as pd

from pybaselines import Baseline

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QFileDialog, QLabel, QTabWidget, QMessageBox,
    QHBoxLayout, QListWidget, QListWidgetItem, QSplitter, QTreeView, QListView,
    QFileSystemModel, QAbstractItemView, QLineEdit, QToolButton, QStyle
)
from PySide6.QtCore import Qt, QSize

import plotly.express as px
from plotly.subplots import make_subplots
from PySide6.QtWebEngineWidgets import QWebEngineView


class FilePickerWidget(QWidget):
    """Integrated file explorer to pick .txt spectra files."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_files = []  # absolute paths

        # Main splitter: left = explorer, right = selection list
        splitter = QSplitter(self)

        # --- Left: file system icon view ---
        self.model = QFileSystemModel(self)
        self.model.setRootPath(os.path.expanduser("~"))
        self.model.setNameFilters(["*.txt"])  # only show .txt files
        self.model.setNameFilterDisables(False)  # hide non-matching files

        self.view = QListView(self)
        self.view.setModel(self.model)
        self.view.setRootIndex(self.model.index(os.path.expanduser("~")))
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setViewMode(QListView.IconMode)
        self.view.setResizeMode(QListView.Adjust)
        self.view.setUniformItemSizes(True)
        self.view.setIconSize(QSize(56, 56))
        self.view.setMovement(QListView.Static)
        self.view.doubleClicked.connect(self._open_or_add)

        # --- Right: selected files + controls ---
        right = QWidget(self)
        right_layout = QVBoxLayout(right)

        # Breadcrumb (clickable path)
        self.breadcrumb = QHBoxLayout()
        right_layout.addLayout(self.breadcrumb)

        # Quick path entry bar
        path_row = QHBoxLayout()
        self.path_edit = QLineEdit(os.path.expanduser("~"))
        self.path_edit.setPlaceholderText("Aller au dossier…")
        go_btn = QToolButton()
        go_btn.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
        go_btn.setToolTip("Aller au dossier indiqué")
        go_btn.clicked.connect(self._go_to_path)
        path_row.addWidget(self.path_edit)
        path_row.addWidget(go_btn)
        right_layout.addLayout(path_row)

        self.list = QListWidget(self)
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        right_layout.addWidget(QLabel("Fichiers sélectionnés (.txt)"))
        right_layout.addWidget(self.list)

        # Control buttons
        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("Ajouter la sélection")
        self.btn_remove = QPushButton("Retirer")
        self.btn_clear = QPushButton("Tout effacer")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        right_layout.addLayout(btn_row)

        self.btn_add.clicked.connect(self._add_from_view)
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_clear.clicked.connect(self._clear_all)

        splitter.addWidget(self.view)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # initialize path label at bottom and breadcrumb
        self.path_label = QLabel(os.path.expanduser("~"))
        main_layout.addWidget(self.path_label)
        self._rebuild_breadcrumb(os.path.expanduser("~"))

    # --- helpers ---
    def _go_to_path(self):
        path = self.path_edit.text().strip()
        if not path:
            return
        if os.path.isdir(path):
            self._set_root(path)
        else:
            QMessageBox.warning(self, "Chemin invalide", f"Le dossier n'existe pas:\n{path}")

    def _set_root(self, path: str):
        """Set the current root of the icon view and update UI."""
        if not os.path.isdir(path):
            return
        idx = self.model.index(path)
        if not idx.isValid():
            return
        self.view.setRootIndex(idx)
        self.path_edit.setText(path)
        self.path_label.setText(path)
        self._rebuild_breadcrumb(path)

    def _rebuild_breadcrumb(self, path: str):
        # clear previous crumbs
        while self.breadcrumb.count():
            item = self.breadcrumb.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        # build clickable segments
        parts = []
        head = path
        # Normalize and split
        head = os.path.abspath(path)
        root = os.path.splitdrive(head)[0] + os.sep if os.name == 'nt' else os.sep
        rel = os.path.relpath(head, root)
        segs = [] if rel == os.curdir else rel.split(os.sep)
        current = root
        # root button
        btn = QToolButton()
        btn.setText(root if root != os.sep else "/")
        btn.clicked.connect(lambda _=None, p=current: self._set_root(p))
        self.breadcrumb.addWidget(btn)
        for seg in segs:
            # separator label
            sep = QLabel(" / ")
            self.breadcrumb.addWidget(sep)
            current = os.path.join(current, seg)
            b = QToolButton()
            b.setText(seg)
            b.clicked.connect(lambda _=None, p=current: self._set_root(p))
            self.breadcrumb.addWidget(b)
        self.breadcrumb.addStretch(1)

    def _open_or_add(self, index):
        """Double-click behavior: enter folders, add .txt files."""
        if not index.isValid():
            return
        path = self.model.filePath(index)
        if os.path.isdir(path):
            self._set_root(path)
            return
        if path.lower().endswith('.txt'):
            if path not in self.selected_files:
                self.selected_files.append(path)
                self.list.addItem(QListWidgetItem(path))
                self._notify_count()

    def _add_from_view(self):
        """Add current selection from the icon view."""
        indexes = self.view.selectionModel().selectedIndexes()
        added = 0
        for idx in indexes:
            if not idx.isValid():
                continue
            path = self.model.filePath(idx)
            if os.path.isdir(path):
                continue
            if not path.lower().endswith('.txt'):
                continue
            if path not in self.selected_files:
                self.selected_files.append(path)
                self.list.addItem(QListWidgetItem(path))
                added += 1
        if added:
            self._notify_count()

    def _remove_selected(self):
        for item in self.list.selectedItems():
            path = item.text()
            if path in self.selected_files:
                self.selected_files.remove(path)
            self.list.takeItem(self.list.row(item))
        self._notify_count()

    def _clear_all(self):
        self.selected_files.clear()
        self.list.clear()
        self._notify_count()

    def _notify_count(self):
        count = len(self.selected_files)
        main = self.window()
        if isinstance(main, QMainWindow):
            main.statusBar().showMessage(f"{count} fichier(s) sélectionné(s)")


class RamanApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spectroscopie Raman - Analyse")
        self.setGeometry(100, 100, 1200, 800)

        self.metadata_path = None
        self.combined_df = None

        # Layout principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Barre d'actions
        actions = QHBoxLayout()
        self.btn_load_metadata = QPushButton("Charger les métadonnées")
        self.btn_load_metadata.clicked.connect(self.load_metadata)
        self.btn_run = QPushButton("Analyser (plus tard)")
        self.btn_run.clicked.connect(self.run_analysis)
        actions.addWidget(self.btn_load_metadata)
        actions.addStretch(1)
        actions.addWidget(self.btn_run)
        layout.addLayout(actions)

        # Onglets
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Tab 1: Fichiers (file explorer)
        self.file_tab = QWidget()
        self.file_picker = FilePickerWidget(self.file_tab)
        file_tab_layout = QVBoxLayout(self.file_tab)
        file_tab_layout.addWidget(QLabel("Choisissez les fichiers .txt à analyser"))
        file_tab_layout.addWidget(self.file_picker)
        self.tabs.addTab(self.file_tab, "Fichiers")

        # Tab 2: Spectres (placeholder for plots)
        self.tab_spectra = QWidget()
        self.tabs.addTab(self.tab_spectra, "Spectres")
        self.plot_view = QWebEngineView()
        vbox = QVBoxLayout(self.tab_spectra)
        vbox.addWidget(QLabel("Spectres Raman corrigés (aperçu)"))
        # Button to plot selected spectra with baseline correction
        self.btn_plot_selected = QPushButton("Tracer les spectres sélectionnés (avec baseline)")
        self.btn_plot_selected.clicked.connect(self.plot_selected_with_baseline)
        vbox.addWidget(self.btn_plot_selected)
        vbox.addWidget(self.plot_view)

        # Tab 3: Ratios (placeholder)
        self.tab_ratios = QWidget()
        self.tabs.addTab(self.tab_ratios, "Ratios")

        # Status bar
        self.statusBar().showMessage("Prêt")

        # Simple stylesheet for a cleaner look
        self.setStyleSheet(
            """
            QMainWindow { background: #f7f7f9; }
            QLabel { font-size: 13px; }
            QPushButton { padding: 6px 10px; }
            QTreeView { font-size: 12px; }
            QListWidget { font-family: Menlo, Consolas, monospace; font-size: 12px; }
            """
        )

    def load_metadata(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Choisir le fichier Excel metadata", filter="Excel Files (*.xlsx)")
        if file_path:
            self.metadata_path = file_path
            QMessageBox.information(self, "Info", f"Metadata chargé : {file_path}")

    def run_analysis(self):
        QMessageBox.information(self, "Analyse", "L'analyse sera implémentée ultérieurement.")
    def plot_selected_with_baseline(self):
        """Read currently selected .txt files, compute baseline (modpoly, deg=5),
        and display raw/baseline/corrected spectra in the Spectres tab."""
        selected = self.file_picker.selected_files if hasattr(self, 'file_picker') else []
        if not selected:
            QMessageBox.information(self, "Aucun fichier", "Sélectionnez d'abord des fichiers .txt dans l'onglet Fichiers.")
            return

        all_traces = []
        try:
            for path in selected:
                # read file with dynamic header detection
                with open(path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                try:
                    header_idx = next(i for i, line in enumerate(lines) if line.strip().startswith('Pixel;'))
                except StopIteration:
                    # if no explicit header, default to 0
                    header_idx = 0
                df = pd.read_csv(path, skiprows=header_idx, sep=';', decimal=',', encoding='utf-8')
                if df.columns[-1].startswith('Unnamed'):
                    df = df.iloc[:, :-1]

                # pick columns if present
                if not set(["Raman Shift", "Dark Subtracted #1"]).issubset(df.columns):
                    # skip files that do not have expected structure
                    continue
                temp = df[["Raman Shift", "Dark Subtracted #1"]].dropna()
                if temp.empty:
                    continue

                # ensure numeric
                temp["Raman Shift"] = pd.to_numeric(temp["Raman Shift"], errors='coerce')
                temp["Dark Subtracted #1"] = pd.to_numeric(temp["Dark Subtracted #1"], errors='coerce')
                temp = temp.dropna()
                if temp.empty:
                    continue

                x = temp["Raman Shift"].values
                y = temp["Dark Subtracted #1"].values

                # baseline (modpoly deg=5)
                try:
                    bl_fitter = Baseline(x)
                    baseline, _ = bl_fitter.modpoly(y, poly_order=5)
                    ycorr = y - baseline
                except Exception:
                    # fallback: no baseline
                    baseline = np.zeros_like(y)
                    ycorr = y

                name = os.path.basename(path)
                # build traces (raw, baseline, corrected)
                all_traces.append({"x": x, "y": y, "name": f"{name} – brut", "mode": "lines", "line": {"width": 1, "dash": "solid"}})
                all_traces.append({"x": x, "y": baseline, "name": f"{name} – baseline", "mode": "lines", "line": {"width": 1, "dash": "dash"}})
                all_traces.append({"x": x, "y": ycorr, "name": f"{name} – corrigé", "mode": "lines", "line": {"width": 2}})
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Problème lors du tracé:\n{e}")
            return

        if not all_traces:
            QMessageBox.information(self, "Rien à tracer", "Aucun spectre valide n'a été trouvé dans la sélection.")
            return

        # Assemble Plotly figure
        import plotly.graph_objects as go
        fig = go.Figure()
        for tr in all_traces:
            fig.add_trace(go.Scatter(x=tr["x"], y=tr["y"], name=tr["name"], mode=tr["mode"], line=tr["line"]))

        fig.update_layout(
            title="Spectres sélectionnés (brut / baseline / corrigé)",
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité (a.u.)",
            template="plotly_white",
            width=1100,
            height=700,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))
        # For now, just validate selected files and show a tiny preview message
        selected = self.file_picker.selected_files if hasattr(self, 'file_picker') else []
        if not selected:
            QMessageBox.information(self, "Aucun fichier", "Sélectionnez d'abord des fichiers .txt dans l'onglet Fichiers.")
            return

        # Build a simple combined DataFrame preview (first 50 rows per file)
        try:
            previews = []
            for path in selected[:10]:  # cap preview to 10 files
                with open(path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                header_idx = next(i for i, line in enumerate(lines) if line.strip().startswith('Pixel;'))
                df = pd.read_csv(path, skiprows=header_idx, sep=';', decimal=',', encoding='utf-8')
                if df.columns[-1].startswith('Unnamed'):
                    df = df.iloc[:, :-1]
                # keep only minimal columns if present
                print("Colonnes pour", path, ":", df.columns.tolist()[:10])
                keep = [c for c in ["Raman Shift", "Dark Subtracted #1"] if c in df.columns]
                if keep:
                    df = df[keep].head(50)
                df["file"] = os.path.basename(path)
                previews.append(df)
            preview_df = pd.concat(previews, ignore_index=True) if previews else pd.DataFrame()
        except Exception as e:
            QMessageBox.warning(self, "Erreur de lecture", f"Impossible de prévisualiser les fichiers:\n{e}")
            return

        # Quick interactive line preview if possible
        if not preview_df.empty and set(["Raman Shift", "Dark Subtracted #1"]).issubset(preview_df.columns):
            fig = px.line(preview_df, x="Raman Shift", y="Dark Subtracted #1", color="file", title="Aperçu rapide (brut)")
            fig.update_layout(width=1000, height=600)
            self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))
        else:
            QMessageBox.information(self, "OK", f"{len(selected)} fichier(s) prêt(s). Le traitement sera ajouté ultérieurement.")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Charger la feuille de style externe
    qss_path = os.path.join(os.path.dirname(__file__), "./styles/styles.qss")
    if os.path.exists(qss_path):
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    window = RamanApp()
    window.show()
    ret = app.exec()
    sys.exit(ret)
