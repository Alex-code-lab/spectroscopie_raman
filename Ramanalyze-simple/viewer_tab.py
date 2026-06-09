"""Onglet Visualiseur : on ouvre des .txt, ils s'affichent en grand,
légende = nom du fichier.

Les fichiers ouverts ici restent locaux au Visualiseur : ils ne sont
envoyés à l'onglet Titration que via le bouton « Envoyer vers Titration ».
"""

import os

from PySide6.QtCore import Qt, QSettings
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QSplitter,
    QScrollArea,
    QFileDialog,
    QMessageBox,
)
import plotly.graph_objects as go

from spectrum_loader import load_spectrum
from plot_view import PlotlyView

_SETTINGS_KEY = "simple/last_dir"


class ViewerTab(QWidget):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store  # store partagé : alimenté seulement via le bouton

        # Données locales au Visualiseur (indépendantes de la Titration).
        self._spectra: dict[str, tuple] = {}   # path -> (x, y)
        self._checked: dict[str, bool] = {}     # path -> affiché ?

        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(QLabel("<b>Spectres</b>", self))

        self.btn_open = QPushButton("📂  Ouvrir des fichiers…", self)
        self.btn_open.setStyleSheet(
            "background-color: #0057b8; color: white; font-weight: 700; padding: 8px;"
        )
        self.btn_open.clicked.connect(self.open_files)
        left_layout.addWidget(self.btn_open)

        hint = QLabel("Cochez / décochez un spectre pour l'afficher ou le masquer.", self)
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888;")
        left_layout.addWidget(hint)

        self.file_list = QListWidget(self)
        self.file_list.setMinimumHeight(220)
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.itemChanged.connect(self._on_item_changed)
        left_layout.addWidget(self.file_list, 1)

        self.btn_send = QPushButton("➡  Envoyer vers Titration", self)
        self.btn_send.setToolTip(
            "Envoie les fichiers sélectionnés (ou tous si aucun n'est sélectionné) "
            "vers l'onglet Titration."
        )
        self.btn_send.setStyleSheet("font-weight: 600; padding: 6px;")
        self.btn_send.clicked.connect(self.send_to_titration)
        left_layout.addWidget(self.btn_send)

        btn_row = QHBoxLayout()
        self.btn_remove = QPushButton("Retirer", self)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("Tout vider", self)
        self.btn_clear.clicked.connect(self.clear_all)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        left_layout.addLayout(btn_row)

        self.status = QLabel("Aucun spectre", self)
        self.status.setStyleSheet("color: #888;")
        self.status.setWordWrap(True)
        left_layout.addWidget(self.status)

        left_scroll = QScrollArea(self)
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        left_scroll.setMinimumWidth(340)

        self.plot_view = PlotlyView(self)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(left_scroll)
        splitter.addWidget(self.plot_view)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([380, 900])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)

        self.refresh_plot()

    # ------------------------------------------------------------------
    def _last_dir(self) -> str:
        saved = QSettings("Ramanalyze", "Simple").value(_SETTINGS_KEY, "")
        if saved and os.path.isdir(saved):
            return saved
        return os.path.expanduser("~")

    def open_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Choisir des spectres (.txt)",
            self._last_dir(),
            "Spectres Raman (*.txt);;Tous les fichiers (*)",
        )
        if not paths:
            return
        QSettings("Ramanalyze", "Simple").setValue(_SETTINGS_KEY, os.path.dirname(paths[0]))

        failed = []
        self.file_list.blockSignals(True)
        for path in paths:
            if path in self._spectra:
                continue
            data = load_spectrum(path)
            if data is None:
                failed.append(os.path.basename(path))
                continue
            self._spectra[path] = data
            self._checked[path] = True
            self._add_item(path)
        self.file_list.blockSignals(False)

        self.refresh_plot()
        if failed:
            self.status.setText(
                f"{len(failed)} fichier(s) non lisible(s) : "
                + ", ".join(failed[:3]) + ("…" if len(failed) > 3 else "")
            )

    def _add_item(self, path: str):
        item = QListWidgetItem(os.path.basename(path))
        item.setToolTip(path)
        item.setData(Qt.UserRole, path)
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked if self._checked.get(path, True) else Qt.Unchecked)
        self.file_list.addItem(item)

    def _on_item_changed(self, item: QListWidgetItem):
        path = item.data(Qt.UserRole)
        self._checked[path] = item.checkState() == Qt.Checked
        self.refresh_plot()

    def remove_selected(self):
        for item in self.file_list.selectedItems():
            path = item.data(Qt.UserRole)
            self._spectra.pop(path, None)
            self._checked.pop(path, None)
            self.file_list.takeItem(self.file_list.row(item))
        self.refresh_plot()

    def clear_all(self):
        self._spectra.clear()
        self._checked.clear()
        self.file_list.clear()
        self.refresh_plot()

    # ------------------------------------------------------------------
    def send_to_titration(self):
        """Envoie les fichiers sélectionnés (ou tous si aucun) vers la Titration."""
        selected = [it.data(Qt.UserRole) for it in self.file_list.selectedItems()]
        paths = selected if selected else list(self._spectra.keys())
        if not paths:
            QMessageBox.information(self, "Rien à envoyer", "Ouvrez d'abord des spectres.")
            return
        n_new = 0
        for path in paths:
            if path not in self.store.spectra:
                n_new += 1
            self.store.add(path, self._spectra[path])
        scope = "sélectionné(s)" if selected else "chargé(s)"
        self.status.setText(
            f"{len(paths)} spectre(s) {scope} envoyé(s) vers Titration "
            f"({n_new} nouveau(x))."
        )

    # ------------------------------------------------------------------
    def _checked_paths(self) -> list[str]:
        return [p for p in self._spectra if self._checked.get(p, True)]

    def refresh_plot(self):
        paths = self._checked_paths()
        fig = go.Figure()
        for path in paths:
            x, y = self._spectra[path]
            fig.add_trace(go.Scatter(x=x, y=y, mode="lines", name=os.path.basename(path)))

        fig.update_layout(
            title="Spectres Raman" if paths else "Ouvrez des fichiers .txt pour les visualiser",
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité (a.u.)",
            template="plotly_white",
            margin=dict(l=70, r=30, t=60, b=60),
            legend=dict(title="Fichiers"),
            font=dict(size=14),
        )
        self.plot_view.show_figure(fig)

        n_total = len(self._spectra)
        if n_total == 0:
            self.status.setText("Aucun spectre")
        else:
            self.status.setText(f"{len(paths)} affiché(s) / {n_total} chargé(s)")
