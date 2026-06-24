"""Onglet « Suivi de pic ».

But : dans une fenêtre spectrale (par défaut **1200–1500 cm⁻¹**), repérer **tous**
les pics présents, puis suivre **chacun individuellement** d'un spectre à l'autre.

Démarche :
1. On détecte les pics (scipy.find_peaks) dans la fenêtre, pour chaque spectre.
2. On **apparie** ces pics d'un spectre à l'autre en regroupant ceux dont la
   position est proche (tolérance réglable) → chaque groupe = un « pic suivi ».
3. Pour chaque pic suivi, on mesure son intensité dans chaque spectre et on trace
   **une courbe par (série × pic)**, en abscisse la quantité de matière de titrant.

Volume de titrant, série et réglages titrant sont partagés avec l'onglet Titration
(via le store) et peuvent être enregistrés / rechargés.
"""

import csv
import os

import numpy as np
from PySide6.QtCore import Qt, QSettings
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QSplitter,
    QScrollArea,
    QFileDialog,
    QDoubleSpinBox,
    QSpinBox,
    QCheckBox,
    QComboBox,
    QMessageBox,
    QGroupBox,
    QFormLayout,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
)
import plotly.graph_objects as go

from spectrum_loader import load_spectrum
from plot_view import PlotlyView
import titrant_utils as tu
import plot_style as ps

_SETTINGS_KEY = "simple/last_dir"

# Chaque série a sa propre FAMILLE de couleurs (teinte distincte) ; à l'intérieur
# d'une série, chaque pic prend une nuance différente de cette famille.
# Les données sont tracées en trait PLEIN ; seules les sigmoïdes sont en tireté.
_FAMILIES = [
    ["#08519c", "#3182bd", "#6baed6", "#2171b5", "#4292c6", "#9ecae1", "#c6dbef"],  # bleus
    ["#a63603", "#e6550d", "#fd8d3c", "#d94801", "#f16913", "#fdae6b", "#fdd0a2"],  # oranges
    ["#006d2c", "#31a354", "#74c476", "#238b45", "#41ab5d", "#a1d99b", "#c7e9c0"],  # verts
    ["#54278f", "#756bb1", "#9e9ac8", "#6a51a3", "#807dba", "#bcbddc", "#dadaeb"],  # violets
    ["#99000d", "#cb181d", "#ef3b2c", "#a50f15", "#fb6a4a", "#fc9272", "#fcbba1"],  # rouges
    ["#b35806", "#8c510a", "#bf812d", "#dfc27d", "#7f3b08", "#e08214", "#fdb863"],  # bruns/ocres
]
_NO_SERIES = "(sans série)"


def _baseline_corrected(x, y, poly_order=5):
    """Soustrait une ligne de base modpoly (pybaselines). Renvoie y brut en cas d'échec."""
    try:
        from pybaselines import Baseline

        baseline, _ = Baseline(x).modpoly(y, poly_order=poly_order)
        return y - baseline
    except Exception as exc:  # noqa: BLE001
        print(f"[suivi-pic] Baseline impossible : {exc}")
        return y


def _detect_peaks(x, y, wmin, wmax, prominence_frac):
    """Positions des pics détectés dans [wmin, wmax] (liste de cm⁻¹)."""
    mask = (x >= wmin) & (x <= wmax)
    if not mask.any():
        return []
    xb, yb = x[mask], y[mask]
    try:
        from scipy.signal import find_peaks

        span = float(np.ptp(yb)) or 1.0
        idx, _ = find_peaks(yb, prominence=prominence_frac * span)
        return [float(xb[i]) for i in idx]
    except Exception:  # noqa: BLE001
        return [float(xb[int(np.argmax(yb))])]


def _cluster(detections, tol):
    """Regroupe des pics proches en « pics suivis ».

    detections : liste de (position_cm, indice_spectre).
    Renvoie une liste de (centre_cm, support) triée par centre.
    """
    if not detections:
        return []
    pts = sorted(detections, key=lambda d: d[0])
    clusters = []
    cur = [pts[0]]
    for pos, si in pts[1:]:
        if pos - cur[-1][0] > tol:
            clusters.append(cur)
            cur = []
        cur.append((pos, si))
    clusters.append(cur)

    out = []
    for cl in clusters:
        center = float(np.mean([p for p, _ in cl]))
        support = len({si for _, si in cl})
        out.append((center, support))
    return out


def _measure_at(x, y, center, tol):
    """Intensité maximale dans [center-tol, center+tol], ou NaN."""
    mask = (x >= center - tol) & (x <= center + tol)
    if not mask.any():
        return np.nan
    return float(np.max(y[mask]))


class PeakTrackerTab(QWidget):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store
        self._corr = {}            # path -> (x, y_corrigé) calculé à la détection
        self._centers = []         # liste de (centre_cm, support) détectés
        self._last_fig = None
        self._last_file_base = "evolution_pics"
        self._populating = False
        self._emitting_meta = False

        # ---------------- Panneau gauche ----------------
        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(QLabel("<b>Suivi individuel des pics</b>", self))
        hint = QLabel(
            "On repère tous les pics de la fenêtre, on les apparie d'un spectre à "
            "l'autre, puis on suit chacun séparément. Volumes et séries sont partagés "
            "avec l'onglet Titration.", self
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888;")
        left_layout.addWidget(hint)

        self.btn_open = QPushButton("📂  Ajouter des fichiers…", self)
        self.btn_open.clicked.connect(self.open_files)
        left_layout.addWidget(self.btn_open)

        # --- Fichiers chargés + volume de titrant + série par échantillon ---
        files_box = QGroupBox("Fichiers chargés · volume de titrant · série", self)
        files_layout = QVBoxLayout(files_box)
        files_hint = QLabel(
            "Double-cliquez les colonnes <i>Volume</i> et <i>Série</i> pour les saisir. "
            "Chaque série est tracée indépendamment. Sélectionnez des lignes pour les "
            "retirer ou leur affecter une série.", self
        )
        files_hint.setWordWrap(True)
        files_hint.setStyleSheet("color: #888;")
        files_layout.addWidget(files_hint)

        self.file_table = QTableWidget(0, 3, self)
        self.file_table.setMinimumHeight(150)
        self.file_table.setHorizontalHeaderLabels(["Fichier", "Volume titrant", "Série"])
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_table.itemChanged.connect(self._on_table_edited)
        files_layout.addWidget(self.file_table)

        # affectation groupée d'une série à la sélection
        series_row = QHBoxLayout()
        self.edit_series = QLineEdit(self)
        self.edit_series.setPlaceholderText("Nom de série (ex. Série 1)")
        series_row.addWidget(self.edit_series, 1)
        self.btn_assign = QPushButton("Affecter à la sélection", self)
        self.btn_assign.clicked.connect(self.assign_series)
        series_row.addWidget(self.btn_assign)
        files_layout.addLayout(series_row)

        files_btns = QHBoxLayout()
        self.btn_remove = QPushButton("Retirer la sélection", self)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("Tout vider", self)
        self.btn_clear.clicked.connect(self.clear_all)
        files_btns.addWidget(self.btn_remove)
        files_btns.addWidget(self.btn_clear)
        files_layout.addLayout(files_btns)

        # enregistrer / charger le tableau (partagé avec Titration)
        session_btns = QHBoxLayout()
        self.btn_save_session = QPushButton("💾 Enregistrer le tableau…", self)
        self.btn_save_session.clicked.connect(self.save_session)
        self.btn_load_session = QPushButton("📥 Charger un tableau…", self)
        self.btn_load_session.clicked.connect(self.load_session)
        session_btns.addWidget(self.btn_save_session)
        session_btns.addWidget(self.btn_load_session)
        files_layout.addLayout(session_btns)
        left_layout.addWidget(files_box)

        # --- Titrant : concentration de la solution + unités ---
        titrant_box = QGroupBox("Titrant (pour l'axe « quantité de matière »)", self)
        titrant_form = QFormLayout(titrant_box)

        conc_row = QHBoxLayout()
        self.spin_conc = QDoubleSpinBox(self)
        self.spin_conc.setRange(0.0, 1e9)
        self.spin_conc.setDecimals(4)
        self.spin_conc.setValue(1.0)
        conc_row.addWidget(self.spin_conc, 1)
        self.cmb_conc_unit = QComboBox(self)
        self.cmb_conc_unit.addItems(list(tu.CONC_FACTORS.keys()))  # nM, µM, mM, M
        self.cmb_conc_unit.setCurrentText("µM")
        conc_row.addWidget(self.cmb_conc_unit)
        titrant_form.addRow("Concentration :", conc_row)

        self.cmb_vol_unit = QComboBox(self)
        self.cmb_vol_unit.addItems(list(tu.VOL_FACTORS.keys()))    # µL, mL, L
        self.cmb_vol_unit.setCurrentText("µL")
        titrant_form.addRow("Unité des volumes :", self.cmb_vol_unit)

        self.cmb_xaxis = QComboBox(self)
        self.cmb_xaxis.addItems(["Quantité de matière de titrant", "Nom de spectre"])
        self.cmb_xaxis.setToolTip(
            "Abscisse du graphe. Par défaut : quantité de matière = concentration × "
            "volume de titrant (axe numérique → les séries de même quantité se superposent)."
        )
        titrant_form.addRow("Axe X :", self.cmb_xaxis)
        left_layout.addWidget(titrant_box)

        # synchro des réglages titrant -> store
        self.spin_conc.valueChanged.connect(self._on_titrant_changed)
        self.cmb_conc_unit.currentTextChanged.connect(self._on_titrant_changed)
        self.cmb_vol_unit.currentTextChanged.connect(self._on_titrant_changed)

        # --- Fenêtre + détection ---
        params = QGroupBox("Fenêtre & détection", self)
        form = QFormLayout(params)

        self.spin_min = QDoubleSpinBox(self)
        self.spin_min.setRange(0.0, 5000.0)
        self.spin_min.setDecimals(0)
        self.spin_min.setSingleStep(10.0)
        self.spin_min.setValue(1200.0)
        self.spin_min.setSuffix(" cm⁻¹")
        form.addRow("Borne min :", self.spin_min)

        self.spin_max = QDoubleSpinBox(self)
        self.spin_max.setRange(0.0, 5000.0)
        self.spin_max.setDecimals(0)
        self.spin_max.setSingleStep(10.0)
        self.spin_max.setValue(1500.0)
        self.spin_max.setSuffix(" cm⁻¹")
        form.addRow("Borne max :", self.spin_max)

        self.spin_prom = QDoubleSpinBox(self)
        self.spin_prom.setRange(0.1, 100.0)
        self.spin_prom.setDecimals(1)
        self.spin_prom.setSingleStep(1.0)
        self.spin_prom.setValue(5.0)
        self.spin_prom.setSuffix(" %")
        self.spin_prom.setToolTip("Proéminence minimale d'un pic, en % de l'amplitude de la fenêtre.")
        form.addRow("Proéminence :", self.spin_prom)

        self.spin_tol = QDoubleSpinBox(self)
        self.spin_tol.setRange(1.0, 200.0)
        self.spin_tol.setDecimals(0)
        self.spin_tol.setSingleStep(1.0)
        self.spin_tol.setValue(10.0)
        self.spin_tol.setSuffix(" cm⁻¹")
        self.spin_tol.setToolTip("Tolérance d'appariement des pics entre spectres et fenêtre de mesure.")
        form.addRow("Tolérance :", self.spin_tol)

        self.spin_support = QSpinBox(self)
        self.spin_support.setRange(1, 999)
        self.spin_support.setValue(2)
        self.spin_support.setToolTip("Un pic n'est gardé que s'il apparaît dans au moins ce nombre de spectres.")
        form.addRow("Présence min. :", self.spin_support)

        self.chk_baseline = QCheckBox("Corriger la ligne de base", self)
        self.chk_baseline.setChecked(True)
        form.addRow(self.chk_baseline)

        left_layout.addWidget(params)

        self.btn_detect = QPushButton("1) Détecter les pics", self)
        self.btn_detect.clicked.connect(self.detect_peaks)
        left_layout.addWidget(self.btn_detect)

        # --- Liste des pics détectés (cochables) ---
        left_layout.addWidget(QLabel("<b>Pics détectés</b> (cochez ceux à suivre)", self))
        self.list_peaks = QListWidget(self)
        self.list_peaks.setMinimumHeight(160)
        self.list_peaks.itemChanged.connect(self._on_peak_toggled)
        left_layout.addWidget(self.list_peaks, 1)

        sel_row = QHBoxLayout()
        self.btn_all = QPushButton("Tout cocher", self)
        self.btn_all.clicked.connect(lambda: self._set_all_checked(True))
        self.btn_none = QPushButton("Tout décocher", self)
        self.btn_none.clicked.connect(lambda: self._set_all_checked(False))
        sel_row.addWidget(self.btn_all)
        sel_row.addWidget(self.btn_none)
        left_layout.addLayout(sel_row)

        # --- Titre + tracé ---
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Titre :", self))
        self.edit_title = QLineEdit(self)
        self.edit_title.setPlaceholderText("Titre automatique (laisser vide)")
        title_row.addWidget(self.edit_title, 1)
        left_layout.addLayout(title_row)

        self.chk_sigmoid = QCheckBox("Fit en tirets + x_eq (fin de palier)", self)
        self.chk_sigmoid.setToolTip(
            "Trace le fit en tirets (paliers + descente/montée) et marque x_eq = "
            "abscisse où le signal a fini de bouger (nouveau palier atteint). "
            "Nécessite l'axe « Quantité de matière » et au moins 4 points."
        )
        self.chk_sigmoid.toggled.connect(self._on_sigmoid_toggled)
        left_layout.addWidget(self.chk_sigmoid)

        fit_row = QHBoxLayout()
        fit_row.addWidget(QLabel("Courbe à ajuster :", self))
        self.cmb_fit_curve = QComboBox(self)
        self.cmb_fit_curve.setToolTip(
            "Courbe (série × pic) sur laquelle marquer x_eq (fin de palier). "
            "« Toutes les courbes » marque x_eq pour chacune (sans trait vertical)."
        )
        self.cmb_fit_curve.setEnabled(False)
        self.cmb_fit_curve.currentIndexChanged.connect(self._on_fit_curve_changed)
        fit_row.addWidget(self.cmb_fit_curve, 1)
        left_layout.addLayout(fit_row)

        plateau_row = QHBoxLayout()
        plateau_row.addWidget(QLabel("Palier atteint à :", self))
        self.spin_plateau = QDoubleSpinBox(self)
        self.spin_plateau.setRange(50.0, 99.9)
        self.spin_plateau.setDecimals(1)
        self.spin_plateau.setSingleStep(1.0)
        self.spin_plateau.setValue(95.0)
        self.spin_plateau.setSuffix(" %")
        self.spin_plateau.setToolTip(
            "Seuil définissant « le signal a fini de bouger » : abscisse où la "
            "sigmoïde a parcouru ce pourcentage de son changement total."
        )
        self.spin_plateau.valueChanged.connect(self._on_fit_curve_changed)
        plateau_row.addWidget(self.spin_plateau)
        plateau_row.addStretch(1)
        left_layout.addLayout(plateau_row)

        self.btn_plot = QPushButton("2) Tracer l'évolution", self)
        self.btn_plot.setStyleSheet(
            "background-color: #0057b8; color: white; font-weight: 700; padding: 8px;"
        )
        self.btn_plot.clicked.connect(self.plot_evolution)
        left_layout.addWidget(self.btn_plot)

        btn_row = QHBoxLayout()
        self.btn_export_csv = QPushButton("Exporter (CSV)…", self)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_csv.setEnabled(False)
        self.btn_export_graph = QPushButton("Exporter le graphique…", self)
        self.btn_export_graph.clicked.connect(self.export_graph)
        self.btn_export_graph.setEnabled(False)
        btn_row.addWidget(self.btn_export_csv)
        btn_row.addWidget(self.btn_export_graph)
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
        left_scroll.setMinimumWidth(460)

        # ---------------- Panneau droit ----------------
        self.plot_view = PlotlyView(self)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(left_scroll)
        splitter.addWidget(self.plot_view)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([480, 800])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)

        self.store.changed.connect(self._on_store_changed)
        self.store.meta_changed.connect(self._on_meta_changed)
        self._sync_titrant_controls()
        self._on_store_changed()

    # ------------------------------------------------------------------
    def _last_dir(self) -> str:
        saved = QSettings("Ramanalyze", "Simple").value(_SETTINGS_KEY, "")
        if saved and os.path.isdir(saved):
            return saved
        return os.path.expanduser("~")

    def open_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Choisir des spectres (.txt)", self._last_dir(),
            "Spectres Raman (*.txt);;Tous les fichiers (*)",
        )
        if not paths:
            return
        QSettings("Ramanalyze", "Simple").setValue(_SETTINGS_KEY, os.path.dirname(paths[0]))
        failed = []
        for path in paths:
            if path in self.store.spectra:
                continue
            data = load_spectrum(path)
            if data is None:
                failed.append(os.path.basename(path))
                continue
            self.store.add(path, data)
        if failed:
            self.status.setText("Non lisible(s) : " + ", ".join(failed))

    def remove_selected(self):
        rows = {idx.row() for idx in self.file_table.selectedIndexes()}
        paths = []
        for r in rows:
            item = self.file_table.item(r, 0)
            if item is not None:
                paths.append(item.data(Qt.UserRole))
        if not paths:
            QMessageBox.information(
                self, "Rien à retirer",
                "Sélectionnez d'abord une ou plusieurs lignes dans le tableau.",
            )
            return
        for p in paths:
            self.store.remove(p)

    def clear_all(self):
        if not self.store.paths():
            return
        if QMessageBox.question(self, "Tout vider", "Retirer tous les fichiers chargés ?") == QMessageBox.Yes:
            self.store.clear()

    # ------------------------------------------------------------------
    # Persistance du tableau
    # ------------------------------------------------------------------
    def save_session(self):
        if not self.store.paths():
            QMessageBox.information(self, "Rien à enregistrer", "Aucun spectre chargé.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Enregistrer le tableau", os.path.join(self._last_dir(), "tableau_titration.json"),
            "Tableau Ramanalyze (*.json)",
        )
        if not path:
            return
        if not path.lower().endswith(".json"):
            path += ".json"
        try:
            self.store.save_session(path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Enregistrement impossible", str(exc))
            return
        self.status.setText(f"Tableau enregistré : {os.path.basename(path)}")

    def load_session(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Charger un tableau", self._last_dir(),
            "Tableau Ramanalyze (*.json);;Tous les fichiers (*)",
        )
        if not path:
            return
        try:
            missing = self.store.load_session(path)   # émet « changed »
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Chargement impossible", str(exc))
            return
        self._sync_titrant_controls()
        msg = f"Tableau chargé : {os.path.basename(path)}."
        if missing:
            msg += (f" {len(missing)} fichier(s) introuvable(s) — rechargez-les "
                    f"(même nom) pour récupérer leur volume/série : "
                    + ", ".join(missing[:3]) + ("…" if len(missing) > 3 else ""))
        self.status.setText(msg)

    # ------------------------------------------------------------------
    # Synchronisation des métadonnées (volumes / séries / titrant) via le store
    # ------------------------------------------------------------------
    def _emit_meta(self):
        self._emitting_meta = True
        self.store.meta_changed.emit()
        self._emitting_meta = False

    def _on_meta_changed(self):
        if self._emitting_meta:
            return
        # un autre onglet a modifié volumes/séries/titrant : on rafraîchit l'affichage
        self._sync_titrant_controls()
        self._refresh_file_list()
        self._refresh_fit_curves()

    def _on_titrant_changed(self, *_):
        if self._populating:
            return
        self.store.titrant["conc"] = self.spin_conc.value()
        self.store.titrant["conc_unit"] = self.cmb_conc_unit.currentText()
        self.store.titrant["vol_unit"] = self.cmb_vol_unit.currentText()
        self._emit_meta()

    def _sync_titrant_controls(self):
        self._populating = True
        t = self.store.titrant
        self.spin_conc.setValue(float(t.get("conc", 1.0)))
        self.cmb_conc_unit.setCurrentText(t.get("conc_unit", "µM"))
        self.cmb_vol_unit.setCurrentText(t.get("vol_unit", "µL"))
        self._populating = False

    def _on_table_edited(self, item):
        if self._populating:
            return
        path = item.data(Qt.UserRole)
        if path is None:
            return
        if item.column() == 1:
            self.store.volumes[path] = item.text().strip()
        elif item.column() == 2:
            self.store.series[path] = item.text().strip()
        self._emit_meta()

    def assign_series(self):
        label = self.edit_series.text().strip()
        rows = {idx.row() for idx in self.file_table.selectedIndexes()}
        if not rows:
            QMessageBox.information(
                self, "Aucune sélection",
                "Sélectionnez d'abord des lignes, saisissez un nom de série, puis affectez.",
            )
            return
        self._populating = True
        for r in rows:
            path = self.file_table.item(r, 0).data(Qt.UserRole)
            self.store.series[path] = label
            self.file_table.item(r, 2).setText(label)
        self._populating = False
        self._emit_meta()
        self.status.setText(f"Série « {label or _NO_SERIES} » affectée à {len(rows)} fichier(s).")

    def _refresh_file_list(self):
        self._populating = True
        self.file_table.setRowCount(0)
        for path in self.store.paths():
            row = self.file_table.rowCount()
            self.file_table.insertRow(row)

            name_item = QTableWidgetItem(self.store.name(path))
            name_item.setData(Qt.UserRole, path)
            name_item.setToolTip(path)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.file_table.setItem(row, 0, name_item)

            vol_item = QTableWidgetItem(self.store.volumes.get(path, ""))
            vol_item.setData(Qt.UserRole, path)
            self.file_table.setItem(row, 1, vol_item)

            ser_item = QTableWidgetItem(self.store.series.get(path, ""))
            ser_item.setData(Qt.UserRole, path)
            self.file_table.setItem(row, 2, ser_item)
        self._populating = False

    def _invalidate_detection(self):
        """Le jeu de fichiers a changé : la détection précédente n'est plus valable."""
        self._populating = True
        self.list_peaks.clear()
        self._populating = False
        self._centers = []
        self._corr = {}
        self._plot_rows = []
        self._last_fig = None
        self.btn_export_csv.setEnabled(False)
        self.btn_export_graph.setEnabled(False)

    def _on_store_changed(self):
        self._refresh_file_list()
        self._invalidate_detection()
        self._refresh_fit_curves()
        n = len(self.store.paths())
        self.status.setText("Aucun spectre" if n == 0 else f"{n} spectre(s) chargé(s). Cliquez « Détecter les pics ».")

    # ------------------------------------------------------------------
    def _corrected_spectra(self):
        do_baseline = self.chk_baseline.isChecked()
        corr = {}
        for path in self.store.paths():
            x, y = self.store.spectra[path]
            x = np.asarray(x, dtype=float)
            y = np.asarray(y, dtype=float)
            corr[path] = (x, _baseline_corrected(x, y) if do_baseline else y)
        return corr

    def detect_peaks(self):
        if not self.store.paths():
            QMessageBox.information(self, "Aucun spectre", "Chargez d'abord des fichiers .txt.")
            return
        wmin, wmax = self.spin_min.value(), self.spin_max.value()
        if wmax <= wmin:
            QMessageBox.warning(self, "Fenêtre invalide", "La borne max doit être supérieure à la borne min.")
            return

        self._corr = self._corrected_spectra()
        prom = self.spin_prom.value() / 100.0
        tol = self.spin_tol.value()

        detections = []
        for si, path in enumerate(self.store.paths()):
            x, y = self._corr[path]
            for pos in _detect_peaks(x, y, wmin, wmax, prom):
                detections.append((pos, si))

        n_spec = len(self.store.paths())
        min_support = min(self.spin_support.value(), n_spec)
        centers = [(c, s) for c, s in _cluster(detections, tol) if s >= min_support]

        self._centers = centers
        self._populate_peak_list()
        self._refresh_fit_curves()

        if not centers:
            self.status.setText("Aucun pic retenu. Baissez la proéminence ou la présence minimale.")
        else:
            self.status.setText(f"{len(centers)} pic(s) détecté(s). Cochez ceux à suivre puis « Tracer ».")

    def _populate_peak_list(self):
        self._populating = True
        self.list_peaks.clear()
        n_spec = len(self.store.paths())
        for center, support in self._centers:
            item = QListWidgetItem(f"{center:.0f} cm⁻¹   (présent dans {support}/{n_spec})")
            item.setData(Qt.UserRole, center)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.list_peaks.addItem(item)
        self._populating = False

    def _set_all_checked(self, checked: bool):
        self._populating = True
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list_peaks.count()):
            self.list_peaks.item(i).setCheckState(state)
        self._populating = False
        self._refresh_fit_curves()
        if self._last_fig is not None:
            self.plot_evolution()

    def _on_peak_toggled(self, item):
        if self._populating:
            return
        self._refresh_fit_curves()
        if self._last_fig is not None:
            self.plot_evolution()

    def _checked_centers(self):
        out = []
        for i in range(self.list_peaks.count()):
            it = self.list_peaks.item(i)
            if it.checkState() == Qt.Checked:
                out.append(float(it.data(Qt.UserRole)))
        return out

    # ------------------------------------------------------------------
    def _on_sigmoid_toggled(self, checked):
        self.cmb_fit_curve.setEnabled(checked)
        self._refresh_fit_curves()
        if self._last_fig is not None:
            self.plot_evolution()

    def _on_fit_curve_changed(self, *_):
        if self._populating:
            return
        if self._last_fig is not None and self.chk_sigmoid.isChecked():
            self.plot_evolution()

    def _refresh_fit_curves(self):
        """Reconstruit la liste des courbes (série × pic coché) pour le sélecteur."""
        self._populating = True
        keep = self.cmb_fit_curve.currentData()
        self.cmb_fit_curve.clear()
        self.cmb_fit_curve.addItem("Toutes les courbes", None)
        series_order = []
        for p in self.store.paths():
            lab = self._series_label(p)
            if lab not in series_order:
                series_order.append(lab)
        centers = sorted(self._checked_centers())
        multi = len(series_order) > 1
        for lab in series_order:
            for c in centers:
                name = f"{lab} — {c:.0f} cm⁻¹" if multi else f"{c:.0f} cm⁻¹"
                self.cmb_fit_curve.addItem(name, (lab, c))
        idx = self.cmb_fit_curve.findData(keep)
        if idx >= 0:
            self.cmb_fit_curve.setCurrentIndex(idx)
        self._populating = False

    def _series_label(self, path):
        return self.store.series.get(path, "").strip() or _NO_SERIES

    def plot_evolution(self):
        try:
            self._plot_evolution_impl()
        except Exception as exc:  # noqa: BLE001
            import traceback
            QMessageBox.critical(
                self, "Erreur de tracé",
                f"Le tracé a échoué :\n{exc}\n\n{traceback.format_exc()}",
            )

    def _plot_evolution_impl(self):
        centers = sorted(self._checked_centers())
        if not centers:
            QMessageBox.information(
                self, "Aucun pic coché",
                "Détectez les pics puis cochez au moins un pic à suivre.",
            )
            return
        if not self._corr:
            self._corr = self._corrected_spectra()

        tol = self.spin_tol.value()
        wmin, wmax = self.spin_min.value(), self.spin_max.value()
        use_quantity = self.cmb_xaxis.currentText().startswith("Quantité")

        amounts = tu.amounts_mol(self.store.paths(), self.store.volumes, self.store.titrant) if use_quantity else None
        if use_quantity:
            valid = list(amounts.values())
            if not valid:
                QMessageBox.warning(
                    self, "Volumes manquants",
                    "Renseignez la concentration du titrant et au moins un volume "
                    "de titrant dans le tableau.",
                )
                return
            unit_label, unit_factor = tu.pick_amount_unit(max(valid))
            x_title = f"Quantité de titrant ({unit_label})"
        else:
            x_title = "Spectre"

        def xval(path):
            if use_quantity:
                return amounts[path] / unit_factor
            return self.store.name(path)

        # regroupement par série (ordre d'apparition)
        series_order = []
        for p in self.store.paths():
            lab = self._series_label(p)
            if lab not in series_order:
                series_order.append(lab)

        inten = {}
        for p in self.store.paths():
            x, y = self._corr[p]
            inten[p] = {c: _measure_at(x, y, c, tol) for c in centers}

        title = self.edit_title.text().strip() or f"Évolution des pics · {wmin:.0f}–{wmax:.0f} cm⁻¹"
        self._last_file_base = f"evolution_pics_{wmin:.0f}_{wmax:.0f}"

        do_fit = self.chk_sigmoid.isChecked()
        fit_target = self.cmb_fit_curve.currentData() if do_fit else None  # None = toutes
        fit_all = do_fit and fit_target is None
        fig = go.Figure()
        self._plot_rows = []
        n_used = n_skipped = n_fits = 0
        eq_points = []
        multi = len(series_order) > 1
        for si, lab in enumerate(series_order):
            grp = [p for p in self.store.paths() if self._series_label(p) == lab]
            if use_quantity:
                grp = [p for p in grp if p in amounts]
                grp.sort(key=lambda p: amounts[p])
            if not grp:
                continue
            family = _FAMILIES[si % len(_FAMILIES)]
            xs = [xval(p) for p in grp]
            names = [self.store.name(p) for p in grp]
            n_used += len(grp)
            for ci, center in enumerate(centers):
                color = family[ci % len(family)]
                ys = [None if np.isnan(inten[p][center]) else inten[p][center] for p in grp]
                trace_name = f"{lab} — {center:.0f} cm⁻¹" if multi else f"{center:.0f} cm⁻¹"
                fig.add_trace(go.Scatter(
                    x=xs, y=ys, mode="lines+markers", name=trace_name,
                    legendgroup=lab, legendgrouptitle_text=(lab if multi else None),
                    connectgaps=False,
                    line=dict(width=ps.LINE_WIDTH, color=color),   # données = trait plein
                    marker=ps.marker(color, si),                   # croix
                    text=names,
                    hovertemplate=f"{lab}<br>pic {center:.0f} cm⁻¹<br>%{{text}}<br>"
                                  f"x=%{{x}}<br>I=%{{y:.4g}}<extra></extra>",
                ))
                if do_fit and use_quantity and (fit_all or fit_target == (lab, center)):
                    x_eq = self._add_sigmoid(fig, xs, ys, color, lab, trace_name,
                                             mark_eq=not fit_all)
                    if x_eq is not None:
                        n_fits += 1
                        eq_points.append((trace_name, x_eq))
            for p in grp:
                self._plot_rows.append((lab, p, xval(p)))

        if use_quantity:
            n_skipped = len(self.store.paths()) - n_used

        ps.apply(
            fig, title=title, x_title=x_title, y_title="Intensité du pic (a.u.)",
            legend_title=("Séries / pics" if multi else "Pics suivis"),
            groupclick="toggleitem",   # clic = 1 seule courbe
        )
        if not use_quantity:
            fig.update_xaxes(tickangle=-45)
        self.plot_view.show_figure(fig)
        self._last_fig = fig

        self._inten = inten
        self._centers_plotted = centers
        self._x_title = x_title

        self.btn_export_graph.setEnabled(True)
        self.btn_export_csv.setEnabled(True)
        msg = f"{len(centers)} pic(s) × {len(series_order)} série(s) sur {n_used} spectre(s)."
        if n_skipped:
            msg += f" {n_skipped} sans volume ignoré(s)."
        if do_fit and not use_quantity:
            msg += " Marquage x_eq ignoré : passez l'axe X sur « Quantité de matière »."
        elif do_fit:
            seuil = self.spin_plateau.value()
            if n_fits and not fit_all:
                nm, xe = eq_points[0]
                msg += f" {nm} — x_eq (fin de palier, {seuil:.0f} %) = {xe:.4g}."
            elif n_fits:
                apercu = ", ".join(f"{nm} : x_eq={xe:.4g}" for nm, xe in eq_points[:3])
                suite = "…" if len(eq_points) > 3 else ""
                msg += f" {n_fits} x_eq marqué(s) ({seuil:.0f} %) — {apercu}{suite}"
            else:
                msg += " x_eq : aucun ajustement convergent (≥ 4 points requis)."
        self.status.setText(msg)

    def _add_sigmoid(self, fig, xs, ys, color, group, label, mark_eq=False):
        """Trace le fit en tirets (paliers + descente) et marque x_eq = fin de palier.

        x_eq = abscisse où le signal a fini de bouger (nouveau palier atteint, seuil
        réglable). Renvoie x_eq ou None.
        """
        popt = tu.fit_sigmoid(xs, ys)
        if popt is None:
            return None
        seuil = self.spin_plateau.value()
        bounds = tu.transition_bounds(popt, seuil / 100.0)
        if bounds is None:
            return None
        x_eq = bounds[1]                      # « x_eq » = fin de palier
        y_eq = float(tu.sigmoid(x_eq, *popt))
        xa = np.array([x for x, y in zip(xs, ys) if y is not None], dtype=float)

        # courbe fittée en tirets : paliers (parties plates) + descente/montée.
        # Pour la courbe choisie, on l'étend jusqu'aux paliers (≈99 %) pour bien les voir.
        if mark_eq:
            b99 = tu.transition_bounds(popt, 0.99) or bounds
            lo, hi = min(xa.min(), b99[0]), max(xa.max(), b99[1])
            width = 1.7
        else:
            lo, hi = xa.min(), xa.max()
            width = 1.3
        xf = np.linspace(lo, hi, 300)
        fig.add_trace(go.Scatter(
            x=xf, y=tu.sigmoid(xf, *popt), mode="lines",
            name=f"{label} (fit)", legendgroup=group, showlegend=False,
            line=dict(width=width, dash="dash", color=color), opacity=0.9,
            hovertemplate=f"{label} · fit (paliers + descente)<extra></extra>",
        ))
        # marqueur x_eq (fin de palier)
        fig.add_trace(go.Scatter(
            x=[x_eq], y=[y_eq], mode="markers",
            name="x_eq (fin de palier)", legendgroup=group, showlegend=False,
            marker=dict(size=13, symbol="diamond", color=color, line=dict(width=1.4, color="#000")),
            hovertemplate=f"x_eq (fin de palier, {seuil:.0f} %)<br>x={x_eq:.4g}<extra></extra>",
        ))
        if mark_eq:
            fig.add_vline(
                x=x_eq, line=dict(color=color, dash="dot", width=1.6),
                annotation_text=f"x_eq ≈ {x_eq:.4g}", annotation_position="top",
                annotation_font=dict(size=12, color=color),
            )
        return x_eq

    # ------------------------------------------------------------------
    def export_csv(self):
        if self._last_fig is None or not getattr(self, "_plot_rows", None):
            QMessageBox.information(self, "Rien à exporter", "Tracez d'abord l'évolution.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter les intensités", os.path.join(self._last_dir(), self._last_file_base + ".csv"),
            "CSV (*.csv)",
        )
        if not path:
            return
        vol_unit = self.store.titrant.get("vol_unit", "µL")
        x_title = getattr(self, "_x_title", "Spectre")
        centers = self._centers_plotted
        include_x = x_title != "Spectre"
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                w = csv.writer(f, delimiter=";")
                header = ["Série", "Fichier", f"Volume titrant ({vol_unit})"]
                if include_x:
                    header.append(x_title)
                header += [f"{c:.0f} cm-1" for c in centers]
                w.writerow(header)
                for lab, p, xv in self._plot_rows:
                    row = [lab, self.store.name(p), self.store.volumes.get(p, "")]
                    if include_x:
                        row.append(f"{xv:.6g}")
                    for c in centers:
                        v = self._inten[p][c]
                        row.append("" if (v is None or np.isnan(v)) else f"{v:.6g}")
                    w.writerow(row)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Export impossible", str(exc))
            return
        self.status.setText(f"Intensités exportées : {os.path.basename(path)}")

    def export_graph(self):
        if self._last_fig is None:
            QMessageBox.information(self, "Rien à exporter", "Tracez d'abord l'évolution.")
            return
        default = os.path.join(self._last_dir(), self._last_file_base + ".png")
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter le graphique", default,
            "Image PNG (*.png);;Image vectorielle SVG (*.svg);;Document PDF (*.pdf);;"
            "Page HTML interactive (*.html)",
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if not ext:
            ext = ".png"
            path += ext
        try:
            if ext == ".html":
                self._last_fig.write_html(path, include_plotlyjs="inline")
            else:
                self._last_fig.write_image(path, width=1200, height=800, scale=2)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "Export impossible",
                f"Impossible d'exporter en {ext} :\n{exc}\n\n"
                "Astuce : l'export HTML fonctionne toujours.",
            )
            return
        self.status.setText(f"Graphique exporté : {os.path.basename(path)}")
