"""Onglet « Suivi de pic ».

But : dans une fenêtre spectrale (par défaut **1200–1500 cm⁻¹**), repérer **tous**
les pics présents, puis suivre **chacun individuellement** d'un spectre à l'autre.

Démarche :
1. On détecte les pics (scipy.find_peaks) dans la fenêtre, pour chaque spectre.
2. On **apparie** ces pics d'un spectre à l'autre en regroupant ceux dont la
   position est proche (tolérance réglable) → chaque groupe = un « pic suivi »,
   caractérisé par une position de référence (centre du groupe).
3. Pour chaque pic suivi, on mesure son intensité dans chaque spectre (maximum
   dans ±tolérance autour du centre) et on trace **une courbe par pic**.

Une correction de ligne de base (modpoly, comme Ramanalyze) est appliquée par
défaut avant toute mesure.
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

_SETTINGS_KEY = "simple/last_dir"

# Facteurs vers les unités SI de base (mol/L et L).
_CONC_FACTORS = {"nM": 1e-9, "µM": 1e-6, "mM": 1e-3, "M": 1.0}
_VOL_FACTORS = {"µL": 1e-6, "mL": 1e-3, "L": 1.0}
# Échelles d'affichage de la quantité de matière (mol).
_AMOUNT_UNITS = [("mol", 1.0), ("mmol", 1e-3), ("µmol", 1e-6),
                 ("nmol", 1e-9), ("pmol", 1e-12), ("fmol", 1e-15)]


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
        # repli : le maximum global de la fenêtre
        return [float(xb[int(np.argmax(yb))])]


def _cluster(detections, tol):
    """Regroupe des pics proches en « pics suivis ».

    detections : liste de (position_cm, indice_spectre).
    Renvoie une liste de (centre_cm, support) triée par centre,
    où *support* = nombre de spectres distincts contribuant au groupe.
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


def _parse_num(text):
    """Convertit un texte (virgule ou point décimal) en float, ou NaN."""
    if text is None:
        return np.nan
    t = str(text).strip().replace(",", ".")
    if not t:
        return np.nan
    try:
        return float(t)
    except ValueError:
        return np.nan


def _pick_amount_unit(max_mol):
    """Choisit l'unité d'affichage (mol→fmol) la plus lisible pour la quantité."""
    if not np.isfinite(max_mol) or max_mol <= 0:
        return "mol", 1.0
    for label, factor in _AMOUNT_UNITS:
        if max_mol / factor >= 1.0:
            return label, factor
    return _AMOUNT_UNITS[-1]  # fmol


class PeakTrackerTab(QWidget):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store
        self._corr = {}            # path -> (x, y_corrigé) calculé à la détection
        self._centers = []         # liste de (centre_cm, support) détectés
        self._volumes = {}         # path -> volume de titrant saisi (texte)
        self._last_fig = None
        self._last_file_base = "evolution_pics"
        self._populating = False

        # ---------------- Panneau gauche ----------------
        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(QLabel("<b>Suivi individuel des pics</b>", self))
        hint = QLabel(
            "On repère tous les pics de la fenêtre, on les apparie d'un spectre à "
            "l'autre, puis on suit chacun séparément. Les fichiers ajoutés ici (ou "
            "envoyés depuis le Visualiseur) sont analysés.", self
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888;")
        left_layout.addWidget(hint)

        self.btn_open = QPushButton("📂  Ajouter des fichiers…", self)
        self.btn_open.clicked.connect(self.open_files)
        left_layout.addWidget(self.btn_open)

        # --- Fichiers chargés + volume de titrant par échantillon ---
        files_box = QGroupBox("Fichiers chargés & volume de titrant", self)
        files_layout = QVBoxLayout(files_box)
        files_hint = QLabel(
            "Double-cliquez la colonne <i>Volume</i> pour saisir le volume de titrant "
            "ajouté à chaque échantillon. Sélectionnez des lignes pour les retirer "
            "(Ctrl/Maj pour une sélection multiple).", self
        )
        files_hint.setWordWrap(True)
        files_hint.setStyleSheet("color: #888;")
        files_layout.addWidget(files_hint)

        self.file_table = QTableWidget(0, 2, self)
        self.file_table.setMinimumHeight(140)
        self.file_table.setHorizontalHeaderLabels(["Fichier", "Volume titrant"])
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_table.itemChanged.connect(self._on_volume_changed)
        files_layout.addWidget(self.file_table)

        files_btns = QHBoxLayout()
        self.btn_remove = QPushButton("Retirer la sélection", self)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("Tout vider", self)
        self.btn_clear.clicked.connect(self.clear_all)
        files_btns.addWidget(self.btn_remove)
        files_btns.addWidget(self.btn_clear)
        files_layout.addLayout(files_btns)
        left_layout.addWidget(files_box)

        # --- Titrant : concentration de la solution + unité + unité de volume ---
        titrant_box = QGroupBox("Titrant (pour l'axe « quantité de matière »)", self)
        titrant_form = QFormLayout(titrant_box)

        conc_row = QHBoxLayout()
        self.spin_conc = QDoubleSpinBox(self)
        self.spin_conc.setRange(0.0, 1e9)
        self.spin_conc.setDecimals(4)
        self.spin_conc.setValue(1.0)
        conc_row.addWidget(self.spin_conc, 1)
        self.cmb_conc_unit = QComboBox(self)
        self.cmb_conc_unit.addItems(list(_CONC_FACTORS.keys()))  # nM, µM, mM, M
        self.cmb_conc_unit.setCurrentText("µM")
        conc_row.addWidget(self.cmb_conc_unit)
        titrant_form.addRow("Concentration :", conc_row)

        self.cmb_vol_unit = QComboBox(self)
        self.cmb_vol_unit.addItems(list(_VOL_FACTORS.keys()))    # µL, mL, L
        self.cmb_vol_unit.setCurrentText("µL")
        titrant_form.addRow("Unité des volumes :", self.cmb_vol_unit)

        self.cmb_xaxis = QComboBox(self)
        self.cmb_xaxis.addItems(["Nom de spectre", "Quantité de matière de titrant"])
        self.cmb_xaxis.setToolTip(
            "Abscisse du graphe : noms des spectres, ou quantité de matière = "
            "concentration × volume de titrant."
        )
        titrant_form.addRow("Axe X :", self.cmb_xaxis)
        left_layout.addWidget(titrant_box)

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
        self.list_peaks.setMinimumHeight(180)
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
            self.store.remove(p)   # émet « changed » -> _on_store_changed

    def clear_all(self):
        if not self.store.paths():
            return
        if QMessageBox.question(
            self, "Tout vider",
            "Retirer tous les fichiers chargés ?",
        ) == QMessageBox.Yes:
            self.store.clear()

    def _on_volume_changed(self, item):
        if self._populating or item.column() != 1:
            return
        path = item.data(Qt.UserRole)
        if path is not None:
            self._volumes[path] = item.text().strip()

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

            vol_item = QTableWidgetItem(self._volumes.get(path, ""))
            vol_item.setData(Qt.UserRole, path)
            self.file_table.setItem(row, 1, vol_item)
        # purge des volumes des fichiers disparus
        live = set(self.store.paths())
        self._volumes = {p: v for p, v in self._volumes.items() if p in live}
        self._populating = False

    def _invalidate_detection(self):
        """Le jeu de fichiers a changé : la détection précédente n'est plus valable."""
        self._populating = True
        self.list_peaks.clear()
        self._populating = False
        self._centers = []
        self._corr = {}
        self._matrix = {}
        self._last_fig = None
        self.btn_export_csv.setEnabled(False)
        self.btn_export_graph.setEnabled(False)

    def _on_store_changed(self):
        self._refresh_file_list()
        self._invalidate_detection()
        n = len(self.store.paths())
        self.status.setText("Aucun spectre" if n == 0 else f"{n} spectre(s) chargé(s). Cliquez « Détecter les pics ».")

    # ------------------------------------------------------------------
    def _corrected_spectra(self):
        """Renvoie {path: (x, y)} avec ligne de base soustraite selon l'option."""
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

        # 1) détecter les pics par spectre, en gardant l'indice du spectre
        detections = []
        for si, path in enumerate(self.store.paths()):
            x, y = self._corr[path]
            for pos in _detect_peaks(x, y, wmin, wmax, prom):
                detections.append((pos, si))

        # 2) apparier (regrouper) + filtrer par présence minimale
        n_spec = len(self.store.paths())
        min_support = min(self.spin_support.value(), n_spec)
        centers = [(c, s) for c, s in _cluster(detections, tol) if s >= min_support]

        self._centers = centers
        self._populate_peak_list()

        if not centers:
            self.status.setText(
                "Aucun pic retenu. Baissez la proéminence ou la présence minimale."
            )
        else:
            self.status.setText(
                f"{len(centers)} pic(s) détecté(s). Cochez ceux à suivre puis « Tracer »."
            )

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
        if self._last_fig is not None:
            self.plot_evolution()

    def _on_peak_toggled(self, item):
        if self._populating:
            return
        # re-tracer automatiquement si un graphe existe déjà
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
    def _amounts_mol(self):
        """{path: quantité de matière de titrant en mol} pour les volumes valides.

        n = concentration(mol/L) × volume(L), avec les unités choisies dans l'UI.
        """
        c_factor = _CONC_FACTORS[self.cmb_conc_unit.currentText()]
        v_factor = _VOL_FACTORS[self.cmb_vol_unit.currentText()]
        conc_molL = self.spin_conc.value() * c_factor
        out = {}
        for path in self.store.paths():
            vol = _parse_num(self._volumes.get(path, ""))
            if np.isnan(vol):
                continue
            out[path] = conc_molL * (vol * v_factor)
        return out

    def plot_evolution(self):
        centers = self._checked_centers()
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

        # --- choix de l'abscisse ---
        n_skipped = 0
        if use_quantity:
            amounts = self._amounts_mol()
            plot_paths = sorted((p for p in self.store.paths() if p in amounts),
                                key=lambda p: amounts[p])
            if not plot_paths:
                QMessageBox.warning(
                    self, "Volumes manquants",
                    "Renseignez la concentration du titrant et au moins un volume "
                    "de titrant dans le tableau.",
                )
                return
            n_skipped = len(self.store.paths()) - len(plot_paths)
            unit_label, unit_factor = _pick_amount_unit(max(amounts[p] for p in plot_paths))
            xvals = [amounts[p] / unit_factor for p in plot_paths]
            x_title = f"Quantité de titrant ({unit_label})"
        else:
            plot_paths = list(self.store.paths())
            xvals = [self.store.name(p) for p in plot_paths]
            x_title = "Spectre"

        names = [self.store.name(p) for p in plot_paths]
        auto_title = f"Évolution des pics · fenêtre {wmin:.0f}–{wmax:.0f} cm⁻¹"
        title = self.edit_title.text().strip() or auto_title
        self._last_file_base = f"evolution_pics_{wmin:.0f}_{wmax:.0f}"

        fig = go.Figure()
        # matrice pour l'export : center -> liste d'intensités (alignée sur plot_paths)
        self._matrix = {}
        for center in sorted(centers):
            ys = []
            for path in plot_paths:
                x, y = self._corr[path]
                val = _measure_at(x, y, center, tol)
                ys.append(None if np.isnan(val) else val)
            self._matrix[center] = ys
            fig.add_trace(go.Scatter(
                x=xvals, y=ys, mode="lines+markers", name=f"{center:.0f} cm⁻¹",
                connectgaps=False, marker=dict(size=9), line=dict(width=2),
                text=names,
                hovertemplate=f"pic {center:.0f} cm⁻¹<br>%{{text}}<br>"
                              f"x=%{{x}}<br>I=%{{y:.4g}}<extra></extra>",
            ))

        fig.update_layout(
            title=title,
            xaxis_title=x_title,
            yaxis_title="Intensité du pic (a.u.)",
            template="plotly_white",
            margin=dict(l=80, r=30, t=60, b=120),
            font=dict(size=14),
            legend=dict(title="Pics suivis"),
        )
        if not use_quantity:
            fig.update_xaxes(tickangle=-45)
        self.plot_view.show_figure(fig)
        self._last_fig = fig

        # mémorisé pour l'export CSV
        self._plot_paths = plot_paths
        self._xvals = xvals
        self._x_title = x_title

        self.btn_export_graph.setEnabled(True)
        self.btn_export_csv.setEnabled(True)
        msg = f"{len(centers)} pic(s) suivi(s) sur {len(plot_paths)} spectre(s)."
        if n_skipped:
            msg += f" {n_skipped} sans volume ignoré(s)."
        self.status.setText(msg)

    # ------------------------------------------------------------------
    def export_csv(self):
        if self._last_fig is None or not getattr(self, "_matrix", None):
            QMessageBox.information(self, "Rien à exporter", "Tracez d'abord l'évolution.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter les intensités", os.path.join(self._last_dir(), self._last_file_base + ".csv"),
            "CSV (*.csv)",
        )
        if not path:
            return
        plot_paths = getattr(self, "_plot_paths", self.store.paths())
        xvals = getattr(self, "_xvals", [self.store.name(p) for p in plot_paths])
        x_title = getattr(self, "_x_title", "Spectre")
        vol_unit = self.cmb_vol_unit.currentText()
        centers = sorted(self._matrix.keys())
        include_x = x_title != "Spectre"   # colonne quantité seulement si pertinent
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                w = csv.writer(f, delimiter=";")
                header = ["Fichier", f"Volume titrant ({vol_unit})"]
                if include_x:
                    header.append(x_title)
                header += [f"{c:.0f} cm-1" for c in centers]
                w.writerow(header)
                for i, p in enumerate(plot_paths):
                    row = [self.store.name(p), self._volumes.get(p, "")]
                    if include_x:
                        row.append(f"{xvals[i]:.6g}")
                    for c in centers:
                        v = self._matrix[c][i]
                        row.append("" if v is None else f"{v:.6g}")
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
            "Image PNG (*.png);;Document PDF (*.pdf);;Page HTML interactive (*.html)",
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
