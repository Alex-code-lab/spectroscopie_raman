"""Onglet Titration.

On saisit, pour chaque spectre, le **volume de titrant** ajouté et (optionnellement)
une **série** ; on indique la **concentration** de la solution de titrant et ses unités.
L'abscisse est alors la **quantité de matière de titrant** (= concentration × volume),
exactement comme dans l'onglet « Suivi de pic » (réglages partagés via le store).

On mesure l'intensité d'un pic (ou le ratio de deux pics) par spectre, et on trace
une **courbe de titration par série**. Options : correction de ligne de base et
ajustement sigmoïde (point d'équivalence).
"""

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
    QComboBox,
    QListWidget,
    QListWidgetItem,
    QTableWidget,
    QTableWidgetItem,
    QSplitter,
    QScrollArea,
    QFileDialog,
    QDoubleSpinBox,
    QCheckBox,
    QMessageBox,
    QGroupBox,
    QFormLayout,
    QHeaderView,
    QAbstractItemView,
)
import plotly.graph_objects as go

from spectrum_loader import load_spectrum
from plot_view import PlotlyView
import peak_presets
import titrant_utils as tu

_SETTINGS_KEY = "simple/last_dir"

_PALETTE = ["#0057b8", "#d9534f", "#5cb85c", "#f0ad4e", "#9b59b6",
            "#17a2b8", "#e83e8c", "#6c757d", "#20c997", "#fd7e14"]
_NO_SERIES = "(sans série)"


def _peak_intensity(x, y, peak, tol):
    """Intensité maximale dans la fenêtre [peak-tol, peak+tol], ou NaN."""
    mask = (x >= peak - tol) & (x <= peak + tol)
    if not mask.any():
        return np.nan
    return float(np.max(y[mask]))


def _baseline_corrected(x, y, poly_order=5):
    """Soustrait une ligne de base modpoly (pybaselines). Renvoie y brut en cas d'échec."""
    try:
        from pybaselines import Baseline

        baseline, _ = Baseline(x).modpoly(y, poly_order=poly_order)
        return y - baseline
    except Exception as exc:  # noqa: BLE001
        print(f"[titration] Baseline impossible : {exc}")
        return y


class TitrationTab(QWidget):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store
        self._last_fig = None
        self._last_file_base = "titration"
        self._populating = False
        self._emitting_meta = False

        # ---------------- Panneau gauche ----------------
        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(QLabel("<b>Volume de titrant & série par échantillon</b>", self))
        hint = QLabel(
            "Double-cliquez les colonnes <i>Volume</i> et <i>Série</i>. L'abscisse est la "
            "quantité de matière (concentration × volume). Réglages partagés avec « Suivi de pic ».",
            self,
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888;")
        left_layout.addWidget(hint)

        self.btn_open = QPushButton("📂  Ajouter des fichiers…", self)
        self.btn_open.clicked.connect(self.open_files)
        left_layout.addWidget(self.btn_open)

        self.table = QTableWidget(0, 3, self)
        self.table.setMinimumHeight(170)
        self.table.setHorizontalHeaderLabels(["Fichier", "Volume titrant", "Série"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.itemChanged.connect(self._on_table_edited)
        left_layout.addWidget(self.table, 1)

        series_row = QHBoxLayout()
        self.edit_series = QLineEdit(self)
        self.edit_series.setPlaceholderText("Nom de série (ex. Série 1)")
        series_row.addWidget(self.edit_series, 1)
        self.btn_assign = QPushButton("Affecter à la sélection", self)
        self.btn_assign.clicked.connect(self.assign_series)
        series_row.addWidget(self.btn_assign)
        left_layout.addLayout(series_row)

        file_btns = QHBoxLayout()
        self.btn_remove = QPushButton("Retirer la sélection", self)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("Tout vider", self)
        self.btn_clear.clicked.connect(self.clear_all)
        file_btns.addWidget(self.btn_remove)
        file_btns.addWidget(self.btn_clear)
        left_layout.addLayout(file_btns)

        session_btns = QHBoxLayout()
        self.btn_save_session = QPushButton("💾 Enregistrer le tableau…", self)
        self.btn_save_session.clicked.connect(self.save_session)
        self.btn_load_session = QPushButton("📥 Charger un tableau…", self)
        self.btn_load_session.clicked.connect(self.load_session)
        session_btns.addWidget(self.btn_save_session)
        session_btns.addWidget(self.btn_load_session)
        left_layout.addLayout(session_btns)

        # --- Titrant : concentration + unités ---
        titrant_box = QGroupBox("Titrant", self)
        titrant_form = QFormLayout(titrant_box)
        conc_row = QHBoxLayout()
        self.spin_conc = QDoubleSpinBox(self)
        self.spin_conc.setRange(0.0, 1e9)
        self.spin_conc.setDecimals(4)
        self.spin_conc.setValue(1.0)
        conc_row.addWidget(self.spin_conc, 1)
        self.cmb_conc_unit = QComboBox(self)
        self.cmb_conc_unit.addItems(list(tu.CONC_FACTORS.keys()))
        self.cmb_conc_unit.setCurrentText("µM")
        conc_row.addWidget(self.cmb_conc_unit)
        titrant_form.addRow("Concentration :", conc_row)
        self.cmb_vol_unit = QComboBox(self)
        self.cmb_vol_unit.addItems(list(tu.VOL_FACTORS.keys()))
        self.cmb_vol_unit.setCurrentText("µL")
        titrant_form.addRow("Unité des volumes :", self.cmb_vol_unit)
        left_layout.addWidget(titrant_box)

        self.spin_conc.valueChanged.connect(self._on_titrant_changed)
        self.cmb_conc_unit.currentTextChanged.connect(self._on_titrant_changed)
        self.cmb_vol_unit.currentTextChanged.connect(self._on_titrant_changed)

        # --- Combinaisons de pics par source lumineuse ---
        src_box = QGroupBox("Combinaisons de pics (source lumineuse)", self)
        src_layout = QVBoxLayout(src_box)
        src_row = QHBoxLayout()
        src_row.addWidget(QLabel("Source :", self))
        self.cmb_source = QComboBox(self)
        self.cmb_source.addItems(peak_presets.sources())
        self.cmb_source.currentTextChanged.connect(self._refresh_pairs)
        src_row.addWidget(self.cmb_source, 1)
        src_layout.addLayout(src_row)
        src_layout.addWidget(QLabel("Double-cliquez une paire pour charger un ratio :", self))
        self.list_pairs = QListWidget(self)
        self.list_pairs.setMaximumHeight(120)
        self.list_pairs.itemDoubleClicked.connect(self._use_pair)
        src_layout.addWidget(self.list_pairs)
        left_layout.addWidget(src_box)

        # --- Pic(s) à mesurer ---
        params = QGroupBox("Pic à mesurer", self)
        form = QFormLayout(params)
        self.spin_peak = QDoubleSpinBox(self)
        self.spin_peak.setRange(0.0, 5000.0)
        self.spin_peak.setDecimals(1)
        self.spin_peak.setValue(1000.0)
        self.spin_peak.setSuffix(" cm⁻¹")
        form.addRow("Pic 1 :", self.spin_peak)
        self.spin_tol = QDoubleSpinBox(self)
        self.spin_tol.setRange(0.5, 200.0)
        self.spin_tol.setDecimals(1)
        self.spin_tol.setValue(5.0)
        self.spin_tol.setSuffix(" cm⁻¹")
        form.addRow("Tolérance :", self.spin_tol)
        self.chk_ratio = QCheckBox("Faire le ratio avec un 2ᵉ pic", self)
        self.chk_ratio.toggled.connect(self._on_ratio_toggled)
        form.addRow(self.chk_ratio)
        self.spin_peak2 = QDoubleSpinBox(self)
        self.spin_peak2.setRange(0.0, 5000.0)
        self.spin_peak2.setDecimals(1)
        self.spin_peak2.setValue(1500.0)
        self.spin_peak2.setSuffix(" cm⁻¹")
        self.spin_peak2.setEnabled(False)
        form.addRow("Pic 2 :", self.spin_peak2)
        self.chk_baseline = QCheckBox("Corriger la ligne de base", self)
        self.chk_baseline.setChecked(True)
        form.addRow(self.chk_baseline)
        self.chk_sigmoid = QCheckBox("Ajuster une sigmoïde (point d'équivalence)", self)
        form.addRow(self.chk_sigmoid)
        left_layout.addWidget(params)

        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Titre :", self))
        self.edit_title = QLineEdit(self)
        self.edit_title.setPlaceholderText("Titre automatique (laisser vide)")
        title_row.addWidget(self.edit_title, 1)
        left_layout.addLayout(title_row)

        self.btn_plot = QPushButton("Tracer la titration", self)
        self.btn_plot.setStyleSheet(
            "background-color: #0057b8; color: white; font-weight: 700; padding: 8px;"
        )
        self.btn_plot.clicked.connect(self.plot_titration)
        left_layout.addWidget(self.btn_plot)

        self.btn_export = QPushButton("Exporter le graphique…", self)
        self.btn_export.clicked.connect(self.export_graph)
        self.btn_export.setEnabled(False)
        left_layout.addWidget(self.btn_export)

        self.status = QLabel("", self)
        self.status.setStyleSheet("color: #888;")
        self.status.setWordWrap(True)
        left_layout.addWidget(self.status)

        left_scroll = QScrollArea(self)
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        left_scroll.setMinimumWidth(450)

        self.plot_view = PlotlyView(self)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(left_scroll)
        splitter.addWidget(self.plot_view)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([470, 810])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)

        self.store.changed.connect(self._refresh_table)
        self.store.meta_changed.connect(self._on_meta_changed)
        self._sync_titrant_controls()
        self._refresh_table()
        self._refresh_pairs()

    # ------------------------------------------------------------------
    def _refresh_pairs(self):
        self.list_pairs.clear()
        for a, b in peak_presets.pairs_for(self.cmb_source.currentText()):
            item = QListWidgetItem(f"{a} / {b} cm⁻¹  →  I({a}) / I({b})")
            item.setData(Qt.UserRole, (a, b))
            self.list_pairs.addItem(item)

    def _use_pair(self, item: QListWidgetItem):
        a, b = item.data(Qt.UserRole)
        self.spin_peak.setValue(float(a))
        self.spin_peak2.setValue(float(b))
        self.chk_ratio.setChecked(True)
        self.status.setText(f"Ratio chargé : I({a}) / I({b}).")

    def _on_ratio_toggled(self, checked: bool):
        self.spin_peak2.setEnabled(checked)

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
        rows = {idx.row() for idx in self.table.selectedIndexes()}
        paths = [self.table.item(r, 0).data(Qt.UserRole) for r in rows if self.table.item(r, 0)]
        if not paths:
            QMessageBox.information(self, "Rien à retirer", "Sélectionnez d'abord des lignes.")
            return
        for p in paths:
            self.store.remove(p)

    def clear_all(self):
        if not self.store.paths():
            return
        if QMessageBox.question(self, "Tout vider", "Retirer tous les fichiers chargés ?") == QMessageBox.Yes:
            self.store.clear()

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
            missing = self.store.load_session(path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Chargement impossible", str(exc))
            return
        self._sync_titrant_controls()
        msg = f"Tableau chargé : {os.path.basename(path)}."
        if missing:
            msg += f" {len(missing)} introuvable(s) : " + ", ".join(missing[:3]) + ("…" if len(missing) > 3 else "")
        self.status.setText(msg)

    # ------------------------------------------------------------------
    def _emit_meta(self):
        self._emitting_meta = True
        self.store.meta_changed.emit()
        self._emitting_meta = False

    def _on_meta_changed(self):
        if self._emitting_meta:
            return
        self._sync_titrant_controls()
        self._refresh_table()

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
        rows = {idx.row() for idx in self.table.selectedIndexes()}
        if not rows:
            QMessageBox.information(self, "Aucune sélection", "Sélectionnez d'abord des lignes.")
            return
        self._populating = True
        for r in rows:
            path = self.table.item(r, 0).data(Qt.UserRole)
            self.store.series[path] = label
            self.table.item(r, 2).setText(label)
        self._populating = False
        self._emit_meta()
        self.status.setText(f"Série « {label or _NO_SERIES} » affectée à {len(rows)} fichier(s).")

    def _refresh_table(self):
        self._populating = True
        self.table.setRowCount(0)
        for path in self.store.paths():
            row = self.table.rowCount()
            self.table.insertRow(row)
            name_item = QTableWidgetItem(self.store.name(path))
            name_item.setData(Qt.UserRole, path)
            name_item.setToolTip(path)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, name_item)
            vol_item = QTableWidgetItem(self.store.volumes.get(path, ""))
            vol_item.setData(Qt.UserRole, path)
            self.table.setItem(row, 1, vol_item)
            ser_item = QTableWidgetItem(self.store.series.get(path, ""))
            ser_item.setData(Qt.UserRole, path)
            self.table.setItem(row, 2, ser_item)
        self._populating = False

    def _series_label(self, path):
        return self.store.series.get(path, "").strip() or _NO_SERIES

    # ------------------------------------------------------------------
    def _measure(self, path, peak1, peak2, tol, use_ratio, do_baseline):
        x, y = self.store.spectra[path]
        x = np.asarray(x, dtype=float)
        y = np.asarray(y, dtype=float)
        yv = _baseline_corrected(x, y) if do_baseline else y
        i1 = _peak_intensity(x, yv, peak1, tol)
        if use_ratio:
            i2 = _peak_intensity(x, yv, peak2, tol)
            if i2 is None or np.isnan(i2) or i2 == 0:
                return np.nan
            return i1 / i2
        return i1

    def plot_titration(self):
        if not self.store.paths():
            QMessageBox.information(self, "Aucun spectre", "Chargez d'abord des fichiers .txt.")
            return

        peak1 = self.spin_peak.value()
        peak2 = self.spin_peak2.value()
        tol = self.spin_tol.value()
        use_ratio = self.chk_ratio.isChecked()
        do_baseline = self.chk_baseline.isChecked()
        do_fit = self.chk_sigmoid.isChecked()

        amounts = tu.amounts_mol(self.store.paths(), self.store.volumes, self.store.titrant)
        if not amounts:
            QMessageBox.warning(
                self, "Volumes manquants",
                "Renseignez la concentration du titrant et au moins un volume de titrant.",
            )
            return
        unit_label, unit_factor = tu.pick_amount_unit(max(amounts.values()))

        if use_ratio:
            y_title = f"I({peak1:.0f}) / I({peak2:.0f}) (a.u.)"
            auto_title = f"Titration · ratio {peak1:.0f}/{peak2:.0f} cm⁻¹"
            self._last_file_base = f"titration_ratio_{peak1:.0f}_{peak2:.0f}"
        else:
            y_title = f"Intensité au pic {peak1:.0f} cm⁻¹ (a.u.)"
            auto_title = f"Titration · pic {peak1:.0f} cm⁻¹"
            self._last_file_base = f"titration_pic_{peak1:.0f}"
        title = self.edit_title.text().strip() or auto_title

        # série -> liste de points (quantité, valeur, nom), triés par quantité
        series_order = []
        for p in self.store.paths():
            lab = self._series_label(p)
            if lab not in series_order:
                series_order.append(lab)

        fig = go.Figure()
        multi = len(series_order) > 1
        n_pts = n_fits = 0
        eq_points = []
        for si, lab in enumerate(series_order):
            rows = []
            for p in self.store.paths():
                if self._series_label(p) != lab or p not in amounts:
                    continue
                val = self._measure(p, peak1, peak2, tol, use_ratio, do_baseline)
                if not np.isnan(val):
                    rows.append((amounts[p] / unit_factor, val, self.store.name(p)))
            if not rows:
                continue
            rows.sort(key=lambda r: r[0])
            xs = [r[0] for r in rows]
            ys = [r[1] for r in rows]
            names = [r[2] for r in rows]
            n_pts += len(rows)
            color = _PALETTE[si % len(_PALETTE)]
            fig.add_trace(go.Scatter(
                x=xs, y=ys, mode="lines+markers",
                name=(lab if multi else "Mesures"),
                legendgroup=lab, line=dict(width=2, color=color), marker=dict(size=9, color=color),
                text=names, hovertemplate="%{text}<br>x=%{x:.4g}<br>val=%{y:.4g}<extra></extra>",
            ))
            if do_fit:
                popt = tu.fit_sigmoid(xs, ys)
                if popt is not None:
                    xf = np.linspace(min(xs), max(xs), 300)
                    fig.add_trace(go.Scatter(
                        x=xf, y=tu.sigmoid(xf, *popt), mode="lines",
                        name=f"{lab} (sigmoïde)", legendgroup=lab, showlegend=False,
                        line=dict(width=1.6, dash="dash", color=color),
                    ))
                    x_eq = float(popt[3])
                    eq_points.append((lab, x_eq))
                    n_fits += 1
                    if not multi:
                        fig.add_vline(
                            x=x_eq, line=dict(color=color, dash="dot"),
                            annotation_text=f"x_eq ≈ {x_eq:.3g}", annotation_position="top",
                        )

        if n_pts == 0:
            QMessageBox.warning(
                self, "Pas de points",
                "Aucun point exploitable. Vérifiez les volumes saisis et la position du pic.",
            )
            return

        fig.update_layout(
            title=title,
            xaxis_title=f"Quantité de titrant ({unit_label})",
            yaxis_title=y_title,
            template="plotly_white",
            margin=dict(l=80, r=30, t=60, b=60),
            font=dict(size=14),
            legend=dict(title="Séries" if multi else None),
        )
        self.plot_view.show_figure(fig)
        self._last_fig = fig
        self.btn_export.setEnabled(True)

        msg = f"{n_pts} point(s) sur {len(series_order)} série(s)."
        if do_fit:
            if n_fits:
                apercu = ", ".join(f"{lab} : x_eq={xe:.4g}" for lab, xe in eq_points[:3])
                msg += f" {n_fits} sigmoïde(s) — {apercu}"
            else:
                msg += " Sigmoïde : aucun ajustement convergent (≥ 4 points requis)."
        self.status.setText(msg)

    # ------------------------------------------------------------------
    def export_graph(self):
        if self._last_fig is None:
            QMessageBox.information(self, "Rien à exporter", "Tracez d'abord la titration.")
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
                f"Impossible d'exporter en {ext} :\n{exc}\n\nAstuce : l'export HTML fonctionne toujours.",
            )
            return
        self.status.setText(f"Graphique exporté : {os.path.basename(path)}")
