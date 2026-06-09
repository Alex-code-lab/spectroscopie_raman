"""Onglet Titration.

On associe à chaque spectre une concentration (tableau éditable à la main),
on mesure l'intensité d'un pic (ou le ratio de deux pics) par fichier, puis
on trace intensité/ratio en fonction de la concentration — la courbe de
titration. Option : correction de ligne de base et ajustement sigmoïde
(point d'équivalence), comme dans Ramanalyze.
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
)
import plotly.graph_objects as go

from spectrum_loader import load_spectrum
from plot_view import PlotlyView
import peak_presets

_SETTINGS_KEY = "simple/last_dir"


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


def _sigmoid(x, A, B, k, x_eq):
    return A + B / (1 + np.exp(-k * (x - x_eq)))


class TitrationTab(QWidget):
    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store

        # ---------------- Panneau gauche : tableau + contrôles ----------------
        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        left_layout.addWidget(QLabel("<b>Concentrations par échantillon</b>", self))
        hint = QLabel(
            "Double-cliquez dans la colonne <i>Concentration</i> pour saisir la "
            "valeur de chaque échantillon.", self
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #888;")
        left_layout.addWidget(hint)

        self.table = QTableWidget(0, 2, self)
        self.table.setMinimumHeight(180)
        self.table.setHorizontalHeaderLabels(["Fichier", "Concentration"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(self._on_cell_changed)
        left_layout.addWidget(self.table, 1)

        file_btns = QHBoxLayout()
        self.btn_open = QPushButton("📂  Ajouter des fichiers…", self)
        self.btn_open.clicked.connect(self.open_files)
        self.btn_import = QPushButton("Importer CSV…", self)
        self.btn_import.setToolTip("CSV à 2 colonnes : nom de fichier ; concentration")
        self.btn_import.clicked.connect(self.import_csv)
        self.btn_export = QPushButton("Exporter CSV…", self)
        self.btn_export.clicked.connect(self.export_csv)
        file_btns.addWidget(self.btn_open)
        file_btns.addWidget(self.btn_import)
        file_btns.addWidget(self.btn_export)
        left_layout.addLayout(file_btns)

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
        src_layout.addWidget(QLabel(
            "Double-cliquez une paire pour la charger comme ratio (Pic 1 / Pic 2) :", self))
        self.list_pairs = QListWidget(self)
        self.list_pairs.setMaximumHeight(130)
        self.list_pairs.itemDoubleClicked.connect(self._use_pair)
        src_layout.addWidget(self.list_pairs)
        left_layout.addWidget(src_box)

        # --- Paramètres du / des pic(s) ---
        params = QGroupBox("Pic à mesurer", self)
        form = QFormLayout(params)

        self.spin_peak = QDoubleSpinBox(self)
        self.spin_peak.setRange(0.0, 5000.0)
        self.spin_peak.setDecimals(1)
        self.spin_peak.setSingleStep(1.0)
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
        self.spin_peak2.setSingleStep(1.0)
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

        # --- Titre du graphique (éditable) ---
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

        self._last_fig = None
        self._last_file_base = "titration"

        self.status = QLabel("", self)
        self.status.setStyleSheet("color: #888;")
        self.status.setWordWrap(True)
        left_layout.addWidget(self.status)

        # Panneau gauche défilant : barre de défilement verticale si le contenu
        # dépasse la hauteur de la fenêtre.
        left_scroll = QScrollArea(self)
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        left_scroll.setMinimumWidth(440)

        # ---------------- Panneau droit : graphe ----------------
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

        self.store.changed.connect(self.sync_from_store)
        self.sync_from_store()
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

    def _on_ratio_toggled(self, checked: bool):
        self.spin_peak2.setEnabled(checked)

    # ------------------------------------------------------------------
    def _on_cell_changed(self, item: QTableWidgetItem):
        if item.column() != 1:
            return
        path = item.data(Qt.UserRole)
        if path is not None:
            self.store.concentrations[path] = item.text().strip()

    def sync_from_store(self):
        """Reconstruit le tableau depuis le store (conserve les concentrations)."""
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        for path in self.store.paths():
            row = self.table.rowCount()
            self.table.insertRow(row)

            name_item = QTableWidgetItem(self.store.name(path))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            name_item.setData(Qt.UserRole, path)
            name_item.setToolTip(path)
            self.table.setItem(row, 0, name_item)

            conc_item = QTableWidgetItem(self.store.concentrations.get(path, ""))
            conc_item.setData(Qt.UserRole, path)
            self.table.setItem(row, 1, conc_item)
        self.table.blockSignals(False)

    # ------------------------------------------------------------------
    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Importer les concentrations", self._last_dir(),
            "CSV (*.csv *.txt);;Tous les fichiers (*)",
        )
        if not path:
            return
        # Index des fichiers connus par nom de base (sans extension aussi).
        by_name = {}
        for p in self.store.paths():
            base = self.store.name(p)
            by_name[base] = p
            by_name[os.path.splitext(base)[0]] = p

        n = 0
        try:
            with open(path, "r", encoding="utf-8", newline="") as f:
                sample = f.read(2048)
                f.seek(0)
                delim = ";" if sample.count(";") >= sample.count(",") else ","
                for fields in csv.reader(f, delimiter=delim):
                    if len(fields) < 2:
                        continue
                    key = fields[0].strip()
                    val = fields[1].strip()
                    target = by_name.get(key) or by_name.get(os.path.splitext(key)[0])
                    if target is not None:
                        self.store.concentrations[target] = val
                        n += 1
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Import impossible", str(exc))
            return
        self.sync_from_store()
        self.status.setText(f"{n} concentration(s) importée(s).")

    def export_csv(self):
        if not self.store.paths():
            QMessageBox.information(self, "Rien à exporter", "Aucun spectre chargé.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter les concentrations", os.path.join(self._last_dir(), "concentrations.csv"),
            "CSV (*.csv)",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Fichier", "Concentration"])
                for p in self.store.paths():
                    w.writerow([self.store.name(p), self.store.concentrations.get(p, "")])
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Export impossible", str(exc))
            return
        self.status.setText(f"Exporté : {os.path.basename(path)}")

    # ------------------------------------------------------------------
    @staticmethod
    def _parse_conc(text: str):
        if text is None:
            return np.nan
        t = str(text).strip().replace(",", ".")
        if not t:
            return np.nan
        try:
            return float(t)
        except ValueError:
            return np.nan

    def plot_titration(self):
        if not self.store.paths():
            QMessageBox.information(self, "Aucun spectre", "Chargez d'abord des fichiers .txt.")
            return

        peak1 = self.spin_peak.value()
        peak2 = self.spin_peak2.value()
        tol = self.spin_tol.value()
        use_ratio = self.chk_ratio.isChecked()
        do_baseline = self.chk_baseline.isChecked()

        rows = []  # (conc, value, name)
        missing_conc = 0
        for path in self.store.paths():
            conc = self._parse_conc(self.store.concentrations.get(path, ""))
            if np.isnan(conc):
                missing_conc += 1
                continue
            x, y = self.store.spectra[path]
            yv = _baseline_corrected(x, y) if do_baseline else y
            i1 = _peak_intensity(x, yv, peak1, tol)
            if use_ratio:
                i2 = _peak_intensity(x, yv, peak2, tol)
                value = i1 / i2 if (i2 not in (0, np.nan) and not np.isnan(i2) and i2 != 0) else np.nan
            else:
                value = i1
            if not np.isnan(value):
                rows.append((conc, value, self.store.name(path)))

        if not rows:
            QMessageBox.warning(
                self, "Pas de points",
                "Aucun point exploitable. Vérifiez les concentrations saisies et la "
                "position du pic (aucun signal trouvé dans la fenêtre de tolérance ?).",
            )
            return

        rows.sort(key=lambda r: r[0])
        xs = np.array([r[0] for r in rows], dtype=float)
        ys = np.array([r[1] for r in rows], dtype=float)
        names = [r[2] for r in rows]

        if use_ratio:
            y_title = f"I({peak1:.0f}) / I({peak2:.0f}) (a.u.)"
            auto_title = f"Titration · ratio pics {peak1:.0f}/{peak2:.0f} cm⁻¹"
            self._last_file_base = f"titration_ratio_{peak1:.0f}_{peak2:.0f}"
        else:
            y_title = f"Intensité au pic {peak1:.0f} cm⁻¹ (a.u.)"
            auto_title = f"Titration · pic {peak1:.0f} cm⁻¹"
            self._last_file_base = f"titration_pic_{peak1:.0f}"
        # Titre éditable : on respecte la saisie de l'utilisateur si présente.
        title = self.edit_title.text().strip() or auto_title

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines+markers", name="Mesures",
            text=names, hovertemplate="%{text}<br>conc=%{x}<br>val=%{y:.3g}<extra></extra>",
            marker=dict(size=9), line=dict(width=2),
        ))

        fit_msg = ""
        if self.chk_sigmoid.isChecked():
            fit_msg = self._add_sigmoid_fit(fig, xs, ys)

        fig.update_layout(
            title=title,
            xaxis_title="Concentration",
            yaxis_title=y_title,
            template="plotly_white",
            margin=dict(l=80, r=30, t=60, b=60),
            font=dict(size=14),
        )
        self.plot_view.show_figure(fig)
        self._last_fig = fig
        self.btn_export.setEnabled(True)

        msg = f"{len(rows)} point(s) tracé(s)."
        if missing_conc:
            msg += f" {missing_conc} échantillon(s) sans concentration ignoré(s)."
        if fit_msg:
            msg += " " + fit_msg
        self.status.setText(msg)

    def _add_sigmoid_fit(self, fig, xs, ys) -> str:
        """Ajoute la sigmoïde ajustée et marque le point d'équivalence. Renvoie un message."""
        if len(xs) < 4:
            return "Sigmoïde : pas assez de points (≥ 4 requis)."
        try:
            from scipy.optimize import curve_fit
        except Exception:
            return "Sigmoïde indisponible (scipy manquant)."

        x_span = float(np.ptp(xs)) or 1.0
        y_min, y_max = float(np.min(ys)), float(np.max(ys))
        p0 = [y_min, y_max - y_min, 1.0 / x_span, float(np.median(xs))]
        try:
            popt, _ = curve_fit(_sigmoid, xs, ys, p0=p0, maxfev=20000)
        except Exception as exc:  # noqa: BLE001
            return f"Ajustement sigmoïde échoué ({exc})."

        x_eq = popt[3]
        x_fit = np.linspace(float(np.min(xs)), float(np.max(xs)), 300)
        fig.add_trace(go.Scatter(
            x=x_fit, y=_sigmoid(x_fit, *popt), mode="lines",
            name="Sigmoïde", line=dict(width=2, dash="dash", color="#d9534f"),
        ))
        if np.min(xs) <= x_eq <= np.max(xs):
            fig.add_vline(
                x=x_eq, line=dict(color="#d9534f", dash="dot"),
                annotation_text=f"x_eq ≈ {x_eq:.3g}", annotation_position="top",
            )
        return f"Point d'équivalence ≈ {x_eq:.4g}."

    # ------------------------------------------------------------------
    def export_graph(self):
        if self._last_fig is None:
            QMessageBox.information(self, "Rien à exporter", "Tracez d'abord la titration.")
            return
        default = os.path.join(self._last_dir(), self._last_file_base + ".png")
        path, selected = QFileDialog.getSaveFileName(
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
            else:  # .png ou .pdf via kaleido
                self._last_fig.write_image(path, width=1200, height=800, scale=2)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, "Export impossible",
                f"Impossible d'exporter en {ext} :\n{exc}\n\n"
                "Astuce : l'export HTML fonctionne toujours.",
            )
            return
        self.status.setText(f"Graphique exporté : {os.path.basename(path)}")
