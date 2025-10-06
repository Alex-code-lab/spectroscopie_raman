from __future__ import annotations

import os
import itertools
from typing import Optional, List

import numpy as np
import pandas as pd
import plotly.express as px
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QMessageBox, QTableWidget, QTableWidgetItem, QSpacerItem, QSizePolicy,
    QComboBox, QDoubleSpinBox, QHeaderView
)
from PySide6.QtWebEngineWidgets import QWebEngineView


class AnalysisTab(QWidget):
    """
    Onglet Analyse – utilise le fichier combiné créé dans l'onglet Métadonnées
    pour extraire les intensités au voisinage de pics d'intérêt et calculer
    les rapports d'intensité.

    ➜ Pas de recombinaison ici. On consomme `metadata_picker.combined_df`.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._combined_df: Optional[pd.DataFrame] = None
        self._peaks: List[int] = []
        self._peak_intensities: Optional[pd.DataFrame] = None
        self._ratios_long: Optional[pd.DataFrame] = None

        # --- UI ---
        layout = QVBoxLayout(self)

        header = QLabel("<b>Analyse de pics sur le fichier combiné</b><br>\n"
                        "Cette page utilise le fichier combiné (txt + métadonnées) validé dans l’onglet Métadonnées.")
        header.setWordWrap(True)
        layout.addWidget(header)

        # état (fichier combiné présent / absent)
        self.lbl_status = QLabel("Fichier combiné : non chargé")
        layout.addWidget(self.lbl_status)

        # Options d'analyse (présets de pics + tolérance)
        opts = QHBoxLayout()
        opts.addWidget(QLabel("Jeu de pics :"))
        self.cmb_presets = QComboBox(self)
        self.cmb_presets.addItems(["532 nm (1231,1327,1342,1358,1450)", "785 nm (412,444,471,547,1561)"])
        self.cmb_presets.currentIndexChanged.connect(self._on_preset_changed)
        opts.addWidget(self.cmb_presets)

        opts.addSpacerItem(QSpacerItem(20, 10, QSizePolicy.Expanding, QSizePolicy.Minimum))
        opts.addWidget(QLabel("Tolérance (cm⁻¹) :"))
        self.spin_tol = QDoubleSpinBox(self)
        self.spin_tol.setDecimals(2)
        self.spin_tol.setRange(0.0, 50.0)
        self.spin_tol.setSingleStep(0.1)
        self.spin_tol.setValue(2.0)
        opts.addWidget(self.spin_tol)
        layout.addLayout(opts)

        # Boutons d'action
        actions = QHBoxLayout()
        self.btn_refresh = QPushButton("Recharger le fichier combiné depuis Métadonnées", self)
        self.btn_analyze = QPushButton("Analyser les pics", self)
        self.btn_export = QPushButton("Exporter résultats (Excel)…", self)
        self.btn_export.setEnabled(False)
        self.btn_refresh.clicked.connect(self._reload_combined)
        self.btn_analyze.clicked.connect(self._run_analysis)
        self.btn_export.clicked.connect(self._export_results)
        actions.addWidget(self.btn_refresh)
        actions.addWidget(self.btn_analyze)
        actions.addWidget(self.btn_export)
        layout.addLayout(actions)

        # Aperçu des résultats (table des intensités de pics)
        layout.addWidget(QLabel("<b>Aperçu – Intensités extraites</b>"))
        self.table = QTableWidget(self)
        # Configuration pour navigation sur tout le fichier
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(self.table, 0)

        # Graphique des ratios
        layout.addWidget(QLabel("<b>Graphique – Rapports d’intensité Raman selon [EGTA]</b>"))
        self.plot_view = QWebEngineView(self)
        layout.addWidget(self.plot_view, 1)

        # Init valeurs
        self._on_preset_changed(self.cmb_presets.currentIndex())
        self._reload_combined()  # tente de charger au démarrage si déjà assemblé

    # ---------- Helpers ----------
    def _on_preset_changed(self, idx: int):
        if idx == 0:
            self._peaks = [1231, 1327, 1342, 1358, 1450]  # 532 nm
        else:
            self._peaks = [412, 444, 471, 547, 1561]  # 785 nm

    def _reload_combined(self):
        # Utiliser la fenêtre principale (MainWindow), pas le parent immédiat (TabWidget)
        main = self.window()
        if main is None:
            QMessageBox.critical(self, "Erreur", "Fenêtre principale introuvable.")
            return
        metadata_picker = getattr(main, 'metadata_picker', None)
        df = getattr(metadata_picker, 'combined_df', None) if metadata_picker is not None else None
        if isinstance(df, pd.DataFrame) and not df.empty:
            self._combined_df = df
            self.lbl_status.setText("Fichier combiné : chargé ✓")
        else:
            self._combined_df = None
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — assemblez et validez dans l’onglet Métadonnées")

    def _set_preview(self, df: pd.DataFrame, n: int = 20):  # n conservé pour compat, ignoré ici
        # Afficher TOUT le DataFrame avec défilement et tri
        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                self.table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))
        self.table.setSortingEnabled(True)

    # ---------- Analyse ----------
    def _run_analysis(self):
        if self._combined_df is None or self._combined_df.empty:
            QMessageBox.warning(self, "Données manquantes", "Aucun fichier combiné chargé. Allez dans l’onglet Métadonnées, assemblez puis validez le choix.")
            return

        df = self._combined_df
        required = {"Raman Shift", "Intensity_corrected", "file"}
        if not required.issubset(df.columns):
            QMessageBox.critical(self, "Colonnes manquantes", "Le fichier combiné doit contenir : Raman Shift, Intensity_corrected, file.")
            return

        peaks = self._peaks
        tol = float(self.spin_tol.value())

        # Calcul des intensités max autour de chaque pic, par fichier
        results: List[dict] = []
        for fname, group in df.groupby("file"):
            spectrum = group.sort_values("Raman Shift")
            record = {"file": fname}
            for target in peaks:
                window = spectrum[(spectrum["Raman Shift"] >= target - tol) & (spectrum["Raman Shift"] <= target + tol)]
                if not window.empty:
                    record[f"I_{target}"] = float(window["Intensity_corrected"].max())
                else:
                    record[f"I_{target}"] = np.nan
            results.append(record)

        peak_intensities = pd.DataFrame(results)

        # Calcul des ratios toutes paires (combinations)
        for (a, b) in itertools.combinations(peaks, 2):
            peak_intensities[f"ratio_I_{a}_I_{b}"] = peak_intensities[f"I_{a}"] / peak_intensities[f"I_{b}"]

        # Joindre aux métadonnées (si colonnes présentes)
        if "Spectrum name" not in peak_intensities.columns:
            peak_intensities["Spectrum name"] = peak_intensities["file"].str.replace(".txt", "", regex=False)
        meta_cols = [c for c in df.columns if c not in {"Raman Shift", "Intensity_corrected"}]
        metadata_df = df[meta_cols].drop_duplicates(subset=["Spectrum name"], keep="first") if "Spectrum name" in df.columns else None
        merged = peak_intensities.merge(metadata_df, on="Spectrum name", how="left") if metadata_df is not None else peak_intensities.copy()

        # Mise en forme long pour les ratios
        ratio_cols = [c for c in merged.columns if c.startswith("ratio_I_")]
        id_vars = [c for c in ["file", "Spectrum name", "Sample description", "n(EGTA) (mol)"] if c in merged.columns]
        df_ratios = merged.melt(id_vars=id_vars, value_vars=ratio_cols, var_name="Ratio", value_name="Value") if ratio_cols else None

        # Mémoriser + aperçu
        self._peak_intensities = merged
        self._ratios_long = df_ratios
        self._set_preview(merged)
        self.btn_export.setEnabled(True)

        # --- Tracé interactif (équivalent au ggplot fourni) ---
        if df_ratios is not None and not df_ratios.empty and "n(EGTA) (mol)" in df_ratios.columns:
            fig = px.line(
                df_ratios.sort_values(["Ratio", "n(EGTA) (mol)"]),
                x="n(EGTA) (mol)",
                y="Value",
                color="Ratio",
                markers=True,
                title="Rapports d’intensité Raman selon [EGTA]",
            )
            fig.update_layout(
                xaxis_title="Quantité EGTA (mol)",
                yaxis_title="Rapport d’intensité (a.u.)",
                width=1200,
                height=600,
                legend_title_text="Ratio",
            )
            # Formatage type scientifique pour l'axe X
            fig.update_xaxes(tickformat=".0e")
            self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))
        else:
            # Effacer le graphique si pas de ratios
            self.plot_view.setHtml("<i>Aucun ratio à afficher.</i>")

        QMessageBox.information(self, "Succès", "Analyse terminée : intensités et ratios calculés.")

    # ---------- Export ----------
    def _export_results(self):
        if self._peak_intensities is None or self._peak_intensities.empty:
            QMessageBox.warning(self, "Rien à exporter", "Lancez d'abord l'analyse.")
            return
        import datetime as _dt
        home = os.path.expanduser('~')
        stamp = _dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        out_path = os.path.join(home, f"raman_analysis_{stamp}.xlsx")
        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                self._peak_intensities.to_excel(writer, index=False, sheet_name="intensites")
                if self._ratios_long is not None:
                    self._ratios_long.to_excel(writer, index=False, sheet_name="ratios")
            QMessageBox.information(self, "Exporté", f"Résultats enregistrés :\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Échec export", f"Impossible d'enregistrer : {e}")