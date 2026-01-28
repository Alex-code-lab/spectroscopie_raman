from __future__ import annotations

import os
import re
import itertools
from typing import Optional, List

import numpy as np
import pandas as pd
import plotly.express as px
from scipy.optimize import curve_fit
import plotly.graph_objects as go
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QMessageBox, QTableWidget, QTableWidgetItem, QSpacerItem, QSizePolicy,
    QComboBox, QDoubleSpinBox, QHeaderView, QLineEdit
)
from PySide6.QtWebEngineWidgets import QWebEngineView

from data_processing import build_combined_dataframe_from_ui

class AnalysisTab(QWidget):
    """
    Onglet Analyse – utilise le fichier combiné créé dans l'onglet Métadonnées
    pour extraire les intensités au voisinage de pics d'intérêt et calculer
    les rapports d'intensité.

    ➜ Ici, on reconstruit le fichier combiné à partir des spectres .txt sélectionnés
    et des tableaux créés dans `MetadataCreatorWidget` (onglet Métadonnées).
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._combined_df: Optional[pd.DataFrame] = None
        self._peaks: List[int] = []
        self._peak_intensities: Optional[pd.DataFrame] = None
        self._ratios_long: Optional[pd.DataFrame] = None
        # Indique que l'analyse doit être (re)lancée (données ou paramètres modifiés)
        self._analysis_dirty: bool = True
        self._sigmoid_params = None   # (A, B, k, x_eq)
        self._sigmoid_cov = None
        self._sigmoid_results = {}  # ratio_name -> {"params": popt, "cov": pcov}

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
        self.cmb_presets.addItems([
            "532 nm (1231,1327,1342,1358,1450)",
            "785 nm (412,444,471,547,1561)",
            "Personnalisé"
        ])
        self.cmb_presets.currentIndexChanged.connect(self._on_preset_changed)
        opts.addWidget(self.cmb_presets)

        # Paramètres pour mode personnalisé
        opts.addWidget(QLabel("Laser (nm) :"))
        self.spin_laser_nm = QDoubleSpinBox(self)
        self.spin_laser_nm.setRange(200.0, 2000.0)
        self.spin_laser_nm.setDecimals(2)
        self.spin_laser_nm.setSingleStep(1.0)
        self.spin_laser_nm.setValue(532.0)
        self.spin_laser_nm.setEnabled(False)
        opts.addWidget(self.spin_laser_nm)

        opts.addWidget(QLabel("Pics à étudier (Raman Shift en cm⁻¹, séparés par des virgules) :"))
        self.edit_custom_waves = QLineEdit(self)
        self.edit_custom_waves.setPlaceholderText("ex : 1231, 1327, 1450")
        self.edit_custom_waves.setEnabled(False)
        opts.addWidget(self.edit_custom_waves)

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
        # Tout changement de paramètres d'analyse => il faudra relancer l'analyse
        self.spin_tol.valueChanged.connect(self._on_analysis_param_changed)
        self.cmb_presets.currentIndexChanged.connect(self._on_analysis_param_changed)
        if hasattr(self, "spin_laser_nm"):
            self.spin_laser_nm.valueChanged.connect(self._on_analysis_param_changed)
        if hasattr(self, "edit_custom_waves"):
            self.edit_custom_waves.textChanged.connect(self._on_analysis_param_changed)
        actions.addWidget(self.btn_refresh)
        actions.addWidget(self.btn_analyze)
        actions.addWidget(self.btn_export)
        layout.addLayout(actions)

        # Bouton pour afficher / masquer le tableau des intensités
        self.btn_toggle_table = QPushButton("Afficher le tableau des intensités", self)
        self.btn_toggle_table.clicked.connect(self._toggle_table_visibility)
        layout.addWidget(self.btn_toggle_table)

        # Aperçu des résultats (table des intensités de pics)
        self.lbl_table_title = QLabel("<b>Aperçu – Intensités extraites</b>")
        layout.addWidget(self.lbl_table_title)
        self.table = QTableWidget(self)
        # Configuration pour navigation sur tout le fichier
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(self.table, 0)

        # Masquer la table et son titre par défaut pour laisser plus de place au graphique
        self.lbl_table_title.setVisible(False)
        self.table.setVisible(False)

        # Graphique des ratios
        layout.addWidget(QLabel("<b>Graphique – Rapports d’intensité Raman selon [titrant]</b>"))
        self.plot_view = QWebEngineView(self)
        layout.addWidget(self.plot_view, 1)

        # --- Fit sigmoïde (équivalence) ---
        fit_layout = QHBoxLayout()

        fit_layout.addWidget(QLabel("Ratio à ajuster :"))
        self.cmb_fit_ratio = QComboBox(self)
        fit_layout.addWidget(self.cmb_fit_ratio)

        self.btn_fit_sigmoid = QPushButton("Ajuster par sigmoïde", self)
        self.btn_fit_sigmoid.clicked.connect(self._fit_sigmoid)
        fit_layout.addWidget(self.btn_fit_sigmoid)

        self.btn_show_equivalence = QPushButton("Afficher l'équivalence", self)
        self.btn_show_equivalence.clicked.connect(self._show_equivalence)
        self.btn_show_equivalence.setEnabled(False)
        fit_layout.addWidget(self.btn_show_equivalence)

        layout.addLayout(fit_layout)

        # Init valeurs
        self._on_preset_changed(self.cmb_presets.currentIndex())
        # Figure Plotly actuellement affichée (pour overlay fit)
        self._current_fig = None   # figure Plotly actuellement affichée
        # Couleurs initiales des boutons
        self._refresh_button_states()

    @staticmethod
    def _sigmoid(x, A, B, k, x_eq):
        return A + B / (1 + np.exp(-k * (x - x_eq)))

    def _fit_sigmoid(self):
        if self._ratios_long is None or self._ratios_long.empty:
            QMessageBox.warning(self, "Aucune donnée", "Aucun ratio disponible pour le fit.")
            return

        if self._current_fig is None:
            QMessageBox.warning(self, "Graphique manquant", "Aucun graphique n'est actuellement affiché.")
            return

        ratio_name = self.cmb_fit_ratio.currentText()
        if not ratio_name:
            QMessageBox.warning(self, "Ratio manquant", "Veuillez sélectionner un ratio à ajuster.")
            return

        # ========== 1) Sécuriser la sélection des données (tri + copie) ==========
        df = (
            self._ratios_long[self._ratios_long["Ratio"] == ratio_name]
            .dropna(subset=["n(titrant) (mol)", "Value"])
            .sort_values("n(titrant) (mol)")
            .copy()
        )

        if len(df) < 5:
            QMessageBox.warning(self, "Données insuffisantes", "Pas assez de points pour un fit sigmoïde.")
            return

        x = df["n(titrant) (mol)"].to_numpy(dtype=float)
        y = df["Value"].to_numpy(dtype=float)
        x_min = float(np.nanmin(x))
        x_max = float(np.nanmax(x))
        x_span = x_max - x_min
        if (not np.isfinite(x_span)) or x_span <= 0:
            x_span = max(1.0, abs(x_min) if np.isfinite(x_min) else 1.0)
        x_center = 0.5 * (x_min + x_max)
        x_scale = x_span if x_span > 0 else 1.0
        y_min = float(np.nanmin(y))
        y_max = float(np.nanmax(y))
        y_span = y_max - y_min
        if (not np.isfinite(y_span)) or y_span <= 0:
            QMessageBox.warning(
                self,
                "Données insuffisantes",
                "Les valeurs du ratio sont quasi constantes, le fit sigmoïde est indéterminé.",
            )
            return

        fig = self._current_fig

        # ========== 2) Supprimer les anciens fits / points du même ratio ==========
        # Supprimer les anciens fits / annotations pour ce ratio
        fig.data = tuple(
            tr for tr in fig.data
            if not (
                isinstance(tr.name, str)
                and (f"Fit sigmoïde – {ratio_name}" in tr.name
                     or f"Points utilisés – {ratio_name}" in tr.name
                     or f"Équivalence – {ratio_name}" in tr.name)
            )
        )
        # Supprimer les annotations liées à ce ratio
        if hasattr(fig.layout, "annotations") and fig.layout.annotations:
            fig.layout.annotations = tuple(
                ann for ann in fig.layout.annotations
                if ratio_name not in str(getattr(ann, "text", ""))
            )

        # ========== 3) Afficher explicitement les points utilisés pour le fit ==========
        # Affichage explicite des points utilisés pour le fit (debug visuel)
        fig.add_trace(go.Scatter(
            x=x,
            y=y,
            mode="markers",
            marker=dict(size=10, symbol="x"),
            name=f"Points utilisés – {ratio_name}",
        ))

        # ========== 4) (Option sécurité) Forcer un refit réel même si données identiques ==========
        # Sécurité : conversion explicite en float64 pour éviter tout cache implicite
        x = np.asarray(x, dtype=np.float64)
        y = np.asarray(y, dtype=np.float64)

        # Estimations initiales (espace normalisé pour stabiliser le fit)
        A0 = y_min
        B0 = y_span
        corr = np.corrcoef(x, y)[0, 1] if len(x) > 1 else 1.0
        sign = -1.0 if (np.isfinite(corr) and corr < 0) else 1.0
        k0 = sign * 8.0
        y_mid = A0 + 0.5 * B0
        try:
            idx_mid = int(np.nanargmin(np.abs(y - y_mid)))
            xeq0 = float((x[idx_mid] - x_center) / x_scale)
        except Exception:
            xeq0 = 0.0

        def _sigmoid_scaled(x_scaled, A, B, k, x_eq_scaled):
            return A + B / (1 + np.exp(-k * (x_scaled - x_eq_scaled)))

        try:
            popt_scaled, pcov_scaled = curve_fit(
                _sigmoid_scaled,
                (x - x_center) / x_scale,
                y,
                p0=[A0, B0, k0, xeq0],
                maxfev=20000,
            )
        except Exception as e:
            QMessageBox.critical(self, "Échec du fit", f"Impossible d'ajuster la sigmoïde :\n{e}")
            return

        k_orig = float(popt_scaled[2]) / x_scale
        x_eq_orig = x_center + x_scale * float(popt_scaled[3])
        popt = np.array([popt_scaled[0], popt_scaled[1], k_orig, x_eq_orig], dtype=float)
        if pcov_scaled is not None:
            scale = np.diag([1.0, 1.0, 1.0 / x_scale, x_scale])
            pcov = scale @ pcov_scaled @ scale
        else:
            pcov = None

        self._sigmoid_params = popt
        self._sigmoid_cov = pcov
        self.btn_show_equivalence.setEnabled(True)
        # Enregistrer le résultat du fit par ratio
        self._sigmoid_results[ratio_name] = {"params": popt, "cov": pcov}

        # -------- Courbe de fit avec plage étendue pour afficher toute la sigmoïde --------
        # Point d’équivalence (point d’inflexion) : coordonnées exactes
        x_eq = float(popt[3])
        y_eq = self._sigmoid(x_eq, *popt)
        k = float(popt[2])
        if np.isfinite(k) and k != 0:
            half_span_k = 6.0 / abs(k)  # ~de 0.25% à 99.75%
        else:
            half_span_k = 0.0
        half_span = max(x_span, half_span_k)
        if (not np.isfinite(half_span)) or half_span <= 0:
            half_span = max(1.0, abs(x_eq) if np.isfinite(x_eq) else 1.0)

        x_fit_min = min(x_min, x_eq - half_span)
        x_fit_max = max(x_max, x_eq + half_span)
        x_fit = np.linspace(x_fit_min, x_fit_max, 1000)
        y_fit = self._sigmoid(x_fit, *popt)

        # -------- Ajout du point d'équivalence (croix) --------

        # Courbe de fit
        fig.add_trace(go.Scatter(
            x=x_fit,
            y=y_fit,
            mode="lines",
            name=f"Fit sigmoïde – {ratio_name}",
            line=dict(dash="dash", width=3)
        ))

        # Marqueur du point d’équivalence (croix)
        fig.add_trace(go.Scatter(
            x=[x_eq],
            y=[y_eq],
            mode="markers",
            name=f"Équivalence – {ratio_name}",
            marker=dict(
                symbol="x",
                size=14,
                color="red",
                line=dict(width=3)
            )
        ))

        # Ligne verticale équivalence (sans annotation automatique)
        fig.add_vline(
            x=x_eq,
            line_dash="dot",
            line_color="red",
        )

        # Annotation lisible empilée verticalement (coordonnées papier)
        n_annot = len(self._sigmoid_results)  # 1, 2, 3...
        y_paper = 1.0 - 0.06 * (n_annot - 1)
        if y_paper < 0.05:
            y_paper = 0.05

        fig.add_annotation(
            x=x_eq,
            y=y_paper,
            xref="x",
            yref="paper",
            text=f"Équiv. {ratio_name} : {x_eq:.3e}",
            showarrow=False,
            xanchor="left",
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="rgba(0,0,0,0.25)",
            borderwidth=1,
        )

        self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))

    def _show_equivalence(self):
        if not getattr(self, "_sigmoid_results", None):
            QMessageBox.information(
                self,
                "Aucune équivalence",
                "Aucun fit sigmoïde n’a encore été effectué."
            )
            return

        lines = ["Points d’équivalence (points d’inflexion) :\n"]
        for ratio_name, res in self._sigmoid_results.items():
            popt = res.get("params")
            pcov = res.get("cov")
            if popt is None:
                continue

            x_eq = float(popt[3])

            if pcov is not None:
                try:
                    err = np.sqrt(np.diag(pcov))
                    xeq_err = float(err[3])
                    lines.append(
                        f"• {ratio_name} : x_eq = {x_eq:.3e} ± {xeq_err:.1e} mol"
                    )
                except Exception:
                    lines.append(
                        f"• {ratio_name} : x_eq = {x_eq:.3e} mol"
                    )
            else:
                lines.append(
                    f"• {ratio_name} : x_eq = {x_eq:.3e} mol"
                )

        QMessageBox.information(self, "Équivalences", "\n".join(lines))
    def _refresh_button_states(self) -> None:
        """Met à jour la couleur des boutons selon l'état du fichier combiné et de l'analyse."""
        red = "background-color: #d9534f; color: white; font-weight: 600;"
        green = "background-color: #5cb85c; color: white; font-weight: 600;"

        combined_ready = (
            self._combined_df is not None
            and isinstance(self._combined_df, pd.DataFrame)
            and not self._combined_df.empty
        )

        analysis_ready = (
            combined_ready
            and (self._analysis_dirty is False)
            and (self._ratios_long is not None)
            and isinstance(self._ratios_long, pd.DataFrame)
            and (not self._ratios_long.empty)
        )

        # Bouton reload : vert si le fichier combiné est chargé
        self.btn_refresh.setStyleSheet(green if combined_ready else red)
        self.btn_refresh.setToolTip(
            "Vert = fichier combiné chargé" if combined_ready else "Rouge = recharger le fichier combiné"
        )

        # Bouton analyse : vert seulement si l'analyse a été faite et est à jour
        self.btn_analyze.setStyleSheet(green if analysis_ready else red)
        self.btn_analyze.setToolTip(
            "Vert = analyse à jour"
            if analysis_ready
            else (
                "Rouge = lancez (ou relancez) l'analyse" if combined_ready else "Rouge = chargez d'abord le fichier combiné"
            )
        )

    def _on_analysis_param_changed(self, *args) -> None:
        """Marque l'analyse comme à relancer si un paramètre change."""
        self._analysis_dirty = True
        self._refresh_button_states()
        # Ne pas recharger automatiquement le fichier combiné au démarrage pour éviter
        # les popups d'avertissement si aucun .txt/metadonnée n'est encore défini.
        # L'utilisateur cliquera sur "Recharger le fichier combiné depuis Métadonnées"
        # quand il sera prêt.
        # self._reload_combined()

    # ---------- Helpers ----------
    def _on_preset_changed(self, idx: int):
        """
        Met à jour la liste des pics en fonction du preset sélectionné.

        idx = 0 : preset 532 nm
        idx = 1 : preset 785 nm
        idx = 2 : mode personnalisé (défini par l'utilisateur)
        """
        if idx == 0:
            # 532 nm – jeu de pics par défaut
            self._peaks = [1231, 1327, 1342, 1358, 1450]
            # Désactiver les champs personnalisés
            if hasattr(self, "spin_laser_nm"):
                self.spin_laser_nm.setEnabled(False)
            if hasattr(self, "edit_custom_waves"):
                self.edit_custom_waves.setEnabled(False)
        elif idx == 1:
            # 785 nm – jeu de pics par défaut
            self._peaks = [412, 444, 471, 547, 1561]
            # Désactiver les champs personnalisés
            if hasattr(self, "spin_laser_nm"):
                self.spin_laser_nm.setEnabled(False)
            if hasattr(self, "edit_custom_waves"):
                self.edit_custom_waves.setEnabled(False)
        else:
            # Mode personnalisé : on laisse _peaks vide ici,
            # il sera calculé dans _run_analysis à partir des λ fournies.
            self._peaks = []
            # Activer les champs personnalisés
            if hasattr(self, "spin_laser_nm"):
                self.spin_laser_nm.setEnabled(True)
            if hasattr(self, "edit_custom_waves"):
                self.edit_custom_waves.setEnabled(True)

    def _reload_combined(self):
        """Reconstruit le DataFrame combiné à partir des métadonnées créées dans l'onglet Métadonnées
        et des fichiers .txt sélectionnés dans l'onglet Fichiers.
        """
        # Utiliser la fenêtre principale (MainWindow), pas le parent immédiat (TabWidget)
        main = self.window()
        if main is None:
            QMessageBox.critical(self, "Erreur", "Fenêtre principale introuvable.")
            return

        # Récupérer le créateur de métadonnées
        metadata_creator = getattr(main, "metadata_creator", None)
        if metadata_creator is None or not hasattr(metadata_creator, "build_merged_metadata"):
            self._combined_df = None
            self._analysis_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — définissez d'abord les métadonnées dans l'onglet Métadonnées")
            QMessageBox.warning(
                self,
                "Métadonnées manquantes",
                "L'onglet Métadonnées n'est pas correctement initialisé ou ne fournit pas encore les tableaux.\n"
                "Créez d'abord les tableaux de compositions et de correspondance dans l'onglet Métadonnées.",
            )
            return

        # Récupérer les fichiers .txt depuis l'onglet Fichiers
        file_picker = getattr(main, "file_picker", None)
        txt_files = file_picker.get_selected_files() if (file_picker is not None and hasattr(file_picker, "get_selected_files")) else []
        if not txt_files:
            self._combined_df = None
            self._analysis_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — sélectionnez d'abord des fichiers .txt")
            QMessageBox.warning(
                self,
                "Fichiers manquants",
                "Aucun fichier .txt sélectionné. Allez dans l'onglet Fichiers et ajoutez des spectres.",
            )
            return

        # Construire le DataFrame de métadonnées fusionnées (correspondance ↔ composition)
        try:
            merged_meta = metadata_creator.build_merged_metadata()
        except Exception as e:
            self._combined_df = None
            self._analysis_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — erreur lors de la fusion des métadonnées")
            QMessageBox.critical(
                self,
                "Erreur métadonnées",
                f"Impossible de construire les métadonnées fusionnées (correspondance ↔ composition) :\n{e}",
            )
            return

        # Construire le DataFrame combiné (spectres + métadonnées)
        try:
            combined = build_combined_dataframe_from_ui(
                txt_files,
                metadata_creator,
                poly_order=5,
                exclude_brb=True,
                apply_baseline=True,
            )
        except Exception as e:
            self._combined_df = None
            self._analysis_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — erreur lors de l'assemblage")
            QMessageBox.critical(
                self,
                "Échec assemblage",
                f"Impossible d'assembler les données spectres + métadonnées :\n{e}",
            )
            return

        self._combined_df = combined
        self.lbl_status.setText("Fichier combiné : chargé ✓")
        # Nouvelles données => analyse doit être relancée
        self._analysis_dirty = True
        self._peak_intensities = None
        self._ratios_long = None
        self.btn_export.setEnabled(False)
        self._refresh_button_states()
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
        # Ajuster les largeurs de colonnes au contenu
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def _toggle_table_visibility(self):
        """Affiche ou masque le tableau des intensités pour laisser plus de place au graphique."""
        is_visible = self.table.isVisible()
        # Bascule visibilité
        self.table.setVisible(not is_visible)
        if hasattr(self, "lbl_table_title"):
            self.lbl_table_title.setVisible(not is_visible)
        # Met à jour le texte du bouton
        if not is_visible:
            self.btn_toggle_table.setText("Masquer le tableau des intensités")
        else:
            self.btn_toggle_table.setText("Afficher le tableau des intensités")

    # ---------- Analyse ----------
    def _run_analysis(self):
        if self._combined_df is None or self._combined_df.empty:
            self._analysis_dirty = True
            self._refresh_button_states()
            QMessageBox.warning(self, "Données manquantes", "Aucun fichier combiné chargé. Allez dans l’onglet Métadonnées, assemblez puis validez le choix.")
            return

        # On travaille sur une copie locale pour ne pas modifier le fichier combiné global
        df = self._combined_df.copy()
        

        required = {"Raman Shift", "Intensity_corrected", "file"}
        if not required.issubset(df.columns):
            QMessageBox.critical(self, "Colonnes manquantes", "Le fichier combiné doit contenir : Raman Shift, Intensity_corrected, file.")
            return

        # Exclure de l'analyse les spectres dont le nom se termine par BI ou BR,
        # mais conserver ceux qui se terminent par BRB.
        # On utilise 'Spectrum name' si disponible.
        if "Spectrum name" in df.columns:
            sn = df["Spectrum name"].astype(str)
            mask_bi_br = sn.str.match(r".*(BI|BR)$")
            mask_brb = sn.str.endswith("BRB")
            # garder tout ce qui n'est pas BI/BR, ou qui est BRB
            df = df[~mask_bi_br | mask_brb].copy()

        # Si mode "Personnalisé" : interpréter directement la saisie comme des décalages Raman (cm⁻¹)
        if self.cmb_presets.currentIndex() == 2:
            text = self.edit_custom_waves.text() if hasattr(self, "edit_custom_waves") else ""
            if not text.strip():
                QMessageBox.warning(
                    self,
                    "Paramètre manquant",
                    "En mode personnalisé, veuillez saisir au moins un pic à étudier\n"
                    "(Raman Shift en cm⁻¹, séparés par des virgules)."
                )
                return

            # Parsing des valeurs en cm⁻¹
            raw_tokens = [t.strip() for t in text.replace(";", ",").split(",") if t.strip()]
            custom_peaks = []
            for token in raw_tokens:
                token_clean = token.replace(",", ".")
                try:
                    rs = float(token_clean)
                    custom_peaks.append(int(round(rs)))
                except ValueError:
                    continue

            if not custom_peaks:
                QMessageBox.warning(
                    self,
                    "Paramètres invalides",
                    "Aucun Raman Shift valide n'a été trouvé dans le champ personnalisé.\n"
                    "Utilisez des valeurs numériques en cm⁻¹, séparées par des virgules."
                )
                return

            self._peaks = custom_peaks

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

        # Harmonisation de la colonne utilisée pour l'axe X (titrant)
        # Objectif : disposer d'une colonne 'n(titrant) (mol)' quand c'est possible.
        #
        # Cas gérés :
        # - anciens fichiers : 'C (EGTA) (M)' + 'V cuvette (mL)'
        # - pipeline MetadataCreator : '[titrant] (M)' + 'V cuvette (µL)'
        # - fallback : détection automatique d'une colonne de concentration en M
        #   (ex: 'C (EGTA) (M)', 'C (PAN) (M)', 'C (Cu) (M)', etc.).

        def _to_num(s: pd.Series) -> pd.Series:
            return pd.to_numeric(s.astype(str).str.replace(",", "."), errors="coerce")

        def _pick_auto_titrant_column(df_: pd.DataFrame) -> str | None:
            explicit = ["[titrant] (M)", "[titrant] (mM)", "C (EGTA) (M)"]

            def is_valid(col):
                v = _to_num(df_[col])
                return (
                    v.notna().any()
                    and v.nunique(dropna=True) > 2
                    and float(v.var(skipna=True)) > 0.0
                )

            for cand in explicit:
                if cand in df_.columns and is_valid(cand):
                    return cand

            cands = [c for c in df_.columns if re.match(r"^C \(.+\) \(M\)$", str(c))]
            scored = []
            for c in cands:
                v = _to_num(df_[c])
                if v.notna().any() and v.nunique(dropna=True) > 2:
                    scored.append((float(v.var(skipna=True)), c))

            if scored:
                return max(scored)[1]

            return None

        if "n(titrant) (mol)" not in merged.columns:
            tit_col = _pick_auto_titrant_column(merged)

            if tit_col is None:
                # Pas de colonne de concentration utilisable -> on ne peut pas tracer en fonction du titrant
                # (les ratios peuvent exister mais pas d'axe X)
                pass
            else:
                c = _to_num(merged[tit_col])

                # Si on a un volume, on calcule n = C * V(L)
                if "V cuvette (µL)" in merged.columns:
                    v_ul = _to_num(merged["V cuvette (µL)"])
                    merged["n(titrant) (mol)"] = c * v_ul * 1e-6
                elif "V cuvette (mL)" in merged.columns:
                    v_ml = _to_num(merged["V cuvette (mL)"])
                    merged["n(titrant) (mol)"] = c * v_ml * 1e-3
                else:
                    # Pas de volume explicite : on utilise la concentration comme proxy d'axe X
                    merged["n(titrant) (mol)"] = c

                # Garder une trace de la colonne utilisée (utile pour debug / UI)
                if "titrant_column" not in merged.columns:
                    merged["titrant_column"] = tit_col

        # Mise en forme long pour les ratios
        ratio_cols = [c for c in merged.columns if c.startswith("ratio_I_")]
        id_vars = [c for c in ["file", "Spectrum name", "Sample description", "n(titrant) (mol)"] if c in merged.columns]
        df_ratios = merged.melt(id_vars=id_vars, value_vars=ratio_cols, var_name="Ratio", value_name="Value") if ratio_cols else None

        # Mémoriser + aperçu
        self._peak_intensities = merged
        self._ratios_long = df_ratios
        self._set_preview(merged)
        self.btn_export.setEnabled(True)
        self._analysis_dirty = False
        self._refresh_button_states()
        # Réinitialiser état du fit sigmoïde
        self._sigmoid_params = None
        self._sigmoid_cov = None
        self.btn_show_equivalence.setEnabled(False)
        self._sigmoid_results = {}

        # --- Tracé interactif (équivalent au ggplot fourni) ---
        if df_ratios is not None and not df_ratios.empty and "n(titrant) (mol)" in df_ratios.columns:
            # Mettre à jour la liste des ratios disponibles pour le fit
            self.cmb_fit_ratio.blockSignals(True)
            self.cmb_fit_ratio.clear()
            for r in sorted(df_ratios["Ratio"].unique()):
                self.cmb_fit_ratio.addItem(r)
            self.cmb_fit_ratio.blockSignals(False)

            df_plot = df_ratios.sort_values(["Ratio", "n(titrant) (mol)"])

            # Palette de couleurs plus riche et contrastée pour éviter les répétitions trop rapides
            # On concatène plusieurs palettes qualitatives de Plotly pour avoir plus de couleurs distinctes.
            color_seq = (
                px.colors.qualitative.Alphabet
                + px.colors.qualitative.Safe
                + px.colors.qualitative.Dark24
            )

            fig = px.line(
                df_plot,
                x="n(titrant) (mol)",
                y="Value",
                color="Ratio",
                markers=True,
                title="Rapports d’intensité Raman selon [titrant]",
                color_discrete_sequence=color_seq,
            )

            # Améliorer la lisibilité : courbes un peu plus épaisses, marqueurs plus visibles,
            # légende positionnée à droite.
            fig.update_traces(line=dict(width=2), marker=dict(size=6))
            fig.update_layout(
                xaxis_title="Quantité de titrant (mol)",
                yaxis_title="Rapport d’intensité (a.u.)",
                width=1200,
                height=600,
                legend_title_text="Ratio",
                legend=dict(
                    orientation="v",
                    x=1.02,
                    y=1,
                    bordercolor="black",
                    borderwidth=1,
                    bgcolor="rgba(255,255,255,0.8)",
                ),
            )
            # Formatage type scientifique pour l'axe X
            fig.update_xaxes(tickformat=".0e")
            self._current_fig = fig
            self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))
        else:
            # Effacer le graphique si pas de ratios
            self.plot_view.setHtml("<i>Aucun ratio à afficher.</i>")

        # (Suppression du reset du fit et de la combo des ratios ici -- déjà fait plus haut)

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
