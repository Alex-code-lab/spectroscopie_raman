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
    QComboBox, QDoubleSpinBox, QHeaderView, QListView, QAbstractItemView, QLineEdit
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QModelIndex

from data_processing import build_combined_dataframe_from_df


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

        # Init valeurs
        self._on_preset_changed(self.cmb_presets.currentIndex())
        # Couleurs initiales des boutons
        self._refresh_button_states()
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
            combined = build_combined_dataframe_from_df(
                txt_files,
                merged_meta,
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
        # Objectif : disposer d'une colonne 'n(titrant) (mol)' quand c'est possible,
        # en restant compatible avec :
        #  - les anciens fichiers (GC514) qui utilisent 'C (EGTA) (M)' et volume en mL
        #  - les nouveaux tableaux générés dans MetadataCreator, qui fournissent
        #    '[titrant] (M)' et 'V cuvette (µL)'.
        if "n(titrant) (mol)" not in merged.columns:
            # 1) Cas anciens fichiers : C (EGTA) (M) et volume en mL
            if "C (EGTA) (M)" in merged.columns and "V cuvette (mL)" in merged.columns:
                c = pd.to_numeric(merged["C (EGTA) (M)"], errors="coerce")
                v_ml = pd.to_numeric(merged["V cuvette (mL)"], errors="coerce")
                # n = C * V(L) = C * V(mL)/1000
                merged["n(titrant) (mol)"] = c * v_ml * 1e-3
            elif "C (EGTA) (M)" in merged.columns:
                # Cas Angelina : pas de volume explicite, on utilise la concentration comme proxy
                c = pd.to_numeric(merged["C (EGTA) (M)"], errors="coerce")
                merged["n(titrant) (mol)"] = c

            # 2) Nouveau pipeline : colonne [titrant] (M) et volume en µL ou mL
            elif "[titrant] (M)" in merged.columns and "V cuvette (µL)" in merged.columns:
                c = pd.to_numeric(merged["[titrant] (M)"], errors="coerce")
                v_ul = pd.to_numeric(merged["V cuvette (µL)"], errors="coerce")
                # n = C * V(L) = C * V(µL) * 1e-6
                merged["n(titrant) (mol)"] = c * v_ul * 1e-6
            elif "[titrant] (M)" in merged.columns and "V cuvette (mL)" in merged.columns:
                c = pd.to_numeric(merged["[titrant] (M)"], errors="coerce")
                v_ml = pd.to_numeric(merged["V cuvette (mL)"], errors="coerce")
                merged["n(titrant) (mol)"] = c * v_ml * 1e-3
            elif "[titrant] (M)" in merged.columns:
                # Pas de volume explicite : on utilise la concentration comme axe,
                # ce qui garde un axe croissant en titrant (unité "mol" approximative ici).
                c = pd.to_numeric(merged["[titrant] (M)"], errors="coerce")
                merged["n(titrant) (mol)"] = c

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

        # --- Tracé interactif (équivalent au ggplot fourni) ---
        if df_ratios is not None and not df_ratios.empty and "n(titrant) (mol)" in df_ratios.columns:
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


class MetadataPickerWidget(QWidget):
    def __init__(self, start_dir: str | None = None, parent=None):
        super().__init__(parent)

        self.metadata_comp_path = None   # fichier compositions des tubes
        self.metadata_map_path = None    # fichier correspondance spectres/tubes
        self.df_comp = None              # DataFrame compositions
        self.df_map = None               # DataFrame correspondance
        self.combined_df = None

        root = start_dir or os.path.expanduser("~")

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel(
            "<b>Sélection des deux fichiers de métadonnées nécessaires</b><br>"
            "<ul>"
            "<li><b>Fichier de composition des tubes :</b> contient les concentrations (ex. C(EGTA)) et les volumes.</li>"
            "<li><b>Fichier de correspondance spectre → tube :</b> associe chaque nom de spectre (GCXXX_…) au tube correspondant.</li>"
            "</ul>"
            "Sélectionnez chaque fichier puis cliquez sur le bouton correspondant pour le valider.",
            self
        ))

        self.model = QFileSystemModel(self)
        self.model.setRootPath(root)
        root_index = self.model.index(root)

        self.view = QListView(self)
        self.view.setModel(self.model)
        self.view.setRootIndex(root_index)
        self.view.setViewMode(QListView.IconMode)
        self.view.setResizeMode(QListView.Adjust)
        self.view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.view.doubleClicked.connect(self._on_double_clicked)

        layout.addWidget(self.view, 1)

        btns_select = QHBoxLayout()
        # Boutons de validation des deux fichiers de métadonnées
        self.btn_set_comp = QPushButton("Valider le fichier de composition des tubes", self)
        self.btn_set_map = QPushButton("Valider le fichier de correspondance des noms de spectres", self)
        self.btn_set_comp.clicked.connect(self._set_comp_file)
        self.btn_set_map.clicked.connect(self._set_map_file)
        btns_select.addWidget(self.btn_set_comp)
        btns_select.addWidget(self.btn_set_map)
        layout.addLayout(btns_select)

        # Info fichiers choisis
        self.info = QLabel("Aucun fichier de métadonnées sélectionné", self)

        paths_layout = QVBoxLayout()
        paths_layout.addWidget(QLabel("Fichier compositions des tubes :", self))
        self.comp_path_edit = QLineEdit("", self)
        self.comp_path_edit.setReadOnly(True)
        paths_layout.addWidget(self.comp_path_edit)

        paths_layout.addWidget(QLabel("Fichier correspondance spectres/tubes :", self))
        self.map_path_edit = QLineEdit("", self)
        self.map_path_edit.setReadOnly(True)
        paths_layout.addWidget(self.map_path_edit)

        layout.addLayout(paths_layout)

        # Boutons pour assembler et valider
        btns = QHBoxLayout()
        self.btn_assemble = QPushButton("Assembler les métadonnées avec les spectres", self)
        self.btn_validate = QPushButton("Valider le fichier combiné", self)
        self.btn_validate.setEnabled(False)
        self.btn_assemble.clicked.connect(self._on_assemble_clicked)
        self.btn_validate.clicked.connect(self._on_validate_clicked)
        btns.addWidget(self.btn_assemble)
        btns.addWidget(self.btn_validate)
        layout.addLayout(btns)
        layout.addWidget(self.info)

        self._update_info_label()

    def _update_info_label(self):
        comp = os.path.basename(self.metadata_comp_path) if self.metadata_comp_path else "non choisi"
        mapping = os.path.basename(self.metadata_map_path) if self.metadata_map_path else "non choisi"
        self.info.setText(f"Compositions : {comp} — Correspondance : {mapping}")

    def _set_comp_file(self):
        indexes = self.view.selectedIndexes()
        if not indexes:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un fichier Excel ou CSV pour les compositions.")
            return

        file_path = self.model.filePath(indexes[0])
        if not (file_path.endswith(".xlsx") or file_path.endswith(".xls") or file_path.endswith(".csv")):
            QMessageBox.warning(self, "Erreur", "Le fichier de compositions doit être un Excel (.xlsx/.xls) ou un CSV (.csv).")
            return

        self.metadata_comp_path = file_path
        self.comp_path_edit.setText(file_path)

        # Charger le fichier de compositions
        try:
            if file_path.endswith(".csv"):
                self.df_comp = pd.read_csv(file_path, sep=None, engine="python")
            else:
                self.df_comp = pd.read_excel(file_path)
            self._update_preview(self.df_comp)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier de compositions : {e}")
            self.df_comp = None

        self._update_info_label()

    def _set_map_file(self):
        indexes = self.view.selectedIndexes()
        if not indexes:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un fichier Excel ou CSV pour la correspondance.")
            return

        file_path = self.model.filePath(indexes[0])
        if not (file_path.endswith(".xlsx") or file_path.endswith(".xls") or file_path.endswith(".csv")):
            QMessageBox.warning(self, "Erreur", "Le fichier de correspondance doit être un Excel (.xlsx/.xls) ou un CSV (.csv).")
            return

        self.metadata_map_path = file_path
        self.map_path_edit.setText(file_path)

        # Charger le fichier de correspondance
        try:
            if file_path.endswith(".csv"):
                self.df_map = pd.read_csv(file_path, sep=None, engine="python")
            else:
                self.df_map = pd.read_excel(file_path)
            self._update_preview(self.df_map)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier de correspondance : {e}")
            self.df_map = None

        self._update_info_label()

    def _on_double_clicked(self, index: QModelIndex):
        path = self.model.filePath(index)
        if self.model.isDir(index):
            self._set_root_path(path)
            return
        # fichier : simple aperçu
        if path.endswith(".xlsx") or path.endswith(".xls") or path.endswith(".csv"):
            try:
                if path.endswith(".csv"):
                    df_preview = pd.read_csv(path, sep=None, engine="python")
                else:
                    df_preview = pd.read_excel(path)
                self._update_preview(df_preview)
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier : {e}")
        else:
            QMessageBox.information(self, "Info", "Sélectionnez un fichier Excel (.xlsx ou .xls) ou un CSV (.csv).")

    def _set_root_path(self, path: str):
        root_index = self.model.index(path)
        self.view.setRootIndex(root_index)

    def _update_preview(self, df: pd.DataFrame):
        # Affiche un aperçu dans une nouvelle fenêtre ou dans un widget prévu
        # Ici on affiche dans un message box simplifié (exemple)
        if df is None or df.empty:
            QMessageBox.information(self, "Aperçu", "Aucun contenu à afficher.")
            return
        preview_str = df.head(10).to_string()
        QMessageBox.information(self, "Aperçu des données", preview_str)

    def _on_assemble_clicked(self):
        """Construit le fichier global (spectres .txt + métadonnées) en utilisant les fichiers sélectionnés dans l'onglet Fichiers."""
        # Récupérer la fenêtre principale pour accéder au FilePickerWidget
        main = self.window()
        file_picker = getattr(main, 'file_picker', None)
        if file_picker is None:
            QMessageBox.critical(self, "Erreur", "Sélecteur de fichiers (.txt) introuvable.")
            return
        txt_files = file_picker.get_selected_files() if hasattr(file_picker, 'get_selected_files') else []
        if not txt_files:
            QMessageBox.warning(self, "Données manquantes", "Sélectionnez d'abord des fichiers .txt dans l'onglet Fichiers.")
            return

        if self.df_comp is None or not isinstance(self.df_comp, pd.DataFrame):
            QMessageBox.warning(self, "Données manquantes", "Sélectionnez et chargez un fichier de compositions des tubes.")
            return
        if self.df_map is None or not isinstance(self.df_map, pd.DataFrame):
            QMessageBox.warning(self, "Données manquantes", "Sélectionnez et chargez un fichier de correspondance spectres/tubes.")
            return

        # Vérification des colonnes attendues
        if "Tube" not in self.df_comp.columns or "Tube" not in self.df_map.columns:
            QMessageBox.critical(self, "Colonnes manquantes",
                                 "Les deux fichiers de métadonnées doivent contenir une colonne 'Tube' (identifiant du tube).")
            return
        if "Spectrum name" not in self.df_map.columns:
            QMessageBox.critical(self, "Colonnes manquantes",
                                 "Le fichier de correspondance doit contenir une colonne 'Spectrum name' "
                                 "correspondant aux noms de spectres (sans l'extension .txt).")
            return

        try:
            # Jointure de la correspondance avec les compositions via la colonne 'Tube'
            merged_meta = self.df_map.merge(self.df_comp, on="Tube", how="left")

            # Construction du DataFrame complet (spectres + métadonnées fusionnées)
            combined = build_combined_dataframe_from_df(txt_files, merged_meta, exclude_brb=True)
            # Filtrer : supprimer les manips dont le nom finit par BI ou BR,
            # mais conserver celles finissant par BRB
            if "Spectrum name" in combined.columns:
                mask_bi_br = combined["Spectrum name"].astype(str).str.match(r".*(BI|BR)$", case=True)
                mask_brb = combined["Spectrum name"].astype(str).str.endswith("BRB")
                # garder tout ce qui n'est pas BI/BR OU qui est BRB
                combined = combined[~mask_bi_br | mask_brb].copy()
        except Exception as e:
            QMessageBox.critical(self, "Échec", f"Impossible d'assembler les données : {e}")
            return

        self.combined_df = combined
        # Aperçu dans le tableau
        self._update_preview(self.combined_df)
        self.info.setText(f"Assemblage réussi — {combined.shape[0]} lignes, {combined.shape[1]} colonnes")
        self.btn_validate.setEnabled(True)

    def _on_validate_clicked(self):
        # Valide le DataFrame combiné et le rend accessible à l'application
        if self.combined_df is None:
            QMessageBox.warning(self, "Aucune donnée", "Aucun fichier combiné à valider.")
            return
        base_dir = os.path.dirname(self.metadata_map_path or self.metadata_comp_path or os.path.expanduser('~'))
        # Sauvegarde ou autre traitement ici
        QMessageBox.information(self, "Validé", "Le fichier combiné a été validé.")

    def load_metadata(self) -> pd.DataFrame | None:
        return self.combined_df