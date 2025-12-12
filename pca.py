from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
from data_processing import build_combined_dataframe_from_df

from sklearn.decomposition import PCA

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QDoubleSpinBox,
    QSpinBox,
    QMessageBox,
    QDialog,
)
from PySide6.QtWebEngineWidgets import QWebEngineView


class PCATab(QWidget):
    """
    Onglet PCA – réalise une analyse en composantes principales sur les spectres Raman
    à partir du fichier combiné (txt + métadonnées) construit dans l'onglet Métadonnées.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self._combined_df: pd.DataFrame | None = None
        self._scores_df: pd.DataFrame | None = None
        self._loadings_df: pd.DataFrame | None = None
        self._pca: PCA | None = None
        self._index_col: str | None = None
        # --- Nouveaux attributs pour reconstruction et sélection loadings ---
        self._X_proc: np.ndarray | None = None
        self._wavenumbers: np.ndarray | None = None
        self._spec_ids: np.ndarray | None = None
        # Indique que la PCA doit être (re)lancée (données ou paramètres modifiés)
        self._pca_dirty: bool = True

        layout = QVBoxLayout(self)

        # ---- En-tête explicatif ----
        header = QLabel(
            "<b>Analyse PCA (réduction de dimension)</b><br>"
            "La PCA (Analyse en Composantes Principales) projette les spectres Raman dans un espace de "
            "dimensions réduites (PC1, PC2, …) en cherchant les combinaisons linéaires de longueurs d’onde "
            "qui expliquent le plus de variance. Cela permet d’identifier des spectres atypiques (hot spots), "
            "des groupes de spectres similaires ou des régions spectrales particulièrement variables."
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        # ---- Statut fichier combiné ----
        self.lbl_status = QLabel("Fichier combiné : non chargé")
        layout.addWidget(self.lbl_status)

        # ---- Contrôles principaux ----
        ctrl1 = QHBoxLayout()
        self.btn_reload = QPushButton("Recharger le fichier combiné depuis Métadonnées", self)
        self.btn_reload.clicked.connect(self._reload_combined)
        ctrl1.addWidget(self.btn_reload)

        ctrl1.addWidget(QLabel("Raman min (cm⁻¹):"))
        self.spin_min_shift = QDoubleSpinBox(self)
        self.spin_min_shift.setRange(0.0, 5000.0)
        self.spin_min_shift.setDecimals(1)
        self.spin_min_shift.setSingleStep(10.0)
        ctrl1.addWidget(self.spin_min_shift)
        self.spin_min_shift.valueChanged.connect(self._on_pca_param_changed)


        ctrl1.addWidget(QLabel("Raman max (cm⁻¹):"))
        self.spin_max_shift = QDoubleSpinBox(self)
        self.spin_max_shift.setRange(0.0, 5000.0)
        self.spin_max_shift.setDecimals(1)
        self.spin_max_shift.setSingleStep(10.0)
        ctrl1.addWidget(self.spin_max_shift)
        self.spin_max_shift.valueChanged.connect(self._on_pca_param_changed)


        layout.addLayout(ctrl1)

        # ---- Contrôles avancés ----
        ctrl2 = QHBoxLayout()

        ctrl2.addWidget(QLabel("Normalisation des spectres :"))
        self.cmb_norm = QComboBox(self)
        self.cmb_norm.addItems([
            "Aucune (centrage seulement)",
            "Norme L2 par spectre",
            "Standardisation (z-score par longueur d'onde)",
        ])
        ctrl2.addWidget(self.cmb_norm)
        self.cmb_norm.currentIndexChanged.connect(self._on_pca_param_changed)


        ctrl2.addWidget(QLabel("Nombre de composantes :"))
        self.spin_n_comp = QSpinBox(self)
        self.spin_n_comp.setRange(2, 20)
        self.spin_n_comp.setValue(5)
        ctrl2.addWidget(self.spin_n_comp)
        self.spin_n_comp.valueChanged.connect(self._on_pca_param_changed)

        ctrl2.addWidget(QLabel("Colorer par :"))
        self.cmb_color = QComboBox(self)
        self.cmb_color.addItem("(aucune)")
        ctrl2.addWidget(self.cmb_color)
        # La coloration est un réglage d'affichage : ne rend pas la PCA "dirty".
        self.cmb_color.currentIndexChanged.connect(self._on_color_changed)
        layout.addLayout(ctrl2)

        # ---- Boutons d'action ----
        ctrl_buttons = QHBoxLayout()
        self.btn_run = QPushButton("Lancer la PCA", self)
        self.btn_run.clicked.connect(self._run_pca)
        ctrl_buttons.addWidget(self.btn_run)

        layout.addLayout(ctrl_buttons)

        # ---- Résumé variance ----
        self.lbl_variance = QLabel("Variance expliquée : -")
        layout.addWidget(self.lbl_variance)

        # ---- Choix des composantes à afficher pour les loadings ----
        load_ctrl = QHBoxLayout()
        load_ctrl.addWidget(QLabel("Composantes à afficher (loadings) :"))
        self.cmb_loadings_mode = QComboBox(self)
        self.cmb_loadings_mode.addItem("PC1 & PC2")
        self.cmb_loadings_mode.currentIndexChanged.connect(self._update_loadings_plot)
        load_ctrl.addWidget(self.cmb_loadings_mode)
        layout.addLayout(load_ctrl)

        # ---- Choix du spectre pour la reconstruction ----
        recon_ctrl = QHBoxLayout()
        recon_ctrl.addWidget(QLabel("Spectre à reconstruire :"))
        self.cmb_spec = QComboBox(self)
        self.cmb_spec.addItem("(aucun)")
        self.cmb_spec.currentIndexChanged.connect(self._update_reconstruction_plot)
        recon_ctrl.addWidget(self.cmb_spec)
        layout.addLayout(recon_ctrl)

        # ---- Graphiques ----
        layout.addWidget(QLabel("<b>Scores PCA (PC1 vs PC2)</b>"))
        self.btn_show_scores = QPushButton("Afficher les scores dans une fenêtre", self)
        self.btn_show_scores.clicked.connect(self._show_scores_window)
        layout.addWidget(self.btn_show_scores)

        layout.addWidget(QLabel("<b>Loadings PCA (poids des composantes en fonction du shift Raman)</b>"))
        self.btn_show_loadings = QPushButton("Afficher les loadings dans une fenêtre", self)
        self.btn_show_loadings.clicked.connect(self._show_loadings_window)
        layout.addWidget(self.btn_show_loadings)

        # ---- Vue pour la reconstruction de spectres ----
        layout.addWidget(QLabel("<b>Spectre original vs reconstruit (dans l'espace normalisé)</b>"))
        self.btn_show_recon = QPushButton("Afficher la reconstruction dans une fenêtre", self)
        self.btn_show_recon.clicked.connect(self._show_recon_window)
        layout.addWidget(self.btn_show_recon)

        self._init_plot_windows()
        # Couleurs initiales des boutons (rouge tant que rien n'est chargé)
        self._refresh_button_states()


    def _refresh_button_states(self) -> None:
        """Met à jour la couleur des boutons selon l'état du fichier combiné et de la PCA."""
        red = "background-color: #d9534f; color: white; font-weight: 600;"
        green = "background-color: #5cb85c; color: white; font-weight: 600;"

        combined_ready = (
            self._combined_df is not None
            and isinstance(self._combined_df, pd.DataFrame)
            and not self._combined_df.empty
        )

        pca_ready = (
            combined_ready
            and (self._pca_dirty is False)
            and (self._pca is not None)
            and (self._scores_df is not None and not self._scores_df.empty)
        )

        # Bouton reload : vert si le fichier combiné est chargé
        self.btn_reload.setStyleSheet(green if combined_ready else red)
        self.btn_reload.setToolTip(
            "Vert = fichier combiné chargé" if combined_ready else "Rouge = recharger le fichier combiné"
        )

        # Bouton PCA : vert seulement si PCA déjà lancée et à jour
        self.btn_run.setStyleSheet(green if pca_ready else red)
        self.btn_run.setToolTip(
            "Vert = PCA à jour"
            if pca_ready
            else ("Rouge = lancez (ou relancez) la PCA" if combined_ready else "Rouge = chargez d'abord le fichier combiné")
        )

    def _on_pca_param_changed(self, *args) -> None:
        """Marque la PCA comme à relancer si un paramètre change."""
        self._pca_dirty = True
        self._refresh_button_states()

    def _on_color_changed(self, *args) -> None:
        """Recolore le nuage PC1 vs PC2 sans relancer la PCA."""
        if self._scores_df is not None and not self._scores_df.empty:
            self._update_scores_plot()


    def _init_plot_windows(self):
        """Crée les fenêtres flottantes (dialogues) pour afficher les graphiques PCA."""
        # Fenêtre pour les scores
        self.scores_dialog = QDialog(self)
        self.scores_dialog.setWindowTitle("Scores PCA (PC1 vs PC2)")
        scores_layout = QVBoxLayout(self.scores_dialog)
        self.scores_view = QWebEngineView(self.scores_dialog)
        scores_layout.addWidget(self.scores_view)
        self.scores_dialog.resize(900, 500)

        # Fenêtre pour les loadings
        self.loadings_dialog = QDialog(self)
        self.loadings_dialog.setWindowTitle("Loadings PCA")
        load_layout = QVBoxLayout(self.loadings_dialog)
        self.loadings_view = QWebEngineView(self.loadings_dialog)
        load_layout.addWidget(self.loadings_view)
        self.loadings_dialog.resize(900, 500)

        # Fenêtre pour la reconstruction
        self.recon_dialog = QDialog(self)
        self.recon_dialog.setWindowTitle("Spectre original vs reconstruit")
        recon_layout = QVBoxLayout(self.recon_dialog)
        self.recon_view = QWebEngineView(self.recon_dialog)
        recon_layout.addWidget(self.recon_view)
        self.recon_dialog.resize(900, 500)

    def _on_pca_param_changed(self, *args) -> None:
        """Marque la PCA comme à relancer si un paramètre change."""
        self._pca_dirty = True
        self._refresh_button_states()

    def _show_scores_window(self):
        """Affiche la fenêtre contenant le nuage de scores PCA."""
        if self._scores_df is None or self._scores_df.empty:
            QMessageBox.information(self, "Info", "Aucun score PCA disponible. Lancez d'abord la PCA.")
            return
        self.scores_dialog.show()
        self.scores_dialog.raise_()
        self.scores_dialog.activateWindow()

    def _show_loadings_window(self):
        """Affiche la fenêtre contenant les loadings PCA."""
        if self._loadings_df is None or self._loadings_df.empty:
            QMessageBox.information(self, "Info", "Aucun loading PCA disponible. Lancez d'abord la PCA.")
            return
        self.loadings_dialog.show()
        self.loadings_dialog.raise_()
        self.loadings_dialog.activateWindow()

    def _show_recon_window(self):
        """Affiche la fenêtre contenant la reconstruction du spectre choisi."""
        if self._X_proc is None or self._pca is None or self._wavenumbers is None or self._spec_ids is None:
            QMessageBox.information(self, "Info", "Aucune donnée PCA disponible pour la reconstruction. Lancez d'abord la PCA.")
            return
        self.recon_dialog.show()
        self.recon_dialog.raise_()
        self.recon_dialog.activateWindow()

    # ------------------------------------------------------------------
    # Chargement et préparation des données
    # ------------------------------------------------------------------
    def _reload_combined(self):
        """Reconstruit le DataFrame combiné à partir des métadonnées créées dans l'onglet Métadonnées
        et des fichiers .txt sélectionnés dans l'onglet Fichiers.
        """
        main = self.window()
        if main is None:
            QMessageBox.critical(self, "Erreur", "Fenêtre principale introuvable.")
            return

        metadata_creator = getattr(main, "metadata_creator", None)
        if metadata_creator is None or not hasattr(metadata_creator, "build_merged_metadata"):
            self._combined_df = None
            self._pca_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — définissez d'abord les métadonnées")
            QMessageBox.warning(
                self,
                "Métadonnées manquantes",
                "L'onglet Métadonnées n'est pas correctement initialisé ou ne fournit pas encore les tableaux.\n"
                "Créez d'abord les tableaux de compositions et de correspondance dans l'onglet Métadonnées.",
            )
            return

        file_picker = getattr(main, "file_picker", None)
        txt_files = file_picker.get_selected_files() if (file_picker is not None and hasattr(file_picker, "get_selected_files")) else []
        if not txt_files:
            self._combined_df = None
            self._pca_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — sélectionnez d'abord des fichiers .txt")
            QMessageBox.warning(
                self,
                "Fichiers manquants",
                "Aucun fichier .txt sélectionné. Allez dans l'onglet Fichiers et ajoutez des spectres.",
            )
            return

        try:
            merged_meta = metadata_creator.build_merged_metadata()
            print("\n=== merged_meta.columns ===")
            print(list(merged_meta.columns))

            print("\n=== merged_meta.head() ===")
            print(merged_meta.head())
        except Exception as e:
            self._combined_df = None
            self._pca_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — erreur lors de la fusion des métadonnées")
            QMessageBox.critical(
                self,
                "Erreur métadonnées",
                f"Impossible de construire les métadonnées fusionnées (correspondance ↔ composition) :\n{e}",
            )
            return

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
            self._pca_dirty = True
            self._refresh_button_states()
            self.lbl_status.setText("Fichier combiné : non chargé ✗ — erreur lors de l'assemblage")
            QMessageBox.critical(
                self,
                "Échec assemblage",
                f"Impossible d'assembler les données spectres + métadonnées :\n{e}",
            )
            return

        self._combined_df = combined
        # Nouvelles données => PCA doit être relancée
        self._pca_dirty = True
        self._pca = None
        self._scores_df = None
        self._loadings_df = None
        self._X_proc = None
        self._wavenumbers = None
        self._spec_ids = None

        self.lbl_variance.setText("Variance expliquée : -")
        self.lbl_status.setText(
            f"Fichier combiné : chargé ✓ ({combined['file'].nunique() if 'file' in combined.columns else len(combined)} spectres)"
        )
        self._populate_controls_from_df(combined_df=combined, meta_df=merged_meta)
        # Conserver "(aucune)" sélectionné si possible
        if self.cmb_color.count() > 0:
            self.cmb_color.setCurrentIndex(0)
        self._refresh_button_states()

    def _populate_controls_from_df(self, combined_df: pd.DataFrame, meta_df: pd.DataFrame | None = None) -> None:
        """Met à jour les bornes Raman (depuis combined_df) et la liste des colonnes de coloration.

        Important : la liste "Colorer par" est construite à partir des métadonnées par spectre (meta_df)
        pour inclure correctement les colonnes de concentrations (Solutions A/B/…).
        Si meta_df n'est pas fourni, on retombe sur combined_df.
        """
        # --- bornes Raman (depuis combined_df) ---
        if "Raman Shift" in combined_df.columns:
            try:
                r_min = float(combined_df["Raman Shift"].min())
                r_max = float(combined_df["Raman Shift"].max())
                self.spin_min_shift.setRange(r_min, r_max)
                self.spin_max_shift.setRange(r_min, r_max)
                self.spin_min_shift.setValue(r_min)
                self.spin_max_shift.setValue(r_max)
            except Exception:
                pass

        # --- colonnes de coloration (depuis meta_df si possible) ---
        src = meta_df if (meta_df is not None and isinstance(meta_df, pd.DataFrame) and not meta_df.empty) else combined_df

        # Colonnes à exclure (spectre / techniques / identifiants bruts)
        excluded = {"Raman Shift", "Intensity_corrected", "Dark Subtracted #1"}

        # On évite aussi de proposer la colonne de jointure si elle existe
        if "Spectrum name" in src.columns:
            excluded.add("Spectrum name")
        if "file" in src.columns:
            excluded.add("file")

        # Alimenter la combo
        self.cmb_color.blockSignals(True)
        self.cmb_color.clear()
        self.cmb_color.addItem("(aucune)")

        # 1) D’abord, ajouter en priorité les colonnes "utiles" (titrant / volumes / tubes)
        priority = []
        for p in ["[titrant] (M)", "[titrant] (mM)", "V cuvette (µL)", "Tube", "Sample description"]:
            if p in src.columns and p not in excluded:
                priority.append(p)

        for col in priority:
            self.cmb_color.addItem(col)

        # 2) Puis toutes les autres métadonnées, y compris les solutions (A/B/…)
        for col in src.columns:
            if col in excluded or col in priority:
                continue
            self.cmb_color.addItem(col)

        self.cmb_color.blockSignals(False)

    # ------------------------------------------------------------------
    # PCA principale
    # ------------------------------------------------------------------

    def _populate_color_combo_from_scores(self) -> None:
        """Remplit la combo 'Colorer par' à partir de self._scores_df (réellement utilisé pour le scatter)."""
        if self._scores_df is None or self._scores_df.empty:
            return

        df = self._scores_df
        current = self.cmb_color.currentText() if self.cmb_color.count() else "(aucune)"

        self.cmb_color.blockSignals(True)
        self.cmb_color.clear()
        self.cmb_color.addItem("(aucune)")

        # Exclure les colonnes PCA elles-mêmes + l'identifiant
        excluded = {c for c in df.columns if str(c).startswith("PC")}
        if self._index_col and self._index_col in df.columns:
            excluded.add(self._index_col)

        for col in df.columns:
            if col in excluded:
                continue
            self.cmb_color.addItem(col)

        # Restaurer la sélection si possible
        if current and current != "(aucune)":
            idx = self.cmb_color.findText(current)
            if idx >= 0:
                self.cmb_color.setCurrentIndex(idx)

        self.cmb_color.blockSignals(False)
    def _build_matrix(self) -> tuple[np.ndarray, np.ndarray, pd.DataFrame, np.ndarray]:
        """
        Construit la matrice X (n_spectres × n_shifts) à partir du DataFrame combiné,
        en pivotant Intensity_corrected par 'Spectrum name' (ou 'file') et 'Raman Shift'.

        Returns
        -------
        X : ndarray
            Matrice des intensités normalisées/préparées.
        wavenumbers : ndarray
            Tableau des shifts Raman (ordonnés).
        meta_per_spec : DataFrame
            Tableau des métadonnées par spectre (une ligne par spectre).
        """
        if self._combined_df is None or self._combined_df.empty:
            raise ValueError("Aucun fichier combiné chargé.")

        df = self._combined_df.copy()

        if "Raman Shift" not in df.columns or "Intensity_corrected" not in df.columns:
            raise ValueError("Le fichier combiné doit contenir les colonnes 'Raman Shift' et 'Intensity_corrected'.")

        # Choix de l'identifiant de spectre : Spectrum name si dispo, sinon file
        if "Spectrum name" in df.columns:
            index_col = "Spectrum name"
        elif "file" in df.columns:
            index_col = "file"
        else:
            raise ValueError("Impossible d'identifier les spectres (ni 'Spectrum name' ni 'file' dans les colonnes).")

        self._index_col = index_col

        # Filtrage sur la plage de shifts choisie
        r_min = float(self.spin_min_shift.value())
        r_max = float(self.spin_max_shift.value())
        df = df[(df["Raman Shift"] >= r_min) & (df["Raman Shift"] <= r_max)].copy()
        if df.empty:
            raise ValueError("Aucune donnée dans la plage de Raman choisie.")

        # Pivot : une ligne par spectre, une colonne par shift Raman
        mat = df.pivot_table(
            index=index_col,
            columns="Raman Shift",
            values="Intensity_corrected",
            aggfunc="mean",
        )

        # Supprimer les colonnes complètement vides
        mat = mat.dropna(axis=1, how="all")
        if mat.empty:
            raise ValueError("Matrice des intensités vide après pivot. Vérifiez les données.")

        # Gestion des NaN restants : remplissage par la moyenne de chaque colonne
        mat = mat.apply(lambda col: col.fillna(col.mean()), axis=0)

        X = mat.to_numpy(dtype=float)
        wavenumbers = mat.columns.to_numpy(dtype=float)
        spec_ids = mat.index.to_numpy()

        # Métadonnées par spectre
        # On exclut explicitement la colonne d'index (index_col) pour éviter des doublons
        # lors du reset_index() (sinon pandas essaie d'insérer deux fois la même colonne).
        # Métadonnées par spectre (on exclut les colonnes spectrales)
        meta_cols = [
                    c for c in df.columns
                    if c not in {"Raman Shift", "Intensity_corrected", "Dark Subtracted #1", index_col}
                    ]
        meta_per_spec = df.groupby(index_col)[meta_cols].first().reset_index()

        return X, wavenumbers, meta_per_spec, spec_ids

    def _run_pca(self):
        """Calcule la PCA et met à jour les graphiques."""
        if self._combined_df is None or self._combined_df.empty:
            QMessageBox.warning(self, "Données manquantes", "Aucun fichier combiné chargé. Utilisez le bouton de rechargement.")
            self._refresh_button_states()
            return

        try:
            X, wavenumbers, meta_per_spec, spec_ids = self._build_matrix()
        except Exception as e:
            QMessageBox.critical(self, "Erreur données", f"Impossible de préparer les données pour la PCA :\n{e}")
            return

        n_samples, n_features = X.shape
        if n_samples < 2 or n_features < 2:
            QMessageBox.warning(self, "Données insuffisantes", "Pas assez de spectres ou de points Raman pour effectuer une PCA.")
            return

        # Normalisation choisie
        norm_mode = self.cmb_norm.currentText()
        if norm_mode.startswith("Norme L2"):
            # Normalisation L2 par spectre (ligne)
            norms = np.linalg.norm(X, axis=1, keepdims=True)
            norms[norms == 0] = 1.0
            X_proc = X / norms
        elif norm_mode.startswith("Standardisation"):
            # Standardisation (z-score) par feature (colonne)
            mean = X.mean(axis=0, keepdims=True)
            std = X.std(axis=0, ddof=0, keepdims=True)
            std[std == 0] = 1.0
            X_proc = (X - mean) / std
        else:
            # Aucune normalisation spécifique : on laisse PCA centrer les colonnes
            X_proc = X

        # Stocker les données préparées pour la reconstruction et les graphiques
        self._X_proc = X_proc
        self._wavenumbers = wavenumbers
        self._spec_ids = spec_ids

        # Ajuster le nombre de composantes
        asked = int(self.spin_n_comp.value())
        n_comp = min(asked, n_samples, n_features)
        if n_comp < asked:
            self.spin_n_comp.setValue(n_comp)

        try:
            pca = PCA(n_components=n_comp)
            scores = pca.fit_transform(X_proc)
        except Exception as e:
            QMessageBox.critical(self, "Erreur PCA", f"Échec du calcul de la PCA :\n{e}")
            return

        self._pca = pca

        # DataFrame des scores (coordonnées des spectres dans l'espace PCA)
        comp_names = [f"PC{i+1}" for i in range(n_comp)]
        scores_df = pd.DataFrame(scores, columns=comp_names)
        if self._index_col is not None and self._spec_ids is not None:
            scores_df[self._index_col] = self._spec_ids

        # Jointure avec les métadonnées
        scores_df = scores_df.merge(meta_per_spec, on=self._index_col, how="left")
        self._scores_df = scores_df
        # Re-remplir la liste 'Colorer par' à partir des colonnes réellement disponibles pour les scores
        self._populate_color_combo_from_scores()

        # DataFrame des loadings (poids des composantes par shift Raman)
        loadings = pca.components_.T  # shape: (n_features, n_components)
        loadings_df = pd.DataFrame(loadings, columns=comp_names)
        loadings_df["Raman Shift"] = wavenumbers
        self._loadings_df = loadings_df

        # Met à jour la liste des composantes disponibles pour les loadings
        if hasattr(self, "cmb_loadings_mode"):
            self.cmb_loadings_mode.blockSignals(True)
            self.cmb_loadings_mode.clear()
            # Option combinée par défaut
            if "PC1" in comp_names and "PC2" in comp_names:
                self.cmb_loadings_mode.addItem("PC1 & PC2")
            for name in comp_names:
                self.cmb_loadings_mode.addItem(name)
            self.cmb_loadings_mode.blockSignals(False)

        # Met à jour la liste des spectres pour la reconstruction
        if hasattr(self, "cmb_spec") and self._spec_ids is not None:
            self.cmb_spec.blockSignals(True)
            self.cmb_spec.clear()
            for sid in self._spec_ids:
                self.cmb_spec.addItem(str(sid))
            self.cmb_spec.blockSignals(False)

        # Résumé de la variance expliquée
        var_ratio = pca.explained_variance_ratio_
        parts = [f"PC{i+1}: {vr*100:.1f}%" for i, vr in enumerate(var_ratio)]
        self.lbl_variance.setText("Variance expliquée : " + " | ".join(parts))

        # Mise à jour des graphiques
        self._update_scores_plot()
        self._update_loadings_plot()

        # Met à jour la reconstruction pour le premier spectre si possible
        if hasattr(self, "cmb_spec") and self.cmb_spec.count() > 0:
            self.cmb_spec.setCurrentIndex(0)
            self._update_reconstruction_plot()
        
        # PCA calculée avec succès => état "à jour"
        self._pca_dirty = False
        self._refresh_button_states()

        QMessageBox.information(self, "PCA terminée", "La PCA a été calculée et les graphiques ont été mis à jour.")

    # ------------------------------------------------------------------
    # Graphiques
    # ------------------------------------------------------------------
    def _update_scores_plot(self):
        """Affiche le nuage de points PC1 vs PC2 coloré par la variable choisie."""
        if self._scores_df is None or self._scores_df.empty:
            self.scores_view.setHtml("<i>Aucun score PCA à afficher.</i>")
            return

        df = self._scores_df
        if "PC1" not in df.columns or "PC2" not in df.columns:
            self.scores_view.setHtml("<i>PC1 et PC2 non disponibles.</i>")
            return

        color_col = self.cmb_color.currentText()
        if color_col == "(aucune)" or color_col not in df.columns:
            color_col = None

        hover_cols = []
        if self._index_col and self._index_col in df.columns:
            hover_cols.append(self._index_col)
        for extra in ["file", "Sample description"]:
            if extra in df.columns and extra not in hover_cols:
                hover_cols.append(extra)

        # title = "Scores PCA – PC1 vs PC2"
        # Nom de la manip pour le titre
        manip_name = None
        main = self.window()
        if main is not None:
            metadata_creator = getattr(main, "metadata_creator", None)
            if metadata_creator is not None and hasattr(metadata_creator, "edit_manip"):
                manip_name = metadata_creator.edit_manip.text().strip()

        base_title = manip_name if manip_name else "Scores PCA"
        title = f"{base_title} – PC1 vs PC2"
        
        if self._pca is not None and hasattr(self._pca, "explained_variance_ratio_"):
            vr = self._pca.explained_variance_ratio_
            if len(vr) >= 2:
                title += f" (PC1: {vr[0]*100:.1f}%, PC2: {vr[1]*100:.1f}%)"

        fig = px.scatter(
            df,
            x="PC1",
            y="PC2",
            color=color_col,
            hover_data=hover_cols if hover_cols else None,
            title=title,
        )
        fig.update_layout(
            xaxis_title="PC1",
            yaxis_title="PC2",
            width=900,
            height=450,
            legend_title_text=color_col if color_col else "Groupe",
        )
        self.scores_view.setHtml(fig.to_html(include_plotlyjs="cdn"))

    def _update_loadings_plot(self):
        """Affiche les loadings des composantes choisies en fonction du shift Raman."""
        if self._loadings_df is None or self._loadings_df.empty:
            self.loadings_view.setHtml("<i>Aucun loading PCA à afficher.</i>")
            return

        df = self._loadings_df.copy()
        if "Raman Shift" not in df.columns:
            self.loadings_view.setHtml("<i>Colonne 'Raman Shift' manquante pour les loadings.</i>")
            return

        # Détermination des composantes à tracer
        available_pcs = [c for c in df.columns if c.startswith("PC")]
        mode = self.cmb_loadings_mode.currentText() if hasattr(self, "cmb_loadings_mode") else "PC1 & PC2"

        if mode.startswith("PC1 & PC2") and {"PC1", "PC2"}.issubset(available_pcs):
            value_vars = ["PC1", "PC2"]
        elif mode in available_pcs:
            value_vars = [mode]
        else:
            # Fallback : PC1 seule si présente
            value_vars = ["PC1"] if "PC1" in available_pcs else available_pcs[:1]

        if not value_vars:
            self.loadings_view.setHtml("<i>Aucune composante PCA disponible pour les loadings.</i>")
            return

        df_melt = df.melt(id_vars="Raman Shift", value_vars=value_vars, var_name="Composante", value_name="Loading")

        fig = px.line(
            df_melt.sort_values(["Composante", "Raman Shift"]),
            x="Raman Shift",
            y="Loading",
            color="Composante",
            title="Loadings PCA – contribution des shifts Raman",
        )
        fig.update_layout(
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Loading",
            width=900,
            height=450,
            legend_title_text="Composante",
        )
        self.loadings_view.setHtml(fig.to_html(include_plotlyjs="cdn"))

    def _update_reconstruction_plot(self):
        """Affiche, pour un spectre choisi, le signal original (après normalisation) et sa reconstruction PCA."""
        if self._X_proc is None or self._pca is None or self._wavenumbers is None or self._spec_ids is None:
            self.recon_view.setHtml("<i>Aucune donnée PCA disponible pour la reconstruction.</i>")
            return

        idx = self.cmb_spec.currentIndex() if hasattr(self, "cmb_spec") else -1
        if idx < 0 or idx >= self._X_proc.shape[0]:
            self.recon_view.setHtml("<i>Sélectionnez un spectre pour la reconstruction.</i>")
            return

        # Spectre original (dans l'espace X_proc)
        x_orig = self._X_proc[idx, :].reshape(1, -1)

        # Scores du spectre sélectionné
        comps = self._pca.components_
        mean = self._pca.mean_.reshape(1, -1)
        # Recalcule les scores pour ce spectre dans l'espace des composantes
        # (équivalent aux scores déjà calculés, mais limité à cette ligne)
        x_centered = x_orig - mean
        scores_single = x_centered @ comps.T  # shape: (1, n_components)
        # Reconstruction dans l'espace normalisé X_proc
        x_recon = scores_single @ comps + mean
        x_recon = x_recon.ravel()
        x_orig = x_orig.ravel()

        df_plot = pd.DataFrame({
            "Raman Shift": self._wavenumbers,
            "Original": x_orig,
            "Reconstruit": x_recon,
        })

        df_melt = df_plot.melt(id_vars="Raman Shift", value_vars=["Original", "Reconstruit"],
                               var_name="Type", value_name="Intensité")

        import plotly.express as px  # local import to avoid issues if not used elsewhere

        fig = px.line(
            df_melt.sort_values(["Type", "Raman Shift"]),
            x="Raman Shift",
            y="Intensité",
            color="Type",
            title=f"Spectre original vs reconstruit – {self.cmb_spec.currentText()}",
        )
        fig.update_layout(
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité (espace normalisé)",
            width=900,
            height=450,
            legend_title_text="Signal",
        )
        self.recon_view.setHtml(fig.to_html(include_plotlyjs="cdn"))