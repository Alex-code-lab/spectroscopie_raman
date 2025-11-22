# metadata_picker.py
import os
import re
import numpy as np
import pandas as pd
from data_processing import build_combined_dataframe_from_df
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileSystemModel, QListView, QAbstractItemView,
    QMessageBox, QTableWidget, QTableWidgetItem, QSizePolicy, QHeaderView, QFileDialog
)
from PySide6.QtCore import QDir, QModelIndex

# Dossier par défaut à l'ouverture (modifiez-le si besoin)
DEFAULT_DIR = os.path.expanduser("~/Documents/Travail/CitizenSers/Spectroscopie/")

class MetadataPickerWidget(QWidget):
    def __init__(self, parent=None, start_dir: str | None = None):
        super().__init__(parent)

        self.metadata_comp_path = None   # fichier compositions des tubes
        self.metadata_map_path = None    # fichier correspondance spectres/tubes
        self.df_comp = None              # DataFrame compositions
        self.df_map = None               # DataFrame correspondance
        root = start_dir or os.path.expanduser("~")

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel(
            "<b>Sélection des deux fichiers de métadonnées nécessaires</b><br>"
            "<ul>"
            "<li><b>Fichier de composition des tubes :</b> tableau contenant C(EGTA), volumes, etc.</li>"
            "<li><b>Fichier de correspondance spectres/tubes :</b> permet d’associer correctement chaque spectre à son tube.</li>"
            "</ul>"
            "Sélectionnez un fichier dans la liste, puis cliquez sur le bouton correspondant pour le valider.",
            self
        ))

        # Barre de chemin + bouton Monter
        path_bar = QHBoxLayout()
        self.path_edit = QLineEdit(root, self)
        self.path_edit.setReadOnly(True)
        btn_up = QPushButton("⬆︎ Monter", self)
        btn_up.clicked.connect(self._go_up)
        path_bar.addWidget(QLabel("Dossier :", self))
        path_bar.addWidget(self.path_edit, 1)
        path_bar.addWidget(btn_up)
        layout.addLayout(path_bar)

        # ----- Navigateur fichiers (comme FilePicker) -----
        self.model = QFileSystemModel(self)
        self.model.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot)
        self.model.setNameFilters(["*.xlsx", "*.xls", "*.csv"])
        self.model.setNameFilterDisables(False)
        root = start_dir or DEFAULT_DIR
        root_index = self.model.setRootPath(root)

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

        # ----- Visualiseur de métadonnées -----
        self.table = QTableWidget(self)
        # Configuration pour navigation complète
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(QLabel("<b>Visualisation des données</b>"))
        layout.addWidget(self.table, 2)
        layout.addStretch(1)

        # ----- Actions d'assemblage et validation -----
        actions = QHBoxLayout()
        self.btn_assemble = QPushButton("Assembler (données au format txt + métadonnées)", self)
        self.btn_validate = QPushButton("Enregistrer le fichier", self)
        self.btn_validate.setEnabled(False)
        self.btn_assemble.clicked.connect(self._on_assemble_clicked)
        self.btn_validate.clicked.connect(self._on_validate_clicked)
        actions.addWidget(self.btn_assemble)
        actions.addWidget(self.btn_validate)
        layout.addLayout(actions)
        layout.addWidget(self.info)

        # stockage du DataFrame fusionné
        self.combined_df = None
        self.last_save_path = None

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

        # Vérification / reconstruction des colonnes attendues
        # On essaie d'abord de "réparer" automatiquement si nécessaire.
        if isinstance(self.df_comp, pd.DataFrame) and "Tube" not in self.df_comp.columns:
            try:
                df_auto = self._extract_egta_from_wide(self.df_comp)
            except Exception:
                df_auto = None
            if isinstance(df_auto, pd.DataFrame) and "Tube" in df_auto.columns:
                self.df_comp = df_auto

        if isinstance(self.df_map, pd.DataFrame) and ("Tube" not in self.df_map.columns or "Spectrum name" not in self.df_map.columns):
            try:
                df_auto_map = self._extract_mapping_gc514(self.df_map)
            except Exception:
                df_auto_map = None
            if isinstance(df_auto_map, pd.DataFrame) and "Tube" in df_auto_map.columns and "Spectrum name" in df_auto_map.columns:
                self.df_map = df_auto_map

        # Si après tentative de réparation ça ne va toujours pas, on l'explique clairement.
        if "Tube" not in self.df_comp.columns or "Tube" not in self.df_map.columns:
            msg = (
                "Les deux fichiers de métadonnées doivent contenir une colonne 'Tube' (identifiant du tube).\n\n"
                f"Colonnes fichier COMPOSITION : {list(self.df_comp.columns)}\n"
                f"Colonnes fichier CORRESPONDANCE : {list(self.df_map.columns)}"
            )
            QMessageBox.critical(self, "Colonnes manquantes", msg)
            return
        if "Spectrum name" not in self.df_map.columns:
            msg = (
                "Le fichier de correspondance doit contenir une colonne 'Spectrum name' "
                "correspondant aux noms de spectres (sans l'extension .txt).\n\n"
                f"Colonnes fichier CORRESPONDANCE : {list(self.df_map.columns)}"
            )
            QMessageBox.critical(self, "Colonnes manquantes", msg)
            return

        try:
            # Jointure de la correspondance avec les compositions via la colonne 'Tube'
            merged_meta = self.df_map.merge(self.df_comp, on="Tube", how="left")

            # Construction du DataFrame complet (spectres + métadonnées fusionnées)
            combined = build_combined_dataframe_from_df(txt_files, merged_meta, exclude_brb=True)
        except Exception as e:
            QMessageBox.critical(self, "Échec", f"Impossible d'assembler les données : {e}")
            return

        self.combined_df = combined
        # Aperçu dans le tableau
        self._update_preview(self.combined_df)
        self.info.setText(f"Assemblage réussi — {combined.shape[0]} lignes, {combined.shape[1]} colonnes")
        self.btn_validate.setEnabled(True)

    def _on_validate_clicked(self):
        """Enregistre le fichier combiné en CSV ou Excel, au choix de l'utilisateur."""
        if self.combined_df is None or self.combined_df.empty:
            QMessageBox.warning(self, "Rien à valider", "Assemblez d'abord les données.")
            return

        # Prépare un chemin suggéré
        base_dir = os.path.dirname(self.metadata_map_path or self.metadata_comp_path or os.path.expanduser('~'))
        default_name = "combined_raman.csv"
        suggested = self.last_save_path or os.path.join(base_dir, default_name)

        # Boîte de dialogue : choix du format
        filters = "CSV (*.csv);;Excel (*.xlsx)"
        file_path, selected_filter = QFileDialog.getSaveFileName(self, "Enregistrer le fichier combiné", suggested, filters)
        if not file_path:
            return

        # Normalise l'extension en fonction du filtre choisi
        if selected_filter.startswith("CSV") and not file_path.lower().endswith(".csv"):
            file_path += ".csv"
        elif selected_filter.startswith("Excel") and not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        # Écrit selon l'extension
        try:
            if file_path.lower().endswith('.csv'):
                self.combined_df.to_csv(file_path, index=False)
            else:
                self.combined_df.to_excel(file_path, index=False)
            self.last_save_path = file_path
            QMessageBox.information(self, "Fichier créé", f"Fichier global enregistré :\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Échec écriture", f"Impossible d'écrire le fichier : {e}")

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

        # Charger le fichier de compositions (brut)
        try:
            if file_path.endswith(".csv"):
                df_raw = pd.read_csv(file_path, sep=None, engine="python")
            else:
                df_raw = pd.read_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier de compositions : {e}")
            self.df_comp = None
            return

        # Par défaut, on garde le tableau tel quel
        self.df_comp = df_raw

        # Mais on tente de détecter et convertir la structure "ligne C (EGTA) / colonnes = tubes"
        try:
            df_egta = self._extract_egta_from_wide(df_raw)
        except Exception as e:
            df_egta = None

        if isinstance(df_egta, pd.DataFrame) and not df_egta.empty:
            # On a reconnu la structure GC514 : on bascule vers un tableau long par tube
            self.df_comp = df_egta

        # Aperçu de ce qui sera réellement utilisé pour l'assemblage
        self._update_preview(self.df_comp)
        self._update_info_label()

    def _extract_egta_from_wide(self, df_raw: pd.DataFrame) -> pd.DataFrame | None:
        """
        Tente de convertir un tableau 'large' de type GC514_spectroscopie.xlsx
        en un tableau long avec une ligne par tube et les colonnes :
        Tube, C (EGTA) (M), éventuellement V cuvette (mL) et n(EGTA) (mol).

        Si la structure attendue n'est pas trouvée, retourne None.
        """
        # Cherche la ligne 'C (EGTA)'
        egta_row_idx = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.fullmatch(r"C \(EGTA\)").any():
                egta_row_idx = i
                break
        if egta_row_idx is None:
            return None

        # Cherche la ligne des tubes (cells du type 'tube 1', 'tube 2', ...)
        tube_row_idx = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains(r"^tube\s*\d+\s*$", case=False, regex=True).any():
                tube_row_idx = i
                break
        if tube_row_idx is None:
            return None

        # Cherche éventuellement la ligne 'V cuvette (mL)'
        vol_row_idx = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.fullmatch(r"V cuvette \(mL\)").any():
                vol_row_idx = i
                break

        tube_row = df_raw.iloc[tube_row_idx]
        egta_row = df_raw.iloc[egta_row_idx]
        vol_row = df_raw.iloc[vol_row_idx] if vol_row_idx is not None else None

        tubes = []
        c_vals = []
        v_vals = []
        n_vals = []

        for col in df_raw.columns:
            cell_tube = tube_row[col]
            if isinstance(cell_tube, str):
                label = cell_tube.strip()
            else:
                label = str(cell_tube).strip() if cell_tube is not None and not pd.isna(cell_tube) else ""

            # On ne garde que les cellules du type "tube X"
            if not re.match(r"^tube\s*\d+", label, flags=re.IGNORECASE):
                continue

            c_val = egta_row[col]
            if pd.isna(c_val):
                # pas de concentration pour ce tube
                continue
            try:
                c_float = float(c_val)
            except Exception:
                continue

            tubes.append(label.title())  # "Tube 1", "Tube 2", ...
            c_vals.append(c_float)

            # Volume, si disponible
            v_float = np.nan
            if vol_row is not None:
                v_val = vol_row[col]
                if not pd.isna(v_val):
                    try:
                        v_float = float(v_val)
                    except Exception:
                        v_float = np.nan
            v_vals.append(v_float)

            # Nombre de moles : C(M) * V(L) = C * (V(mL) / 1000)
            if not np.isnan(v_float):
                n_vals.append(c_float * v_float * 1e-3)
            else:
                n_vals.append(np.nan)

        if not tubes:
            return None

        data = {
            "Tube": tubes,
            "C (EGTA) (M)": c_vals,
        }
        if any(not np.isnan(v) for v in v_vals):
            data["V cuvette (mL)"] = v_vals
            data["n(EGTA) (mol)"] = n_vals

        return pd.DataFrame(data)

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

        # Charger le fichier de correspondance (brut)
        try:
            if file_path.endswith(".csv"):
                df_raw = pd.read_csv(file_path, sep=None, engine="python")
            else:
                df_raw = pd.read_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier de correspondance : {e}")
            self.df_map = None
            return

        # Par défaut on garde tel quel
        self.df_map = df_raw

        # On tente de reconnaître la structure spécifique GC514_spectre_tube :
        # première colonne = "Nom de la manip : ", avec plus bas une ligne "Nom du spectre"
        try:
            df_gc = self._extract_mapping_gc514(df_raw)
        except Exception:
            df_gc = None

        if isinstance(df_gc, pd.DataFrame) and not df_gc.empty:
            self.df_map = df_gc

        self._update_preview(self.df_map)
        self._update_info_label()

    def _update_preview(self, df: pd.DataFrame, nrows: int = 20):  # nrows conservé pour compatibilité, ignoré
        # Afficher TOUT le DataFrame avec défilement et tri
        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str).tolist())
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                val = str(df.iloc[i, j])
                self.table.setItem(i, j, QTableWidgetItem(val))
        self.table.setSortingEnabled(True)

    def _set_root_path(self, path: str):
        if not path:
            return
        idx = self.model.index(path)
        if idx.isValid():
            self.view.setRootIndex(idx)
            self.path_edit.setText(path)

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

    def _go_up(self):
        current_idx = self.view.rootIndex()
        parent_idx = current_idx.parent()
        if parent_idx.isValid():
            self.view.setRootIndex(parent_idx)
            self.path_edit.setText(self.model.filePath(parent_idx))

    def load_metadata(self) -> pd.DataFrame | None:
        return self.combined_df
    def _extract_mapping_gc514(self, df_raw: pd.DataFrame) -> pd.DataFrame | None:
        """
        Détecte et reconstruit un tableau de correspondance de type GC514_spectre_tube.xlsx.

        Structure attendue (comme dans GC514_spectre_tube.xlsx) :
          - la 1ère colonne contient une ligne "Nom du spectre"
          - dans la même ligne, la 2ème colonne contient "Tube" (ou "Tube ")

        On renvoie un DataFrame avec les colonnes :
          Spectrum name, Tube
        """
        # Au minimum 2 colonnes : nom du spectre / tube
        if df_raw.shape[1] < 2:
            return None

        first_col = df_raw.columns[0]
        second_col = df_raw.columns[1]

        # Cherche la ligne d'entête ("Nom du spectre" en première colonne)
        header_idx = None
        for i, val in enumerate(df_raw[first_col].astype(str)):
            if str(val).strip().lower().startswith("nom du spectre"):
                header_idx = i
                break
        if header_idx is None:
            return None

        # Vérifie que la 2ème colonne contient bien "Tube" à cette ligne
        tube_header = str(df_raw.iloc[header_idx, 1])
        tube_header = tube_header.strip().lower()
        if not tube_header.startswith("tube"):
            return None

        # Les données commencent à la ligne suivante
        data = df_raw.iloc[header_idx + 1 :, [0, 1]].copy()
        data.columns = ["Spectrum name", "Tube"]

        # Nettoyage : suppression des lignes totalement vides
        data = data.dropna(how="all")

        # Nettoyage des espaces et conversion en chaînes
        data["Spectrum name"] = data["Spectrum name"].astype(str).str.strip()
        data["Tube"] = data["Tube"].astype(str).str.strip()

        # On ne garde que les lignes avec un vrai nom de spectre (non vide et pas "nan")
        mask_valid = data["Spectrum name"].str.len() > 0
        mask_valid &= ~data["Spectrum name"].str.lower().isin(["nan"])
        data = data[mask_valid]

        # On retourne un tableau propre : une ligne = un spectre
        return data.reset_index(drop=True)