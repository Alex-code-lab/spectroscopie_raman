

# metadata_creator.py

from __future__ import annotations

import pandas as pd

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLineEdit,
)

from PySide6.QtCore import Qt


class TableEditorDialog(QDialog):
    """Boîte de dialogue générique pour éditer un tableau type Excel.

    Le contenu est édité dans un QTableWidget mais peut être converti
    proprement en DataFrame pandas via `to_dataframe()`.
    """

    def __init__(
        self,
        title: str,
        columns: list[str],
        initial_df: pd.DataFrame | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)

        self._columns = list(columns)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(title))

        # --- Table d'édition ---
        self.table = QTableWidget(self)
        self.table.setColumnCount(len(self._columns))
        self.table.setHorizontalHeaderLabels(self._columns)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)

        # Remplissage initial : soit à partir du DataFrame, soit quelques lignes vides
        if initial_df is not None and not initial_df.empty:
            self._load_from_dataframe(initial_df)
        else:
            self.table.setRowCount(8)  # quelques lignes vides par défaut

        layout.addWidget(self.table, 1)

        # --- Boutons de modification du tableau ---
        btns_edit = QHBoxLayout()
        self.btn_add_row = QPushButton("+ Ligne", self)
        self.btn_del_row = QPushButton("− Ligne", self)
        self.btn_add_col = QPushButton("+ Colonne", self)
        self.btn_del_col = QPushButton("− Colonne", self)
        btns_edit.addWidget(self.btn_add_row)
        btns_edit.addWidget(self.btn_del_row)
        btns_edit.addSpacing(20)
        btns_edit.addWidget(self.btn_add_col)
        btns_edit.addWidget(self.btn_del_col)
        btns_edit.addStretch(1)
        layout.addLayout(btns_edit)

        self.btn_add_row.clicked.connect(self._add_row)
        self.btn_del_row.clicked.connect(self._del_row)
        self.btn_add_col.clicked.connect(self._add_col)
        self.btn_del_col.clicked.connect(self._del_col)

        # --- Boutons OK / Annuler ---
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        btn_box.accepted.connect(self._accept_if_not_empty)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.resize(800, 500)

    # ------------------------------------------------------------------
    # Utilitaires internes
    # ------------------------------------------------------------------
    def _load_from_dataframe(self, df: pd.DataFrame) -> None:
        """Charge un DataFrame existant dans le QTableWidget."""
        df = df.copy()
        # Réordonner / compléter selon les colonnes attendues
        for col in self._columns:
            if col not in df.columns:
                df[col] = pd.NA
        df = df[self._columns]

        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(self._columns))
        self.table.setHorizontalHeaderLabels(self._columns)

        for i in range(len(df)):
            for j, col in enumerate(self._columns):
                val = df.iloc[i, j]
                item = QTableWidgetItem("" if pd.isna(val) else str(val))
                self.table.setItem(i, j, item)

    def _add_row(self) -> None:
        self.table.insertRow(self.table.rowCount())

    def _del_row(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        self.table.removeRow(row)

    def _add_col(self) -> None:
        col_count = self.table.columnCount()
        self.table.insertColumn(col_count)
        # Nom de colonne générique
        new_name = f"Col{col_count+1}"
        self._columns.append(new_name)
        self.table.setHorizontalHeaderLabels(self._columns)

    def _del_col(self) -> None:
        col = self.table.currentColumn()
        if col < 0:
            return
        if 0 <= col < len(self._columns):
            self._columns.pop(col)
        self.table.removeColumn(col)
        self.table.setHorizontalHeaderLabels(self._columns)

    def _accept_if_not_empty(self) -> None:
        df = self.to_dataframe()
        if df.empty:
            QMessageBox.warning(self, "Tableau vide", "Le tableau est vide. Ajoutez au moins une ligne avant de valider.")
            return
        self.accept()

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------
    def to_dataframe(self) -> pd.DataFrame:
        """Convertit le contenu de la table en DataFrame pandas.

        Les lignes complètement vides sont supprimées.
        """
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data = {self.table.horizontalHeaderItem(j).text() or f"col{j}": [] for j in range(cols)}

        for i in range(rows):
            row_vals = []
            for j in range(cols):
                item = self.table.item(i, j)
                txt = item.text().strip() if item is not None else ""
                row_vals.append(txt if txt != "" else pd.NA)
            # Ajoute la ligne (on filtrera les lignes vides ensuite)
            for j, key in enumerate(data.keys()):
                data[key].append(row_vals[j])

        df = pd.DataFrame(data)
        # Suppression des lignes entièrement vides
        df = df.dropna(how="all")
        # Normalise quelques colonnes si présentes
        if "Tube" in df.columns:
            df["Tube"] = df["Tube"].astype(str).str.strip()
            df = df[df["Tube"] != ""]
        if "Spectrum name" in df.columns:
            df["Spectrum name"] = df["Spectrum name"].astype(str).str.strip()
        return df.reset_index(drop=True)


class MetadataCreatorWidget(QWidget):
    """Remplaçant de MetadataPicker pour créer les métadonnées directement dans l'application.

    Cette version ne lit plus de fichiers Excel : elle permet de construire
    les deux DataFrames nécessaires au pipeline existant :

    - df_comp : composition des tubes (colonnes typiques : "Tube", "C (EGTA) (M)",
      "V cuvette (mL)", "n(EGTA) (mol)", ...)
    - df_map  : correspondance entre les noms de spectres et les tubes
      (colonnes typiques : "Spectrum name", "Tube", "Sample description", ...)

    Les attributs publics `df_comp` et `df_map` peuvent ensuite être utilisés
    à la place de ceux de `MetadataPickerWidget` pour appeler
    `build_combined_dataframe_from_df` dans le reste du programme.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.df_comp: pd.DataFrame | None = None
        self.df_map: pd.DataFrame | None = None

        layout = QVBoxLayout(self)

        # Champ pour le nom de la manipulation (sera recopié dans le tableau spectres/tubes)
        manip_layout = QHBoxLayout()
        manip_layout.addWidget(QLabel("Nom de la manip :", self))
        self.edit_manip = QLineEdit(self)
        self.edit_manip.setPlaceholderText("ex : GROUPE_20251203")
        manip_layout.addWidget(self.edit_manip)
        layout.addLayout(manip_layout)

        layout.addWidget(QLabel(
            "<b>Création des métadonnées directement dans le programme</b><br>"
            "Utilisez les boutons ci-dessous pour créer / éditer :<ul>"
            "<li>le tableau des volumes (avec concentrations) pour chaque tube</li>"
            "<li>le tableau de correspondance entre les noms de spectres et les tubes</li>"
            "</ul>Les tableaux sont éditables comme dans Excel mais sont stockés en DataFrame pandas.",
            self,
        ))

        # --- Boutons pour ouvrir les éditeurs de tableaux ---
        btns = QHBoxLayout()
        self.btn_edit_comp = QPushButton("Créer / éditer le tableau des volumes", self)
        self.btn_edit_map = QPushButton("Créer / éditer la correspondance spectres ↔ tubes", self)
        self.btn_edit_comp.clicked.connect(self._on_edit_comp_clicked)
        self.btn_edit_map.clicked.connect(self._on_edit_map_clicked)
        btns.addWidget(self.btn_edit_comp)
        btns.addWidget(self.btn_edit_map)
        layout.addLayout(btns)

        # --- État ---
        self.lbl_status_comp = QLabel("Tableau des volumes : non défini", self)
        self.lbl_status_map = QLabel("Tableau de correspondance : non défini", self)
        layout.addWidget(self.lbl_status_comp)
        layout.addWidget(self.lbl_status_map)

        layout.addStretch(1)

    # ------------------------------------------------------------------
    # Gestion du tableau des volumes (avec concentrations)
    # ------------------------------------------------------------------
    def _on_edit_comp_clicked(self) -> None:
        """Ouvre le dialog d'édition pour le tableau des volumes (et concentrations) par tube."""
        # Colonnes exactement dans l'esprit du tableau Excel : réactif + concentration + unités + volumes Tube 1–11
        default_cols = [
            "Réactif",
            "Concentration",
            "Unité",
            "Tube 1",
            "Tube 2",
            "Tube 3",
            "Tube 4",
            "Tube 5",
            "Tube 6",
            "Tube 7",
            "Tube 8",
            "Tube 9",
            "Tube 10",
            "Tube 11",
        ]

        # DataFrame par défaut pré-rempli pour ressembler au tableau fourni (valeurs en µL)
        if self.df_comp is None:
            self.df_comp = pd.DataFrame(
                {
                    "Réactif": [
                        "Echantillon",
                        "Eau de contrôle",
                        "Solution A",
                        "Solution B",
                        "Solution C",
                        "Solution D",
                        "Solution E",
                        "Solution F",
                    ],
                    "Concentration": ["0,5", "", "4", "2", "2", "0,05", "1", "1"],
                    "Unité": ["µM", "", "mM", "µM", "µM", "% en masse", "mM", "mM"],
                    "Tube 1":  [1000,    0, 500,   0,  30, 750, 750,  60],
                    "Tube 2":  [1000,    0, 411,  89,  30, 750, 750,  60],
                    "Tube 3":  [1000,    0, 321, 179,  30, 750, 750,  60],
                    "Tube 4":  [1000,    0, 299, 201,  30, 750, 750,  60],
                    "Tube 5":  [1000,    0, 277, 223,  30, 750, 750,  60],
                    "Tube 6":  [1000,    0, 254, 246,  30, 750, 750,  60],
                    "Tube 7":  [1000,    0, 232, 268,  30, 750, 750,  60],
                    "Tube 8":  [1000,    0, 188, 312,  30, 750, 750,  60],
                    "Tube 9":  [1000,    0,  98, 402,  30, 750, 750,  60],
                    "Tube 10": [1000,    0,  54, 446,  30, 750, 750,  60],
                    "Tube 11": [   0, 1000,  54, 446,  30, 750, 750,  60],
                }
            )

        dlg = TableEditorDialog(
            title="Tableau des volumes par tube",
            columns=default_cols,
            initial_df=self.df_comp,
            parent=self,
        )
        if dlg.exec() == QDialog.Accepted:
            df = dlg.to_dataframe()
            if df.empty:
                QMessageBox.warning(self, "Tableau vide", "Le tableau des volumes est vide.")
                return
            self.df_comp = df
            self.lbl_status_comp.setText(
                f"Tableau des volumes : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
            )

        # ------------------------------------------------------------------
    # Gestion du tableau de correspondance spectres/tubes
    # ------------------------------------------------------------------
    def _on_edit_map_clicked(self) -> None:
        """Ouvre le dialog d'édition pour la correspondance spectres ↔ tubes.

        Le modèle par défaut suit la structure Excel : deux colonnes
        "Nom du spectre" et "Tube", avec les premières lignes utilisées pour
        les métadonnées (Nom de la manip, Date, Lieu, Coordinateur, Opérateur),
        puis une ligne d'en-tête "Nom du spectre" / "Tube" et enfin la liste des spectres.

        Les noms de spectres sont automatiquement générés à partir du nom
        de la manipulation sous la forme NomManip_00, NomManip_01, … (indice à 2 chiffres).
        """
        default_cols = [
            "Nom du spectre",
            "Tube",
        ]

        # Récupérer le nom de la manip saisi sur la page principale
        try:
            manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        except Exception:
            manip_name = ""
        if not manip_name:
            # Valeur par défaut si l'utilisateur n'a encore rien saisi
            manip_name = "MANIP"

        meta_keys = [
            "Nom de la manip :",
            "Date :",
            "Lieu :",
            "Coordinateur :",
            "Opérateur :",
        ]

        # DataFrame par défaut si aucun tableau n'a encore été créé
        if self.df_map is None:
            # 12 spectres par défaut (00 à 11)
            spec_names = [f"{manip_name}_{i:02d}" for i in range(12)]
            self.df_map = pd.DataFrame(
                {
                    "Nom du spectre": [
                        "Nom de la manip :",
                        "Date :",
                        "Lieu :",
                        "Coordinateur :",
                        "Opérateur :",
                        "",
                        "Nom du spectre",
                        *spec_names,
                    ],
                    "Tube": [
                        manip_name,
                        "02/12/2025",
                        "Campus de la transition",
                        "Marcelline",
                        "Dominique",
                        "",
                        "Tube",
                        "Tube BRB",
                        "Tube 1",
                        "Tube 2",
                        "Tube 3",
                        "Tube 4",
                        "Tube 5",
                        "Tube 6",
                        "Tube 7",
                        "Tube 8",
                        "Tube 9",
                        "Tube 10",
                        "Tube MQW",
                    ],
                }
            )
        else:
            # Mettre à jour la ligne "Nom de la manip :" si elle existe
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
                mask_manip = col_ns == "Nom de la manip :"
                if mask_manip.any():
                    self.df_map.loc[mask_manip, "Tube"] = manip_name
            except Exception:
                pass

            # Regénérer les noms de spectres à partir du nom de la manip
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
                # Lignes qui correspondent aux vrais mappings (pas les méta, pas la ligne d'en-tête ni les vides)
                mask_mapping = ~col_ns.isin(meta_keys + ["", "Nom du spectre"])
                idx_mapping = list(self.df_map.index[mask_mapping])
                for k, idx in enumerate(idx_mapping):
                    self.df_map.loc[idx, "Nom du spectre"] = f"{manip_name}_{k:02d}"
            except Exception:
                # Si la structure est inattendue, on ne bloque pas l'utilisateur.
                pass

        dlg = TableEditorDialog(
            title="Tableau de correspondance spectres ↔ tubes",
            columns=default_cols,
            initial_df=self.df_map,
            parent=self,
        )
        if dlg.exec() == QDialog.Accepted:
            df = dlg.to_dataframe()
            if df.empty:
                QMessageBox.warning(self, "Tableau vide", "Le tableau de correspondance est vide.")
                return
            self.df_map = df
            self.lbl_status_map.setText(
                f"Tableau de correspondance : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
            )

            # Synchroniser le champ "Nom de la manip" avec ce qui est écrit dans la ligne correspondante
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
                mask = col_ns == "Nom de la manip :"
                if mask.any():
                    val = str(self.df_map.loc[mask, "Tube"].iloc[0])
                    if hasattr(self, "edit_manip"):
                        self.edit_manip.setText(val)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Méthode utilitaire optionnelle
    # ------------------------------------------------------------------
    def build_merged_metadata(self) -> pd.DataFrame:
        """Construit un DataFrame unique avec une ligne par spectre.

        Cette méthode fusionne `df_map` et `df_comp` sur la colonne "Tube".

        Pour `df_map`, on accepte le modèle Excel suivant :
        - deux colonnes "Nom du spectre" et "Tube" ;
        - les premières lignes peuvent contenir des méta-informations
          ("Nom de la manip :", "Date :", "Lieu :", "Coordinateur :", "Opérateur :");
        - une ligne d'en-tête interne "Nom du spectre" / "Tube" ;
        - la suite correspond aux vrais couples (nom de spectre, tube).

        On extrait automatiquement un DataFrame de mapping avec colonnes
        "Spectrum name" et "Tube", ainsi que des colonnes optionnelles
        pour les méta-informations (mêmes valeurs répétées pour chaque ligne).
        """
        if self.df_comp is None or self.df_map is None:
            raise ValueError("Les deux tableaux (volumes et correspondance) doivent être définis.")

        df_comp = self.df_comp.copy()
        df_map_raw = self.df_map.copy()

        # Vérifier la présence de la colonne Tube dans df_comp
        if "Tube" not in df_comp.columns:
            raise ValueError("Le tableau des volumes doit contenir une colonne 'Tube'.")

        # ----- Normalisation / extraction du mapping spectre ↔ tube -----
        if "Spectrum name" in df_map_raw.columns:
            # Cas simple : on a déjà la bonne colonne
            df_map = df_map_raw.copy()

        elif "Nom du spectre" in df_map_raw.columns:
            meta_keys = [
                "Nom de la manip :",
                "Date :",
                "Lieu :",
                "Coordinateur :",
                "Opérateur :",
            ]

            # Extraire les méta-informations globales
            meta_values: dict[str, str] = {}
            col_ns = df_map_raw["Nom du spectre"].astype(str).str.strip()
            for key in meta_keys:
                mask = col_ns == key
                if mask.any():
                    val = (
                        df_map_raw.loc[mask, "Tube"]
                        .astype(str)
                        .iloc[0]
                    )
                    # On crée une colonne sans les deux-points
                    clean_col = key.replace(":", "").strip()
                    meta_values[clean_col] = val

            # Lignes correspondant au véritable mapping spectre ↔ tube
            col_ns = df_map_raw["Nom du spectre"].astype(str).str.strip()
            mask_mapping = ~col_ns.isin(meta_keys + ["", "Nom du spectre"])
            df_mapping = df_map_raw.loc[mask_mapping, ["Nom du spectre", "Tube"]].copy()

            if df_mapping.empty:
                raise ValueError("Aucun couple (Nom du spectre, Tube) exploitable trouvé dans le tableau de correspondance.")

            # Créer la colonne Spectrum name attendue par le pipeline
            df_mapping = df_mapping.rename(columns={"Nom du spectre": "Spectrum name"})
            df_mapping["Spectrum name"] = df_mapping["Spectrum name"].astype(str).str.strip()
            df_mapping["Tube"] = df_mapping["Tube"].astype(str).str.strip()

            # Ajouter les méta-informations en colonnes si disponibles
            for col_name, value in meta_values.items():
                df_mapping[col_name] = value

            df_map = df_mapping.reset_index(drop=True)

        else:
            raise ValueError(
                "Le tableau de correspondance doit contenir soit une colonne 'Spectrum name', "
                "soit une colonne 'Nom du spectre'."
            )

        # Harmonisation minimale : enlever l'extension .txt si l'utilisateur l'a mise
        if "Spectrum name" in df_map.columns:
            df_map["Spectrum name"] = (
                df_map["Spectrum name"]
                .astype(str)
                .str.strip()
                .str.replace(".txt", "", regex=False)
            )

        # ----- Jointure finale avec df_comp -----
        if "Tube" not in df_map.columns:
            raise ValueError("Le tableau de correspondance doit contenir une colonne 'Tube'.")

        merged_meta = df_map.merge(df_comp, on="Tube", how="left")
        return merged_meta