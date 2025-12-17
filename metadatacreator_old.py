

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
    QDateEdit,
    QFileDialog,
)

from PySide6.QtCore import Qt, QDate


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
        self.table.itemChanged.connect(self._auto_resize_columns)

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
        self.table.resizeColumnsToContents()


    def _auto_resize_columns(self) -> None:
        """Ajuste automatiquement la largeur des colonnes au contenu."""
        self.table.resizeColumnsToContents()

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

        self.df_comp: pd.DataFrame | None = None # tableau de composition des tubes
        self.df_map: pd.DataFrame | None = None # tableau de correspondance spectres ↔ tubes
        self.df_conc: pd.DataFrame | None = None  # tableau de concentrations calculé à partir des volumes
        # Indique que la correspondance spectres↔tubes doit être revue (champs d’en-tête modifiés)
        self.map_dirty: bool = False

        layout = QVBoxLayout(self)

        # Champs d'en-tête pour la manipulation : nom, date, lieu, coordinateur, opérateur
        header_layout = QVBoxLayout()

        # Ligne 1 : nom de la manip + date (avec calendrier)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Nom de la manip :", self))
        self.edit_manip = QLineEdit(self)
        self.edit_manip.setPlaceholderText("ex : GROUPE_20251203")
        row1.addWidget(self.edit_manip)
        # Synchroniser automatiquement le nom de la manip avec df_map
        self.edit_manip.textChanged.connect(self._on_manip_name_changed)

        row1.addWidget(QLabel("Date :", self))
        self.edit_date = QDateEdit(self)
        self.edit_date.setCalendarPopup(True)
        self.edit_date.setDisplayFormat("dd/MM/yyyy")
        self.edit_date.setDate(QDate.currentDate())
        row1.addWidget(self.edit_date)
        self.edit_date.dateChanged.connect(self._on_header_field_changed)

        header_layout.addLayout(row1)

        # Ligne 2 : lieu, coordinateur, opérateur
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Lieu :", self))
        self.edit_location = QLineEdit(self)
        self.edit_location.setPlaceholderText("ex : Laboratoire / Terrain")
        row2.addWidget(self.edit_location)
        self.edit_location.textChanged.connect(self._on_header_field_changed)

        row2.addWidget(QLabel("Coordinateur :", self))
        self.edit_coordinator = QLineEdit(self)
        self.edit_coordinator.setPlaceholderText("ex : Nom du coordinateur")
        row2.addWidget(self.edit_coordinator)
        self.edit_coordinator.textChanged.connect(self._on_header_field_changed)

        row2.addWidget(QLabel("Opérateur :", self))
        self.edit_operator = QLineEdit(self)
        self.edit_operator.setPlaceholderText("ex : Nom de l'opérateur")
        row2.addWidget(self.edit_operator)
        self.edit_operator.textChanged.connect(self._on_header_field_changed)

        header_layout.addLayout(row2)
        layout.addLayout(header_layout)

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
        self.btn_show_conc = QPushButton("Afficher le tableau des concentrations (info)", self)

        # Boutons enregistrer / charger sur une seconde ligne
        self.btn_save_meta = QPushButton("Enregistrer les métadonnées…", self)
        self.btn_load_meta = QPushButton("Charger des métadonnées…", self)

        self.btn_edit_comp.clicked.connect(self._on_edit_comp_clicked)
        self.btn_edit_map.clicked.connect(self._on_edit_map_clicked)
        self.btn_show_conc.clicked.connect(self._on_show_conc_clicked)
        self.btn_save_meta.clicked.connect(self._on_save_metadata_clicked)
        self.btn_load_meta.clicked.connect(self._on_load_metadata_clicked)

        # Première ligne
        btns.addWidget(self.btn_edit_comp)
        btns.addWidget(self.btn_edit_map)
        btns.addWidget(self.btn_show_conc)

        layout.addLayout(btns)

        # Seconde ligne
        btns2 = QHBoxLayout()
        btns2.addWidget(self.btn_save_meta)
        btns2.addWidget(self.btn_load_meta)
        layout.addLayout(btns2)

        # Couleurs des boutons selon l'état (rouge si non défini, vert si prêt)
        self._refresh_button_states()

        # --- État ---
        self.lbl_status_comp = QLabel("Tableau des volumes : non défini", self)
        self.lbl_status_map = QLabel("Tableau de correspondance : non défini", self)
        layout.addWidget(self.lbl_status_comp)
        layout.addWidget(self.lbl_status_map)

        layout.addStretch(1)

    def _refresh_button_states(self) -> None:
        """Met à jour la couleur des boutons selon l'état des DataFrames.

        - rouge si le tableau n'est pas défini
        - vert si le tableau est prêt (df non vide)
        """
        red = "background-color: #d9534f; color: white; font-weight: 600;"
        green = "background-color: #5cb85c; color: white; font-weight: 600;"
        neutral = ""  # laisse le style par défaut

        comp_ready = self.df_comp is not None and isinstance(self.df_comp, pd.DataFrame) and not self.df_comp.empty
        map_ready = (
            self.df_map is not None
            and isinstance(self.df_map, pd.DataFrame)
            and not self.df_map.empty
            and not self.map_dirty
        )
        conc_ready = self.df_conc is not None and isinstance(self.df_conc, pd.DataFrame) and not self.df_conc.empty

        # Bouton volumes
        if comp_ready:
            self.btn_edit_comp.setStyleSheet(green)
        else:
            self.btn_edit_comp.setStyleSheet(red)

        # Bouton correspondance spectres↔tubes
        if map_ready:
            self.btn_edit_map.setStyleSheet(green)
        else:
            self.btn_edit_map.setStyleSheet(red)

        # Bouton concentrations : vert si le tableau des volumes existe (calcul possible),
        # sinon rouge. (Optionnel : vert plus foncé si conc_ready.)
        if comp_ready:
            self.btn_show_conc.setStyleSheet(green if conc_ready else green)
        else:
            self.btn_show_conc.setStyleSheet(red)

        # Mettre à jour les tooltips (optionnel mais utile)
        self.btn_edit_comp.setToolTip("Vert = tableau des volumes défini" if comp_ready else "Rouge = à créer / éditer")
        self.btn_edit_map.setToolTip("Vert = correspondance définie" if map_ready else "Rouge = à créer / éditer")
        if comp_ready:
            self.btn_show_conc.setToolTip("Calculable (volumes définis)" if not conc_ready else "Concentrations calculées")
        else:
            self.btn_show_conc.setToolTip("Rouge = définir d'abord les volumes")
    def _on_header_field_changed(self, *args) -> None:
        """Marque la correspondance spectres↔tubes comme à revoir quand un champ d’en-tête change."""
        # On ne marque dirty que si un df_map existe déjà (sinon c'est déjà rouge)
        if self.df_map is None or self.df_map.empty:
            return
        self.map_dirty = True
        self._refresh_button_states()

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
                        "Contrôle",
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

            # Sauvegarde du tableau des volumes
            self.df_comp = df
            self.lbl_status_comp.setText(
                f"Tableau des volumes : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
            )

            # IMPORTANT : recalcul automatique des concentrations (pas besoin de valider un 2e tableau)
            try:
                self.df_conc = self._compute_concentration_table()
            except Exception as e:
                # On garde df_comp mais on informe l'utilisateur que les concentrations ne sont pas calculables
                self.df_conc = None
                QMessageBox.warning(
                    self,
                    "Concentrations non calculées",
                    "Le tableau des volumes a été enregistré, mais le tableau des concentrations n'a pas pu être calculé.\n"
                    "Vous pourrez corriger le tableau des volumes puis réessayer.\n\n"
                    f"Détail : {e}",
                )

            # Toute modification de volumes invalide la correspondance (dépendance possible pour l'analyse)
            # et met à jour les couleurs
            self._refresh_button_states()

    def _on_show_conc_clicked(self) -> None:
        """Affiche le tableau des concentrations déjà calculé (calcul à la demande uniquement si manquant)."""
        if self.df_comp is None or not isinstance(self.df_comp, pd.DataFrame) or self.df_comp.empty:
            QMessageBox.warning(
                self,
                "Volumes manquants",
                "Le tableau des volumes n'est pas défini. Créez d'abord / éditez le tableau des volumes.",
            )
            return

        # Normalement df_conc est calculé automatiquement quand on valide le tableau des volumes.
        # On garde un calcul à la demande uniquement si df_conc est manquant.
        if self.df_conc is None or not isinstance(self.df_conc, pd.DataFrame) or self.df_conc.empty:
            try:
                self.df_conc = self._compute_concentration_table()
            except Exception as e:
                QMessageBox.critical(self, "Erreur calcul", f"Impossible de calculer le tableau des concentrations :\n{e}")
                return

        # Afficher dans un éditeur (lecture/édition tolérée, mais ce tableau est normalement calculé)
        dlg = TableEditorDialog(
            title="Tableau des concentrations (calculé)",
            columns=list(self.df_conc.columns),
            initial_df=self.df_conc,
            parent=self,
        )
        dlg.exec()

        # Mettre à jour l'état des boutons
        self._refresh_button_states()


    def _on_edit_map_clicked(self) -> None:
        """
        Ouvre le dialog d'édition pour la correspondance spectres ↔ tubes.

        Le modèle par défaut suit la structure Excel d'origine avec :
        - 5 lignes d'en-tête (Nom de la manip, Date, Lieu, Coordinateur, Opérateur),
        - une ligne vide,
        - une ligne d'en-tête interne "Nom du spectre" / "Tube",
        - puis les lignes de correspondance NomSpectre ↔ Tube.

        Les noms de spectres sont automatiquement générés à partir du nom de la manip
        sous la forme NomManip_00, NomManip_01, … (indice à 2 chiffres) lors de la
        première création du tableau.
        """
        default_cols = ["Nom du spectre", "Tube"]

        # 1) Récupération des informations d'en-tête depuis l'UI
        try:
            manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        except Exception:
            manip_name = ""

        if not manip_name:
            QMessageBox.warning(
                self,
                "Nom de la manip manquant",
                "Veuillez renseigner le nom de la manip avant de créer la correspondance spectres ↔ tubes.",
            )
            return

        # Date (format dd/MM/yyyy)
        date_str = ""
        if hasattr(self, "edit_date") and self.edit_date is not None:
            try:
                date_str = self.edit_date.date().toString("dd/MM/yyyy")
            except Exception:
                date_str = ""

        # Lieu
        location = ""
        if hasattr(self, "edit_location") and self.edit_location is not None:
            try:
                location = self.edit_location.text().strip()
            except Exception:
                location = ""

        # Coordinateur
        coordinator = ""
        if hasattr(self, "edit_coordinator") and self.edit_coordinator is not None:
            try:
                coordinator = self.edit_coordinator.text().strip()
            except Exception:
                coordinator = ""

        # Opérateur
        operator = ""
        if hasattr(self, "edit_operator") and self.edit_operator is not None:
            try:
                operator = self.edit_operator.text().strip()
            except Exception:
                operator = ""

        # Si df_map existe, on met à jour les lignes d'en-tête pour refléter l'UI
        if self.df_map is not None and not self.df_map.empty:
            try:
                df_tmp = self.df_map.copy()
                col_ns_tmp = df_tmp["Nom du spectre"].astype(str).str.strip()
                header_values = {
                    "Nom de la manip": manip_name,
                    "Nom de la manip :": manip_name,
                    "Date": date_str,
                    "Date :": date_str,
                    "Lieu": location,
                    "Lieu :": location,
                    "Coordinateur": coordinator,
                    "Coordinateur :": coordinator,
                    "Opérateur": operator,
                    "Opérateur :": operator,
                }
                for idx in df_tmp.index:
                    label = str(df_tmp.at[idx, "Nom du spectre"]).strip()
                    if label in header_values:
                        df_tmp.at[idx, "Tube"] = header_values[label]
                self.df_map = df_tmp
            except Exception:
                pass

        # 2) Lignes d'en-tête construites à partir de l'UI (toujours recalculées)
        header_rows = [
            ["Nom de la manip", manip_name],
            ["Date", date_str],
            ["Lieu", location],
            ["Coordinateur", coordinator],
            ["Opérateur", operator],
        ]
        df_header = pd.DataFrame(header_rows, columns=default_cols)

        blank_row = pd.DataFrame([["", ""]], columns=default_cols)
        internal_header = pd.DataFrame([["Nom du spectre", "Tube"]], columns=default_cols)

        # 3) Récupération de la partie "correspondance" existante si elle a déjà été éditée
        mapping_rows = None
        if self.df_map is not None and not self.df_map.empty:
            existing = self.df_map.copy()

            # On cherche la ligne d'en-tête interne "Nom du spectre" / "Tube"
            hdr_idx = None
            for idx, row in existing.iterrows():
                left = str(row.get("Nom du spectre", "")).strip()
                right = str(row.get("Tube", "")).strip()
                if left == "Nom du spectre" and right == "Tube":
                    hdr_idx = idx
                    break

            if hdr_idx is not None:
                # Tout ce qui est après cette ligne correspond aux couples NomSpectre ↔ Tube
                mapping_rows = existing.loc[hdr_idx + 1 :].reset_index(drop=True)
            else:
                # Pas d'en-tête interne clairement identifiable :
                # on considère tout comme mapping existant
                mapping_rows = existing.reset_index(drop=True)

        # Si aucune correspondance existante, on génère un brouillon par défaut
        if mapping_rows is None:
            # Tubes par défaut : Tube MQW, Tube BRB, puis Tube 1..11
            tube_labels = ["Contrôle BRB"] + [f"Tube {i}" for i in range(1, 11)] + ["Contrôle"]
            mapping_rows = pd.DataFrame(
                {
                    "Nom du spectre": [f"{manip_name}_{i:02d}" for i in range(len(tube_labels))],
                    "Tube": tube_labels,
                },
                columns=default_cols,
            )

        # 4) DataFrame initial passé au TableEditorDialog
        initial_df = pd.concat(
            [df_header, blank_row, internal_header, mapping_rows],
            ignore_index=True,
        )

        # 5) Ouverture du dialogue d'édition
        dlg = TableEditorDialog(
            title="Correspondance spectres ↔ tubes",
            columns=default_cols,
            initial_df=initial_df,
            parent=self,
        )
        if dlg.exec() != QDialog.Accepted:
            return

        # 6) Récupération du tableau édité et MISE À JOUR forcée des lignes d'en-tête
        df = dlg.to_dataframe()
        if df.empty:
            QMessageBox.warning(self, "Tableau vide", "Le tableau de correspondance est vide.")
            return

        # On force les valeurs des lignes d'en-tête à partir de l'UI, pour éviter
        # tout décalage : le clic sur le bouton reflète immédiatement les derniers champs saisis.
        header_values = {
            "Nom de la manip": manip_name,
            "Date": date_str,
            "Lieu": location,
            "Coordinateur": coordinator,
            "Opérateur": operator,
        }
        for idx in df.index:
            label = str(df.at[idx, "Nom du spectre"]).strip()
            if label in header_values:
                df.at[idx, "Tube"] = header_values[label]

        self.df_map = df
        self.map_dirty = False
        self.lbl_status_map.setText(
            f"Tableau de correspondance : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
        )
        self._refresh_button_states()

    def _on_manip_name_changed(self, text: str) -> None:
        """
        Slot appelé automatiquement quand le champ 'Nom de la manip' change.
        Met à jour df_map en mémoire (ligne 'Nom de la manip' + noms des spectres).
        """
        new_name = (text or "").strip()
        if not new_name:
            return
        self._apply_manip_name_to_df_map(new_name)
        # Un changement de nom de manip impacte directement la correspondance (noms de spectres)
        self.map_dirty = True
        self._refresh_button_states()

    def _apply_manip_name_to_df_map(self, manip_name: str) -> None:
        """
        Applique le nom de la manip à df_map :
        - met à jour la ligne 'Nom de la manip'
        - régénère tous les noms de spectres sous la forme NomManip_XX
        sans modifier l'affectation des tubes.
        """
        if self.df_map is None or self.df_map.empty:
            return
        if "Nom du spectre" not in self.df_map.columns or "Tube" not in self.df_map.columns:
            return

        df = self.df_map.copy()
        col_ns = df["Nom du spectre"].astype(str).str.strip()

        # 1) Mettre à jour la ligne d'en-tête "Nom de la manip"
        mask_manip = col_ns.isin(["Nom de la manip", "Nom de la manip :"])
        if mask_manip.any():
            df.loc[mask_manip, "Tube"] = manip_name

        # 2) Trouver la ligne d'en-tête interne "Nom du spectre"
        header_mask = col_ns == "Nom du spectre"
        if not header_mask.any():
            self.df_map = df
            return

        last_header_idx = header_mask[header_mask].index[-1]

        # 3) Régénérer les noms des spectres après cette ligne
        mapping_idx = df.index[df.index > last_header_idx].tolist()
        for i, ridx in enumerate(mapping_idx):
            df.at[ridx, "Nom du spectre"] = f"{manip_name}_{i:02d}"

        self.df_map = df

        # Mettre à jour le label d'état si présent
        if hasattr(self, "lbl_status_map"):
            self.lbl_status_map.setText(
                f"Tableau de correspondance : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
            )

    # ------------------------------------------------------------------
    # Calcul du tableau des concentrations à partir des volumes
    # ------------------------------------------------------------------
    def _compute_concentration_table(self) -> pd.DataFrame:
        """
        Calcule le tableau des concentrations dans chaque cuvette à partir du tableau des volumes.

        Idée physique :
        - Pour chaque tube j, on calcule le volume total V_tot_j (somme des volumes de tous les réactifs).
        - Pour chaque réactif i et chaque tube j, on calcule la concentration finale :
              C_ij = C_stock_i * V_ij / V_tot_j

          où :
            - C_stock_i : concentration du stock du réactif i (en M = mol/L),
            - V_ij : volume de ce réactif i dans le tube j (en µL),
            - V_tot_j : volume total du tube j (en µL).

          Comme V_ij et V_tot_j sont dans la même unité (µL), le ratio V_ij / V_tot_j est sans dimension,
          donc C_ij a bien les mêmes unités que C_stock_i (M).
        """
        # Sécurité : s'assurer qu'on a bien un tableau de volumes
        if self.df_comp is None or self.df_comp.empty:
            raise ValueError("Le tableau des volumes n'est pas encore défini.")

        # Copie de travail pour ne pas modifier self.df_comp
        df = self.df_comp.copy()

        # Colonnes qui correspondent aux tubes (ex : "Tube 1", "Tube 2", ...)
        tube_cols = [c for c in df.columns if c.lower().startswith("tube")]
        if not tube_cols:
            raise ValueError("Le tableau des volumes ne contient pas de colonnes 'Tube X'.")

        # Conversion des volumes en µL vers du numérique (float)
        # - pd.to_numeric(..., errors="coerce") transforme ce qui n'est pas convertible en NaN
        vol = df[tube_cols].apply(lambda col: pd.to_numeric(col, errors="coerce"))
        # On remplace tous les NaN par 0 µL (cas où la case est vide dans le tableau)
        vol = vol.fillna(0.0)

        # Volume total par tube (µL) = somme des volumes de tous les réactifs pour chaque colonne "Tube X"
        total_ul = vol.sum(axis=0)
        # Remplacer 0 par NaN pour éviter les divisions par zéro plus bas
        total_ul = total_ul.replace(0, pd.NA)

        # ---------- Helpers pour parser les concentrations stock ----------

        def _to_float(val):
            """
            Convertit une valeur de concentration écrite en texte (ex : '0,5') en float Python (0.5).
            - Gère les virgules françaises en les remplaçant par un point.
            - Renvoie NaN si ce n'est pas convertible.
            """
            if pd.isna(val):
                return float("nan")
            s = str(val).strip().replace(",", ".")
            try:
                return float(s)
            except ValueError:
                return float("nan")

        def _conc_to_M(val, unit):
            """
            Convertit une concentration (val, unité) en M (mol/L) si possible.

            Règles:
            - 'M', 'mol/L', 'mol.l-1'    -> M (inchangé)
            - 'mM', 'mmol/L', 'mmol.l-1' -> M en divisant par 1000
            - 'µM', 'um', etc.           -> M en divisant par 1 000 000
            - '% en masse'               -> traité ici comme "pas exploitable" -> NaN
            - Autres unités              -> renvoie la valeur brute v (comportement par défaut)
            """
            v = _to_float(val)
            if pd.isna(v):
                return float("nan")

            # Normalisation de l'unité : on enlève les espaces, met en minuscule
            u = "" if pd.isna(unit) else str(unit).strip().replace(" ", "").lower()

            # Déjà en M (mol/L)
            if u in ("m", "mol/l", "mol.l-1"):
                return v
            # mM -> M
            if u in ("mm", "mmol/l", "mmol.l-1"):
                return v * 1e-3
            # µM -> M
            if u in ("µm", "um", "µmol/l", "umol/l"):
                return v * 1e-6
            # Cas particulier : % en masse → on choisit ici de ne PAS l'utiliser dans le calcul
            if "%enmasse" in u:
                return float("nan")  # ou 0.0 selon ta stratégie

            # Pour tout ce qui ne matche pas, on renvoie la valeur brute
            # (comportement par défaut ; à ajuster si nécessaire)
            return v

        # Liste de dictionnaires qui serviront à construire le DataFrame de sortie
        out_rows: list[dict] = []

        # 1) Première ligne : Volume total en µL pour chaque tube
        row_v = {"Réactif": "V cuvette (µL)"}
        for col in tube_cols:
            v_ul = total_ul[col]
            # Si total_ul est NaN, on met NaN ; sinon, on force en float
            row_v[col] = float(v_ul) if not pd.isna(v_ul) else float("nan")
        out_rows.append(row_v)

        # 2) Lignes de concentrations pour chaque réactif
        for idx, row in df.iterrows():
            # Nom "lisible" du réactif (sinon "Ligne i")
            name = row.get("Réactif", f"Ligne {idx}")
            # Conversion de la concentration stock en M
            c_stock = _conc_to_M(row.get("Concentration"), row.get("Unité"))

            # Dictionnaire pour cette ligne de sortie
            out = {"Réactif": name}

            if pd.isna(c_stock):
                # Si on n'a pas de concentration stock exploitable :
                # on met 0 partout (ou NaN si tu préfères signaler "non défini")
                for col in tube_cols:
                    out[col] = 0.0
            else:
                # Sinon, on calcule la concentration finale dans chaque tube
                for col in tube_cols:
                    v_ij = vol.at[idx, col]   # volume de ce réactif dans ce tube (µL)
                    v_tot = total_ul[col]     # volume total du tube (µL)
                    if pd.isna(v_tot) or v_tot == 0:
                        out[col] = float("nan")
                    else:
                        # Formule : C_final_ij = C_stock_i * (V_ij / V_tot_j)
                        # -> C_stock en M, V_ij et V_tot_j en µL -> ratio sans dimension
                        # -> C_final_ij en M
                        out[col] = float(c_stock) * (float(v_ij) / float(v_tot))

            out_rows.append(out)

        # Construction finale du DataFrame de concentrations
        conc_df = pd.DataFrame(out_rows, columns=["Réactif"] + tube_cols)

        # ---- Formatage : écriture scientifique à 2 décimales ----
        for col in tube_cols:
            conc_df[col] = conc_df[col].apply(
                lambda v: f"{float(v):.2e}" if pd.notna(v) else ""
            )

        return conc_df
    
    def _on_show_conc_clicked(self) -> None:
        """Calcule et affiche le tableau des concentrations lié au tableau des volumes."""
        try:
            conc_df = self._compute_concentration_table()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erreur de calcul",
                f"Impossible de calculer le tableau des concentrations :\n{e}",
            )
            return

        self.df_conc = conc_df
        self._refresh_button_states()

        # Affichage dans une boîte de dialogue type Excel
        cols = ["Réactif"] + [c for c in conc_df.columns if c != "Réactif"]
        dlg = TableEditorDialog(
            title="Tableau des concentrations dans les cuvettes (M)",
            columns=cols,
            initial_df=conc_df,
            parent=self,
        )
        dlg.exec()


    # ------------------------------------------------------------------
    # Méthode utilitaire optionnelle
    # ------------------------------------------------------------------
    def build_merged_metadata(self) -> pd.DataFrame:
        """
        Construit un DataFrame de métadonnées à partir des tableaux créés
        dans l'onglet Métadonnées :

        - df_map  : correspondance spectres ↔ tubes (Nom du spectre / Tube)
        - df_conc : tableau des concentrations par tube (réactifs en lignes, tubes en colonnes)

        Le DataFrame retourné contient au minimum :
            - "Spectrum name"      : nom du spectre (pour la jointure avec les .txt)
            - "Tube"               : identifiant du tube (ex : "Tube 1", "Tube BRB", "Tube MQW")
            - "Sample description" : description textuelle pour les tracés (ex : "Cuvette BRB")
        Et, si df_conc est défini :
            - une colonne par réactif (pivot de df_conc)
            - des colonnes dérivées "[EGTA] (M)" et "[EGTA] (mM)" si un réactif contenant "EGTA" est trouvé.
        """
        if self.df_map is None or self.df_map.empty:
            raise ValueError("Le tableau de correspondance spectres ↔ tubes n'est pas encore défini.")

        df_map = self.df_map.copy()

        # Vérifier la présence des colonnes attendues
        if "Nom du spectre" not in df_map.columns or "Tube" not in df_map.columns:
            raise ValueError("Le tableau de correspondance doit contenir les colonnes 'Nom du spectre' et 'Tube'.")

        # Normaliser les textes
        df_map["Nom du spectre"] = df_map["Nom du spectre"].astype(str)
        df_map["Tube"] = df_map["Tube"].astype(str)

        # Repérer la ligne d'en-tête interne "Nom du spectre" / "Tube"
        col_ns = df_map["Nom du spectre"].str.strip()
        header_mask = col_ns.eq("Nom du spectre")
        if header_mask.any():
            # On garde uniquement les lignes situées APRÈS cette ligne d'en-tête
            last_header_idx = header_mask[header_mask].index[-1]
            df_map = df_map.loc[df_map.index > last_header_idx].copy()

        # Nettoyage : retirer les lignes vides ou incomplètes
        df_map["Nom du spectre"] = df_map["Nom du spectre"].str.strip()
        df_map["Tube"] = df_map["Tube"].str.strip()
        df_map = df_map[(df_map["Nom du spectre"] != "") & (df_map["Tube"] != "")].copy()

        # Base des métadonnées : Spectrum name + Tube
        meta = df_map.rename(columns={"Nom du spectre": "Spectrum name"})

        # Construire une "Sample description" à partir du nom de tube
        tube_to_desc = {
            "Contrôle BRB": "Contrôle BRB",
            "Contrôle": "Contrôle",
        }
        meta["Sample description"] = meta["Tube"].map(tube_to_desc).fillna(meta["Tube"])

        # Si on dispose du tableau de concentrations, l'exploiter pour ajouter des colonnes
        if self.df_conc is not None and isinstance(self.df_conc, pd.DataFrame) and not self.df_conc.empty:
            conc = self.df_conc.copy()
            if "Réactif" in conc.columns:
                # Colonnes de tubes dans df_conc (ex : "Tube 1", "Tube 2", ...)
                tube_cols = [c for c in conc.columns if c != "Réactif"]
                if tube_cols:
                    # Mise au format long : (Réactif, Tube, C_M)
                    long = conc.melt(
                        id_vars="Réactif",
                        value_vars=tube_cols,
                        var_name="Tube",
                        value_name="C_M",
                    )
                    # Pivot : un réactif par colonne, indexé par "Tube"
                    conc_wide = long.pivot(index="Tube", columns="Réactif", values="C_M")
                    # Forcer toutes les colonnes de conc_wide en numérique (au cas où df_conc contient des chaînes)
                    conc_wide = conc_wide.apply(pd.to_numeric, errors="coerce")

                    # Joindre ces colonnes par tube
                    meta = meta.merge(conc_wide, on="Tube", how="left")

                    # Tenter de détecter une colonne correspondant au titrant.
                    # On considère comme titrant :
                    #  - tout réactif dont le nom contient "titrant"
                    #  - ou, par compatibilité descendante, "egta"
                    #  - ou les variantes de "solution b" utilisées dans ton protocole.
                    titrant_candidates: list[str] = []
                    for col in conc_wide.columns:
                        name_low = str(col).strip().lower()
                        # 1) Cas général : le nom du réactif contient explicitement "titrant" ou "egta"
                        if ("titrant" in name_low) or ("egta" in name_low):
                            titrant_candidates.append(col)
                        # 2) Cas spécifique à ton protocole : la solution B est le titrant
                        elif name_low in {"solution b", "solutionb", "sol b", "solb", "b"}:
                            titrant_candidates.append(col)

                    # Si rien n'a été trouvé, titrant_candidates restera vide et on n'ajoutera
                    # pas les colonnes [titrant].
                    if titrant_candidates:
                        titrant_col = titrant_candidates[0]
                        # S'assurer que la colonne titrant est bien numérique
                        meta[titrant_col] = pd.to_numeric(meta[titrant_col], errors="coerce")
                        # Colonnes explicites pour l'analyse : [titrant] (M) et [titrant] (mM)
                        meta["[titrant] (M)"] = meta[titrant_col]
                        meta["[titrant] (mM)"] = meta[titrant_col] * 1e3

        # Exclure explicitement les tubes de type "Tube BRB" et "Tube MQW"
        # des métadonnées utilisées pour l'assemblage (tracés, analyse)
        meta = meta[~meta["Tube"].isin(["Tube BRB"])].copy()

        # Réindexer proprement
        meta = meta.reset_index(drop=True)

        return meta
    def _update_header_fields_from_df_map(self) -> None:
        """Met à jour les champs Nom de la manip, Date, Lieu, Coordinateur, Opérateur
        à partir des premières lignes du tableau de correspondance df_map, si possible.
        """
        if self.df_map is None or self.df_map.empty:
            return
        if "Nom du spectre" not in self.df_map.columns or "Tube" not in self.df_map.columns:
            return

        df = self.df_map
        col_ns = df["Nom du spectre"].astype(str).str.strip()
        col_tube = df["Tube"].astype(str)

        def _value_for(label: str) -> str:
            mask = col_ns == label
            if mask.any():
                return col_tube[mask].iloc[0].strip()
            return ""

        # Nom de la manip
        name_val = _value_for("Nom de la manip :")
        if name_val and hasattr(self, "edit_manip"):
            self.edit_manip.setText(name_val)

        # Date
        date_val = _value_for("Date :")
        if date_val and hasattr(self, "edit_date"):
            qd = QDate.fromString(date_val, "dd/MM/yyyy")
            if qd.isValid():
                self.edit_date.setDate(qd)

        # Lieu
        loc_val = _value_for("Lieu :")
        if loc_val and hasattr(self, "edit_location"):
            self.edit_location.setText(loc_val)

        # Coordinateur
        coord_val = _value_for("Coordinateur :")
        if coord_val and hasattr(self, "edit_coordinator"):
            self.edit_coordinator.setText(coord_val)

        # Opérateur
        oper_val = _value_for("Opérateur :")
        if oper_val and hasattr(self, "edit_operator"):
            self.edit_operator.setText(oper_val)

    # ------------------------------------------------------------------
    # Sauvegarde / chargement des métadonnées complètes
    # ------------------------------------------------------------------
    def _on_save_metadata_clicked(self) -> None:
        """Enregistre df_comp, df_conc (si dispo) et df_map dans un fichier Excel."""
        if self.df_comp is None or self.df_map is None:
            QMessageBox.warning(
                self,
                "Métadonnées incomplètes",
                "Impossible d'enregistrer : le tableau des volumes ou la correspondance spectres ↔ tubes n'est pas encore défini.",
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Enregistrer les métadonnées",
            "",
            "Fichier Excel (*.xlsx)"
        )
        if not path:
            return

        try:
            with pd.ExcelWriter(path) as writer:
                self.df_comp.to_excel(writer, sheet_name="Volumes", index=False)
                if self.df_conc is not None and isinstance(self.df_conc, pd.DataFrame) and not self.df_conc.empty:
                    self.df_conc.to_excel(writer, sheet_name="Concentrations", index=False)
                self.df_map.to_excel(writer, sheet_name="Correspondance", index=False)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erreur d'enregistrement",
                f"Impossible d'enregistrer les métadonnées dans le fichier :\n{e}",
            )
            return

        QMessageBox.information(self, "Enregistrement réussi", f"Métadonnées enregistrées dans :\n{path}")

    def _on_load_metadata_clicked(self) -> None:
        """Charge df_comp, df_conc et df_map à partir d'un fichier Excel précédemment sauvegardé."""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Charger des métadonnées",
            "",
            "Fichier Excel (*.xlsx)"
        )
        if not path:
            return

        try:
            xls = pd.ExcelFile(path)
            # Volumes
            if "Volumes" in xls.sheet_names:
                self.df_comp = pd.read_excel(xls, sheet_name="Volumes")
                self.lbl_status_comp.setText(
                    f"Tableau des volumes : {self.df_comp.shape[0]} ligne(s), {self.df_comp.shape[1]} colonne(s)"
                )
            # Concentrations (optionnel)
            if "Concentrations" in xls.sheet_names:
                self.df_conc = pd.read_excel(xls, sheet_name="Concentrations")
            else:
                # Recalculer à partir de df_comp si possible
                try:
                    self.df_conc = self._compute_concentration_table()
                except Exception:
                    self.df_conc = None

            # Correspondance
            if "Correspondance" in xls.sheet_names:
                self.df_map = pd.read_excel(xls, sheet_name="Correspondance")
                self.lbl_status_map.setText(
                    f"Tableau de correspondance : {self.df_map.shape[0]} ligne(s), {self.df_map.shape[1]} colonne(s)"
                )
                # Mettre à jour les champs d'en-tête d'après df_map
                self._update_header_fields_from_df_map()
                self.map_dirty = False
            else:
                QMessageBox.warning(
                    self,
                    "Feuille manquante",
                    "La feuille 'Correspondance' est absente du fichier Excel chargé.\n"
                    "Le tableau de correspondance n'a pas pu être restauré.",
                )
            self._refresh_button_states()
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erreur de chargement",
                f"Impossible de charger les métadonnées depuis le fichier :\n{e}",
            )