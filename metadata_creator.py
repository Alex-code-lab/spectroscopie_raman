

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

        layout = QVBoxLayout(self)

        # Champs d'en-tête pour la manipulation : nom, date, lieu, coordinateur, opérateur
        header_layout = QVBoxLayout()

        # Ligne 1 : nom de la manip + date (avec calendrier)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Nom de la manip :", self))
        self.edit_manip = QLineEdit(self)
        self.edit_manip.setPlaceholderText("ex : GROUPE_20251203")
        row1.addWidget(self.edit_manip)

        row1.addWidget(QLabel("Date :", self))
        self.edit_date = QDateEdit(self)
        self.edit_date.setCalendarPopup(True)
        self.edit_date.setDisplayFormat("dd/MM/yyyy")
        self.edit_date.setDate(QDate.currentDate())
        row1.addWidget(self.edit_date)

        header_layout.addLayout(row1)

        # Ligne 2 : lieu, coordinateur, opérateur
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Lieu :", self))
        self.edit_location = QLineEdit(self)
        self.edit_location.setPlaceholderText("ex : Laboratoire / Terrain")
        row2.addWidget(self.edit_location)

        row2.addWidget(QLabel("Coordinateur :", self))
        self.edit_coordinator = QLineEdit(self)
        self.edit_coordinator.setPlaceholderText("ex : Nom du coordinateur")
        row2.addWidget(self.edit_coordinator)

        row2.addWidget(QLabel("Opérateur :", self))
        self.edit_operator = QLineEdit(self)
        self.edit_operator.setPlaceholderText("ex : Nom de l'opérateur")
        row2.addWidget(self.edit_operator)

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
        self.btn_show_conc = QPushButton("Calculer / afficher le tableau des concentrations", self)

        self.btn_edit_comp.clicked.connect(self._on_edit_comp_clicked)
        self.btn_edit_map.clicked.connect(self._on_edit_map_clicked)
        self.btn_show_conc.clicked.connect(self._on_show_conc_clicked)

        btns.addWidget(self.btn_edit_comp)
        btns.addWidget(self.btn_edit_map)
        btns.addWidget(self.btn_show_conc)
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
    # Gestion du tableau de correspondance spectres ↔ tubes
    # ------------------------------------------------------------------
    def _on_edit_map_clicked(self) -> None:
        """
        Ouvre le dialog d'édition pour la correspondance spectres ↔ tubes.

        Le modèle par défaut suit la structure Excel d'origine avec :
        - 5 lignes d'en-tête (Nom de la manip, Date, Lieu, Coordinateur, Opérateur),
        - une ligne vide,
        - une ligne d'en-tête interne "Nom du spectre" / "Tube",
        - puis les lignes de correspondance NomSpectre ↔ Tube.

        Les noms de spectres sont automatiquement générés à partir du nom de la manip
        sous la forme NomManip_00, NomManip_01, … (indice à 2 chiffres).
        """
        default_cols = ["Nom du spectre", "Tube"]

        # Récupération du nom de manipulation (champ utilisateur)
        try:
            manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        except Exception:
            manip_name = ""

        if not manip_name:
            manip_name = "MANIP"

        meta_keys = [
            "Nom de la manip :",
            "Date :",
            "Lieu :",
            "Coordinateur :",
            "Opérateur :",
        ]

        # ---- Cas 1 : aucun tableau existant -> créer un gabarit par défaut ----
        if self.df_map is None:
            # 12 spectres par défaut
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
                        "01/01/2024",
                        "Campus de la transition",
                        "Dupont",
                        "Durand",
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

        # ---- Cas 2 : un tableau existe déjà -> on le réutilise tel quel ----
        # Dans tous les cas à partir d'ici, self.df_map existe.
        # On met à jour les lignes d'en-tête à partir des champs UI.
        try:
            col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()

            # Nom de la manip
            mask_manip = col_ns == "Nom de la manip :"
            if mask_manip.any():
                self.df_map.loc[mask_manip, "Tube"] = manip_name

            # Date
            date_str = ""
            if hasattr(self, "edit_date") and self.edit_date is not None:
                try:
                    date_str = self.edit_date.date().toString("dd/MM/yyyy")
                except Exception:
                    date_str = ""
            mask_date = col_ns == "Date :"
            if mask_date.any():
                self.df_map.loc[mask_date, "Tube"] = date_str

            # Lieu
            location = ""
            if hasattr(self, "edit_location") and self.edit_location is not None:
                try:
                    location = (self.edit_location.text() or "").strip()
                except Exception:
                    location = ""
            mask_lieu = col_ns == "Lieu :"
            if mask_lieu.any():
                self.df_map.loc[mask_lieu, "Tube"] = location

            # Coordinateur
            coordinator = ""
            if hasattr(self, "edit_coordinator") and self.edit_coordinator is not None:
                try:
                    coordinator = (self.edit_coordinator.text() or "").strip()
                except Exception:
                    coordinator = ""
            mask_coord = col_ns == "Coordinateur :"
            if mask_coord.any():
                self.df_map.loc[mask_coord, "Tube"] = coordinator

            # Opérateur
            operator = ""
            if hasattr(self, "edit_operator") and self.edit_operator is not None:
                try:
                    operator = (self.edit_operator.text() or "").strip()
                except Exception:
                    operator = ""
            mask_oper = col_ns == "Opérateur :"
            if mask_oper.any():
                self.df_map.loc[mask_oper, "Tube"] = operator

        except Exception:
            pass

        # Regénérer automatiquement les noms de spectres (NomManip_XX)
        # try:
        #     col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
        #     mask_mapping = ~col_ns.isin(meta_keys + ["", "Nom du spectre"])
        #     idx_mapping = list(self.df_map.index[mask_mapping])
        #     for k, idx in enumerate(idx_mapping):
        #         self.df_map.loc[idx, "Nom du spectre"] = f"{manip_name}_{k:02d}"
        # except Exception:
        #     pass

        # Ouvrir le dialogue d'édition avec le DataFrame actuel
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

            # Synchroniser le champ "Nom de la manip" depuis le tableau, si présent
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
                mask = col_ns == "Nom de la manip :"
                if mask.any() and hasattr(self, "edit_manip"):
                    val = str(self.df_map.loc[mask, "Tube"].iloc[0])
                    self.edit_manip.setText(val)
            except Exception:
                pass

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
    # Gestion du tableau de correspondance spectres/tubes
    # ------------------------------------------------------------------
    def _on_edit_map_clicked(self) -> None:
        """
        Ouvre le dialog d'édition pour la correspondance spectres ↔ tubes.

        Le modèle par défaut suit la structure Excel : deux colonnes
        "Nom du spectre" et "Tube", avec les premières lignes utilisées pour
        les métadonnées (Nom de la manip, Date, Lieu, Coordinateur, Opérateur),
        puis une ligne d'en-tête "Nom du spectre" / "Tube" et enfin la liste des spectres.

        Les noms de spectres sont automatiquement générés à partir du nom de la manipulation
        sous la forme NomManip_00, NomManip_01, … (indice à 2 chiffres).
        """
        default_cols = ["Nom du spectre", "Tube"]

        # ───────────────────────────────────────────────────────────────
        # Récupération du nom de manipulation (champ utilisateur)
        # ───────────────────────────────────────────────────────────────
        try:
            manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        except Exception:
            manip_name = ""

        if not manip_name:
            manip_name = "MANIP"

        # Entrées méta selon le format Excel d'origine
        meta_keys = [
            "Nom de la manip :",
            "Date :",
            "Lieu :",
            "Coordinateur :",
            "Opérateur :",
        ]

        # ───────────────────────────────────────────────────────────────
        # Cas 1 : Aucun tableau existant → créer une version par défaut
        # ───────────────────────────────────────────────────────────────
        if self.df_map is None:
            # Créer 12 spectres par défaut
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
                        "01/01/2024",
                        "Campus de la transition",
                        "Dupont",
                        "Durand",
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

        # ───────────────────────────────────────────────────────────────
        # Cas 2 : Un tableau existe déjà → mettre à jour méta + spectres
        # ───────────────────────────────────────────────────────────────
        else:
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()

                # --- Mise à jour : Nom de la manip ---
                mask = col_ns == "Nom de la manip :"
                if mask.any():
                    self.df_map.loc[mask, "Tube"] = manip_name

                # --- Mise à jour : Date ---
                date_str = ""
                if hasattr(self, "edit_date") and self.edit_date is not None:
                    try:
                        date_str = self.edit_date.date().toString("dd/MM/yyyy")
                    except Exception:
                        date_str = ""
                mask = col_ns == "Date :"
                if mask.any():
                    self.df_map.loc[mask, "Tube"] = date_str

                # --- Mise à jour : Lieu ---
                location = ""
                if hasattr(self, "edit_location") and self.edit_location is not None:
                    try:
                        location = (self.edit_location.text() or "").strip()
                    except Exception:
                        location = ""
                mask = col_ns == "Lieu :"
                if mask.any():
                    self.df_map.loc[mask, "Tube"] = location

                # --- Mise à jour : Coordinateur ---
                coordinator = ""
                if hasattr(self, "edit_coordinator") and self.edit_coordinator is not None:
                    try:
                        coordinator = (self.edit_coordinator.text() or "").strip()
                    except Exception:
                        coordinator = ""
                mask = col_ns == "Coordinateur :"
                if mask.any():
                    self.df_map.loc[mask, "Tube"] = coordinator

                # --- Mise à jour : Opérateur ---
                operator = ""
                if hasattr(self, "edit_operator") and self.edit_operator is not None:
                    try:
                        operator = (self.edit_operator.text() or "").strip()
                    except Exception:
                        operator = ""
                mask = col_ns == "Opérateur :"
                if mask.any():
                    self.df_map.loc[mask, "Tube"] = operator

            except Exception:
                pass

            # --- Regénérer automatiquement les noms de spectres (NomManip_XX) ---
            try:
                col_ns = self.df_map["Nom du spectre"].astype(str).str.strip()
                mask_mapping = ~col_ns.isin(meta_keys + ["", "Nom du spectre"])
                idx_mapping = list(self.df_map.index[mask_mapping])
                for k, idx in enumerate(idx_mapping):
                    self.df_map.loc[idx, "Nom du spectre"] = f"{manip_name}_{k:02d}"
            except Exception:
                pass

        # ───────────────────────────────────────────────────────────────
        # Ouverture du tableau d'édition (dialog)
        # ───────────────────────────────────────────────────────────────
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

            # Synchroniser le champ "Nom de la manip"
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
            "Tube BRB": "Cuvette BRB",
            "Tube MQW": "MQW",
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

                    # Tenter de détecter une colonne EGTA (nom de réactif contenant "egta")
                    egta_candidates: list[str] = []
                    for col in conc_wide.columns:
                        name_low = str(col).strip().lower()
                        # 1) Cas général : le nom du réactif contient explicitement "egta"
                        if "egta" in name_low:
                            egta_candidates.append(col)
                        # 2) Cas spécifique à ton protocole : la solution B est l'EGTA
                        elif name_low in {"solution b", "solutionb", "sol b", "solb", "b"}:
                            egta_candidates.append(col)

                    # Si rien n'a été trouvé, egta_candidates restera vide et on n'ajoutera
                    # pas les colonnes [EGTA].
                    if egta_candidates:
                        egta_col = egta_candidates[0]
                        # S'assurer que la colonne EGTA est bien numérique (sinon NaN)
                        meta[egta_col] = pd.to_numeric(meta[egta_col], errors="coerce")
                        # Colonnes explicites pour l'analyse : [EGTA] (M) et [EGTA] (mM)
                        meta["[EGTA] (M)"] = meta[egta_col]
                        meta["[EGTA] (mM)"] = meta[egta_col] * 1e3

        # Exclure explicitement les tubes de type "Tube BRB" et "Tube MQW"
        # des métadonnées utilisées pour l'assemblage (tracés, analyse)
        meta = meta[~meta["Tube"].isin(["Tube BRB", "Tube MQW"])].copy()

        # Réindexer proprement
        meta = meta.reset_index(drop=True)

        #TEST
        print("\n=== Spectrum name depuis les métadonnées ===")
        print(meta["Spectrum name"].unique()[:20])

        return meta