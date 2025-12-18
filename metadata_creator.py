

# metadata_creator.py

from __future__ import annotations

import pandas as pd
import numpy as np
import re
from scipy.stats import norm


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
    QComboBox,
    QSpinBox,
    QDoubleSpinBox,
    QFormLayout,
    QCheckBox,
    QGroupBox,
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
        reset_df: pd.DataFrame | None = None,
        reset_button_text: str = "Reset",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)

        self._columns = list(columns)
        # self.conc_user_overridden = False
        self._reset_df = reset_df.copy() if reset_df is not None else None
        self._reset_button_text = reset_button_text

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

        btns_edit.addSpacing(20)
        self.btn_reset = QPushButton(self._reset_button_text, self)
        self.btn_reset.setEnabled(self._reset_df is not None)
        btns_edit.addWidget(self.btn_reset)

        btns_edit.addStretch(1)
        layout.addLayout(btns_edit)

        self.btn_add_row.clicked.connect(self._add_row)
        self.btn_del_row.clicked.connect(self._del_row)
        self.btn_add_col.clicked.connect(self._add_col)
        self.btn_del_col.clicked.connect(self._del_col)
        self.btn_reset.clicked.connect(self._reset_table)

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
                if pd.isna(val):
                    txt = ""
                else:
                    # Affichage lisible : floats en notation scientifique à 2 décimales
                    if isinstance(val, (float, int)):
                        try:
                            txt = f"{float(val):.2e}"
                        except Exception:
                            txt = str(val)
                    else:
                        txt = str(val)
                item = QTableWidgetItem(txt)
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

    def _reset_table(self) -> None:
        """Restaure le tableau par défaut (si fourni)."""
        if self._reset_df is None:
            return
        self._load_from_dataframe(self._reset_df)

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


def sers_gaussian_volumes(
    *,
    n_tubes: int,
    C0_nM: float,
    V_echant_uL: float,
    C_titrant_uM: float,
    Vtot_uL: float,
    V_Solution_C_uL: float = 30.0,
    V_Solution_D_uL: float = 750.0,
    V_Solution_E_uL: float = 750.0,
    V_Solution_F_uL: float = 60.0,
    margin_uL: float = 20.0,
    pipette_step_uL: float = 1.0,
    include_zero: bool = True,
    duplicate_last: bool = True,
) -> tuple[pd.DataFrame, dict]:
    if norm is None:
        raise ImportError("scipy n'est pas disponible (scipy.stats.norm requis).")

    if n_tubes < 2:
        raise ValueError("n_tubes doit être >= 2")

    V_fixe_uL = (
        float(V_echant_uL)
        + float(V_Solution_C_uL)
        + float(V_Solution_D_uL)
        + float(V_Solution_E_uL)
        + float(V_Solution_F_uL)
    )

    V_AB = float(Vtot_uL) - V_fixe_uL
    if V_AB <= 0:
        raise ValueError("Impossible: Vtot <= V_fixe (V_AB <= 0).")

    Vmin_uL = float(margin_uL)
    Vmax_uL = float(V_AB - margin_uL)
    if Vmax_uL < Vmin_uL:
        raise ValueError("Impossible: margin_uL trop grand par rapport à V_AB.")

    # équivalence
    C0_M = float(C0_nM) * 1e-9
    V_echant_L = float(V_echant_uL) * 1e-6
    C_titrant_M = float(C_titrant_uM) * 1e-6
    if C_titrant_M <= 0:
        raise ValueError("C_titrant_uM doit être > 0.")

    n_Cu = C0_M * V_echant_L
    Veq_uL = (n_Cu / C_titrant_M) * 1e6

    n_free = int(n_tubes) - int(include_zero) - int(duplicate_last)
    if n_free <= 0:
        raise ValueError("Pas assez de tubes libres (n_free <= 0).")

    p = (np.arange(1, n_free + 1) - 0.5) / n_free
    z = norm.ppf(p)
    zmax = np.max(np.abs(z)) if n_free > 1 else 1.0

    sigma_V = max(
        (Vmax_uL - Veq_uL) / zmax,
        (Veq_uL - Vmin_uL) / zmax,
    )
    sigma_V = max(0.0, float(sigma_V))

    V_B_free = Veq_uL + sigma_V * z
    V_B_free = np.clip(V_B_free, Vmin_uL, Vmax_uL)

    step = float(pipette_step_uL)
    if step <= 0:
        raise ValueError("pipette_step_uL doit être > 0.")

    def _quantize_unique_increasing(vals: np.ndarray) -> np.ndarray:
        """Quantifie `vals` sur une grille (Vmin..Vmax, pas=step) en garantissant
        une suite strictement croissante (pas de doublons).
        """
        vals = np.asarray(vals, dtype=float)
        if vals.size == 0:
            return vals
        if vals.size == 1:
            out = np.round(vals / step) * step
            return np.clip(out, Vmin_uL, Vmax_uL)

        span = float(Vmax_uL) - float(Vmin_uL)
        max_idx = int(np.floor(span / step + 1e-12))
        if max_idx < 0:
            max_idx = 0
        n_grid = max_idx + 1
        if vals.size > n_grid:
            raise ValueError(
                "Impossible de générer des volumes tous distincts avec ce pas de pipette "
                "(pipette_step_uL trop grand et/ou marge trop grande)."
            )

        idx = np.rint((vals - float(Vmin_uL)) / step).astype(int)
        idx = np.clip(idx, 0, max_idx)

        # Forcer une suite strictement croissante
        idx_adj = idx.copy()
        for i in range(1, idx_adj.size):
            if idx_adj[i] <= idx_adj[i - 1]:
                idx_adj[i] = idx_adj[i - 1] + 1

        # Si on dépasse la grille, on "recolle" par le haut
        if idx_adj[-1] > max_idx:
            idx_adj[-1] = max_idx
            for i in range(idx_adj.size - 2, -1, -1):
                if idx_adj[i] >= idx_adj[i + 1]:
                    idx_adj[i] = idx_adj[i + 1] - 1
            if idx_adj[0] < 0:
                raise ValueError(
                    "Impossible de générer des volumes tous distincts avec ces paramètres "
                    "(pas trop grand / trop de tubes)."
                )

        out = float(Vmin_uL) + idx_adj.astype(float) * step
        return np.clip(out, Vmin_uL, Vmax_uL)

    # Si on ne veut PAS de duplicat, on force des volumes distincts (après quantification).
    # Même si duplicate_last=True, on garde les valeurs libres distinctes et on duplique uniquement le dernier tube.
    V_B_free = _quantize_unique_increasing(V_B_free)

    V_B = []
    if include_zero:
        V_B.append(0.0)
    V_B.extend(V_B_free.tolist())
    if duplicate_last:
        V_B.append(V_B[-1])
    V_B = np.sort(np.array(V_B, dtype=float))

    V_A = V_AB - V_B
    if np.any(V_A < -1e-9):
        raise ValueError("Impossible: certains tubes ont V_A négatif.")
    V_A = np.round(V_A / step) * step
    V_A += (V_AB - (V_A + V_B))

    C_titrant_nM = float(C_titrant_uM) * 1e3
    C_final_nM = C_titrant_nM * (V_B / float(Vtot_uL))

    df = pd.DataFrame(
        {
            "Tube": np.arange(1, n_tubes + 1),
            "V_B_titrant_uL": V_B,
            "V_A_tampon_uL": V_A,
            "C_final_from_B_nM": np.round(C_final_nM, 2),
        }
    )

    meta = {
        "V_fixe_uL": float(V_fixe_uL),
        "V_AB_uL": float(V_AB),
        "Veq_uL": float(Veq_uL),
        "sigma_V_uL": float(sigma_V),
        "Vmin_nonzero_uL": float(Vmin_uL),
        "Vmax_nonzero_uL": float(Vmax_uL),
        "n_free": int(n_free),
    }
    return df, meta

class GaussianVolumesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Générer volumes (gaussien)")
        self.resize(520, 420)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Génère automatiquement les volumes de titrant (Solution B) autour de l'équivalence.\n"
            "Seules Solution A (tampon) et Solution B (titrant) varient."
        ))

        box = QGroupBox("Paramètres")
        form = QFormLayout(box)

        self.spin_n = QSpinBox(self); self.spin_n.setRange(2, 200); self.spin_n.setValue(11)
        self.spin_C0 = QDoubleSpinBox(self); self.spin_C0.setDecimals(3); self.spin_C0.setRange(0.0, 1e9); self.spin_C0.setValue(492.0)
        self.spin_Vech = QDoubleSpinBox(self); self.spin_Vech.setDecimals(1); self.spin_Vech.setRange(0.0, 1e9); self.spin_Vech.setValue(1000.0)
        self.spin_Ctit = QDoubleSpinBox(self); self.spin_Ctit.setDecimals(3); self.spin_Ctit.setRange(0.0, 1e9); self.spin_Ctit.setValue(2.0)
        self.spin_Vtot = QDoubleSpinBox(self); self.spin_Vtot.setDecimals(1); self.spin_Vtot.setRange(0.0, 1e9); self.spin_Vtot.setValue(3090.0)

        form.addRow("Nombre de tubes", self.spin_n)
        form.addRow("C0 Cu (nM)", self.spin_C0)
        form.addRow("V échantillon (µL)", self.spin_Vech)
        form.addRow("C titrant (µM)", self.spin_Ctit)
        form.addRow("V total par tube (µL)", self.spin_Vtot)

        layout.addWidget(box)

        box2 = QGroupBox("Solutions fixes")
        form2 = QFormLayout(box2)

        unit_choices = ["M", "mM", "µM", "nM", "% en masse"]

        self.spin_VC = QDoubleSpinBox(self); self.spin_VC.setDecimals(1); self.spin_VC.setRange(0.0, 1e9); self.spin_VC.setValue(30.0)
        self.spin_VD = QDoubleSpinBox(self); self.spin_VD.setDecimals(1); self.spin_VD.setRange(0.0, 1e9); self.spin_VD.setValue(750.0)
        self.spin_VE = QDoubleSpinBox(self); self.spin_VE.setDecimals(1); self.spin_VE.setRange(0.0, 1e9); self.spin_VE.setValue(750.0)
        self.spin_VF = QDoubleSpinBox(self); self.spin_VF.setDecimals(1); self.spin_VF.setRange(0.0, 1e9); self.spin_VF.setValue(60.0)

        self.spin_conc_C = QDoubleSpinBox(self); self.spin_conc_C.setDecimals(6); self.spin_conc_C.setRange(0.0, 1e12); self.spin_conc_C.setValue(2.0)
        self.spin_conc_D = QDoubleSpinBox(self); self.spin_conc_D.setDecimals(6); self.spin_conc_D.setRange(0.0, 1e12); self.spin_conc_D.setValue(0.5)
        self.spin_conc_E = QDoubleSpinBox(self); self.spin_conc_E.setDecimals(6); self.spin_conc_E.setRange(0.0, 1e12); self.spin_conc_E.setValue(1.0)
        self.spin_conc_F = QDoubleSpinBox(self); self.spin_conc_F.setDecimals(6); self.spin_conc_F.setRange(0.0, 1e12); self.spin_conc_F.setValue(1.0)

        self.combo_unit_C = QComboBox(self); self.combo_unit_C.addItems(unit_choices); self.combo_unit_C.setCurrentText("µM")
        self.combo_unit_D = QComboBox(self); self.combo_unit_D.addItems(unit_choices); self.combo_unit_D.setCurrentText("% en masse")
        self.combo_unit_E = QComboBox(self); self.combo_unit_E.addItems(unit_choices); self.combo_unit_E.setCurrentText("mM")
        self.combo_unit_F = QComboBox(self); self.combo_unit_F.addItems(unit_choices); self.combo_unit_F.setCurrentText("mM")

        def _row(vol_spin: QDoubleSpinBox, conc_spin: QDoubleSpinBox, unit_combo: QComboBox) -> QWidget:
            w = QWidget(self)
            h = QHBoxLayout(w)
            h.setContentsMargins(0, 0, 0, 0)
            h.addWidget(QLabel("V (µL)", w))
            h.addWidget(vol_spin)
            h.addSpacing(12)
            h.addWidget(QLabel("C", w))
            h.addWidget(conc_spin)
            h.addWidget(unit_combo)
            h.addStretch(1)
            return w

        form2.addRow("Solution C", _row(self.spin_VC, self.spin_conc_C, self.combo_unit_C))
        form2.addRow("Solution D", _row(self.spin_VD, self.spin_conc_D, self.combo_unit_D))
        form2.addRow("Solution E", _row(self.spin_VE, self.spin_conc_E, self.combo_unit_E))
        form2.addRow("Solution F", _row(self.spin_VF, self.spin_conc_F, self.combo_unit_F))

        layout.addWidget(box2)

        box3 = QGroupBox("Contraintes")
        form3 = QFormLayout(box3)

        self.spin_margin = QDoubleSpinBox(self); self.spin_margin.setDecimals(1); self.spin_margin.setRange(0.0, 1e9); self.spin_margin.setValue(20.0)
        self.spin_step = QDoubleSpinBox(self); self.spin_step.setDecimals(3); self.spin_step.setRange(0.001, 1e6); self.spin_step.setValue(1.0)

        self.chk_zero = QCheckBox("Inclure un tube à 0 µL de titrant", self); self.chk_zero.setChecked(True)
        self.chk_dup = QCheckBox("Dupliquer le dernier volume pour contrôle", self); self.chk_dup.setChecked(True)

        form3.addRow("Marge (µL)", self.spin_margin)
        form3.addRow("Pas pipette (µL)", self.spin_step)
        form3.addRow(self.chk_zero)
        form3.addRow(self.chk_dup)

        layout.addWidget(box3)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def params(self) -> dict:
        return {
            "n_tubes": int(self.spin_n.value()),
            "C0_nM": float(self.spin_C0.value()),
            "V_echant_uL": float(self.spin_Vech.value()),
            "C_titrant_uM": float(self.spin_Ctit.value()),
            "Vtot_uL": float(self.spin_Vtot.value()),
            "V_Solution_C_uL": float(self.spin_VC.value()),
            "V_Solution_D_uL": float(self.spin_VD.value()),
            "V_Solution_E_uL": float(self.spin_VE.value()),
            "V_Solution_F_uL": float(self.spin_VF.value()),
            "margin_uL": float(self.spin_margin.value()),
            "pipette_step_uL": float(self.spin_step.value()),
            "include_zero": bool(self.chk_zero.isChecked()),
            "duplicate_last": bool(self.chk_dup.isChecked()),
        }

    def fixed_solution_stocks(self) -> dict[str, tuple[float, str]]:
        return {
            "Solution C": (float(self.spin_conc_C.value()), str(self.combo_unit_C.currentText()).strip()),
            "Solution D": (float(self.spin_conc_D.value()), str(self.combo_unit_D.currentText()).strip()),
            "Solution E": (float(self.spin_conc_E.value()), str(self.combo_unit_E.currentText()).strip()),
            "Solution F": (float(self.spin_conc_F.value()), str(self.combo_unit_F.currentText()).strip()),
        }
    
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

    # Modèles prédéfinis de tableaux des volumes
    VOLUME_PRESETS: dict[str, pd.DataFrame] = {
       "Titration classique (historique)": pd.DataFrame(
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
                "Concentration": ["0,5", "", "4", "2", "2", "0,5", "1", "1"],
                "Unité": ["µM", "", "mM", "µM", "µM", "% en masse", "mM", "mM"],
                "Tube 1":  [1000, 0, 500,   0,  30, 750, 750,  60],
                "Tube 2":  [1000, 0, 411,  89,  30, 750, 750,  60],
                "Tube 3":  [1000, 0, 321, 179,  30, 750, 750,  60],
                "Tube 4":  [1000, 0, 299, 201,  30, 750, 750,  60],
                "Tube 5":  [1000, 0, 277, 223,  30, 750, 750,  60],
                "Tube 6":  [1000, 0, 254, 246,  30, 750, 750,  60],
                "Tube 7":  [1000, 0, 232, 268,  30, 750, 750,  60],
                "Tube 8":  [1000, 0, 188, 312,  30, 750, 750,  60],
                "Tube 9":  [1000, 0,  98, 402,  30, 750, 750,  60],
                "Tube 10": [1000, 0,  54, 446,  30, 750, 750,  60],
                "Tube 11": [   0, 1000,  54, 446,  30, 750, 750,  60],
            }
        ),
        "Titration Angelina": pd.DataFrame(
            {
                "Réactif": [
                    "Echantillon",
                    "Contrôle",
                    "Solution B",
                    "Solution C",
                    "Solution D",
                    "Solution E",
                ],
                "Concentration": ["0,5", "", "4", "2", "0,5", "1",],
                "Unité": ["µM", "", "mM",  "µM", "% en masse","mM"],
                "Tube 1":  [1000, 0, 500,   0, 750, 750],
                "Tube 2":  [1000, 0, 411,  89, 750, 750],
                "Tube 3":  [1000, 0, 321, 179, 750, 750],
                "Tube 4":  [1000, 0, 299, 201, 750, 750],
                "Tube 5":  [1000, 0, 277, 223, 750, 750],
                "Tube 6":  [1000, 0, 254, 246, 750, 750],
                "Tube 7":  [1000, 0, 232, 268, 750, 750],
                "Tube 8":  [1000, 0, 188, 312, 750, 750],
                "Tube 9":  [1000, 0,  98, 402, 750, 750],
                "Tube 10": [1000, 0,  54, 446, 750, 750],
                "Tube 11": [1000, 0,  54, 446, 750, 750],
            }
        ),
    }

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.df_comp: pd.DataFrame | None = None # tableau de composition des tubes
        self.df_map: pd.DataFrame | None = None # tableau de correspondance spectres ↔ tubes
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
        row1.addWidget(QLabel("Départ index spectres :", self))
        self.spin_spec_start = QSpinBox(self)
        self.spin_spec_start.setRange(0, 999)
        self.spin_spec_start.setValue(0)  # par défaut : 00
        self.spin_spec_start.setToolTip("Index de départ (0 => _00, 7 => _07).")
        row1.addWidget(self.spin_spec_start)

        # Changer l'index déclenche la mise à jour des noms
        self.spin_spec_start.valueChanged.connect(self._on_manip_name_changed)

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
            "</ul>"
            "<b>Glossaire des solutions utilisées :</b><ul>"
            "<li><b>Solution A</b> : Tampon (milieu de base, contrôle du pH)</li>"
            "<li><b>Solution B</b> : Titrant (variable expérimentale principale)</li>"
            "<li><b>Solution C</b> : Indicateur (molécule rapporteur suivie par spectroscopie)</li>"
            "<li><b>Solution D</b> : PEG (agent de crowding / modification du milieu)</li>"
            "<li><b>Solution E</b> : NPs (nanoparticules)</li>"
            "<li><b>Solution F</b> : Crosslinker (ex. spermine, agent de réticulation)</li>"
            "</ul>"
            "Les tableaux sont éditables comme dans Excel mais sont stockés en interne sous forme de DataFrame pandas.",
            self,
        ))

        # --- Sélecteur de modèle de tableau des volumes ---
        preset_layout = QHBoxLayout()
        preset_layout.addWidget(QLabel("Modèle de tableau des volumes :", self))

        self.combo_volume_preset = QComboBox(self)
        self.combo_volume_preset.addItems(self.VOLUME_PRESETS.keys())
        preset_layout.addWidget(self.combo_volume_preset)

        # Le changement de modèle ne s'applique pas automatiquement : il faut valider.
        self.btn_validate_volume_preset = QPushButton("Valider le modèle", self)
        self.btn_validate_volume_preset.clicked.connect(self._on_validate_volume_preset_clicked)
        preset_layout.addWidget(self.btn_validate_volume_preset)

        preset_layout.addStretch(1)
        layout.addLayout(preset_layout)
    
        self.btn_generate_volumes = QPushButton("Générer volumes (gaussien)…", self)
        self.btn_generate_volumes.clicked.connect(self._on_generate_volumes_clicked)
        preset_layout.addWidget(self.btn_generate_volumes)

        # --- Boutons pour ouvrir les éditeurs de tableaux ---
        btns = QHBoxLayout()
        self.btn_edit_comp = QPushButton("Créer / éditer le tableau des volumes", self)
        self.btn_edit_map = QPushButton("Créer / éditer la correspondance spectres ↔ tubes", self)
        self.btn_show_conc = QPushButton("Concentration dans les cuvettes (info)", self)

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
        # sinon rouge.
        if comp_ready:
            self.btn_show_conc.setStyleSheet(green)
        else:
            self.btn_show_conc.setStyleSheet(red)

        # Mettre à jour les tooltips
        self.btn_edit_comp.setToolTip("Vert = tableau des volumes défini" if comp_ready else "Rouge = à créer / éditer")
        self.btn_edit_map.setToolTip("Vert = correspondance définie" if map_ready else "Rouge = à créer / éditer")
        if comp_ready:
            self.btn_show_conc.setToolTip("Calculable (volumes définis)")
        else:
            self.btn_show_conc.setToolTip("Rouge = définir d'abord les volumes")
# ------------------------------------------------------------------
# Helpers – recalcul à partir d'un tableau des concentrations édité
# ------------------------------------------------------------------
    @staticmethod
    def _to_float(x):
        """Convertit une valeur (str avec virgule, etc.) en float ou retourne None."""
        if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and x.strip() == ""):
            return None
        try:
            if isinstance(x, str):
                x = x.replace("\xa0", " ").strip().replace(",", ".")
            return float(x)
        except Exception:
            return None

    @staticmethod
    def _normalize_tube_label(label: str) -> str:
        s = ("" if label is None else str(label)).strip()
        if not s:
            return ""

        # A1, A01, B3 ... => Tube N (on prend le nombre)
        m = re.match(r"^[A-Za-z](\d{1,2})$", s)
        if m:
            return f"Tube {int(m.group(1))}"

        # Tube1 / TUBE 01 / tube   2 => Tube N
        m = re.match(r"^(?:tube)\s*(\d{1,2})$", s, flags=re.IGNORECASE)
        if m:
            return f"Tube {int(m.group(1))}"

        # juste un nombre
        m = re.match(r"^(\d{1,2})$", s)
        if m:
            return f"Tube {int(m.group(1))}"

        return s
    
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

        # Nombre de tubes : on s'aligne sur le tableau des volumes si disponible
        n_tubes = self._infer_n_tubes_for_mapping()
        tube_labels_default = self._default_tube_labels_for_mapping(n_tubes)

        start_index = int(self.spin_spec_start.value()) if hasattr(self, "spin_spec_start") else 0

        def _make_mapping_rows(labels: list[str]) -> pd.DataFrame:
            names = [f"{manip_name}_{i:02d}" for i in range(start_index, start_index + len(labels))]
            return pd.DataFrame({"Nom du spectre": names, "Tube": labels}, columns=default_cols)

        # --- DataFrame RESET (valeurs par défaut) ---
        mapping_rows_default = _make_mapping_rows(tube_labels_default)
        reset_df = pd.concat(
            [df_header, blank_row, internal_header, mapping_rows_default],
            ignore_index=True,
        )

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
            mapping_rows = _make_mapping_rows(tube_labels_default)
        else:
            # Si le nombre de tubes a changé (ex : 15), on complète automatiquement
            # pour que l'utilisateur voie toutes les lignes nécessaires.
            def _clean_text(v) -> str:
                if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
                    return ""
                s = str(v).strip()
                return "" if s == "<NA>" else s

            existing_by_tube: dict[str, dict[str, str]] = {}
            extra_rows: list[dict] = []
            desired_keys = {self._tube_merge_key_for_mapping(t, n_tubes) for t in tube_labels_default}

            for _, row in mapping_rows.iterrows():
                tube_raw = _clean_text(row.get("Tube", ""))
                name_raw = _clean_text(row.get("Nom du spectre", ""))
                tube_key = self._tube_merge_key_for_mapping(tube_raw, n_tubes)
                if not tube_key:
                    continue
                if tube_key in desired_keys:
                    # Conserver la première ligne rencontrée pour ce tube (sans écraser les valeurs)
                    if tube_key not in existing_by_tube:
                        existing_by_tube[tube_key] = {"Nom du spectre": name_raw, "Tube": tube_raw}
                    else:
                        # Compléter le nom si la première occurrence était vide
                        if name_raw and not existing_by_tube[tube_key].get("Nom du spectre"):
                            existing_by_tube[tube_key]["Nom du spectre"] = name_raw
                else:
                    # Conserver les lignes hors "tubes par défaut" (ne pas les perdre)
                    if tube_raw or name_raw:
                        extra_rows.append({"Nom du spectre": name_raw, "Tube": tube_raw})

            aligned_rows: list[dict] = []
            for i, tube_label in enumerate(tube_labels_default):
                tube_key = self._tube_merge_key_for_mapping(tube_label, n_tubes)
                existing = existing_by_tube.get(tube_key, {})
                name = existing.get("Nom du spectre") or f"{manip_name}_{start_index + i:02d}"
                existing_tube = existing.get("Tube") or ""
                # Forcer l'affichage du dernier tube comme "Contrôle" si le duplicat est actif,
                # et inversement forcer "Tube N" si le duplicat n'est plus demandé.
                if self._is_control_label(tube_label):
                    tube_val = "Contrôle"
                else:
                    tube_val = tube_label if self._is_control_label(existing_tube) else (existing_tube or tube_label)
                aligned_rows.append({"Nom du spectre": name, "Tube": tube_val})

            # Si le nombre de tubes change, le dernier "Contrôle" peut changer d'index et
            # récupérer un ancien nom déjà utilisé (ex: ..._13 en Tube 13 ET en Contrôle).
            # Dans ce cas, on renumérote automatiquement comme le bouton Reset.
            names_seen: set[str] = set()
            has_dup = False
            for r in aligned_rows:
                nm = str(r.get("Nom du spectre", "")).strip()
                if not nm:
                    continue
                if nm in names_seen:
                    has_dup = True
                    break
                names_seen.add(nm)

            need_renumber = has_dup or (len(existing_by_tube) != len(desired_keys))
            if need_renumber:
                for i in range(len(aligned_rows)):
                    aligned_rows[i]["Nom du spectre"] = f"{manip_name}_{start_index + i:02d}"

            mapping_rows = pd.DataFrame(aligned_rows, columns=default_cols)
            if extra_rows:
                mapping_rows = pd.concat(
                    [mapping_rows, pd.DataFrame(extra_rows, columns=default_cols)],
                    ignore_index=True,
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
            reset_df=reset_df,
            reset_button_text="Reset (valeurs par défaut)",
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

    def _apply_spectrum_names_to_df_map(self, manip_name: str, start_index: int) -> None:
        """Met à jour la colonne 'Nom du spectre' dans df_map sans modifier la colonne 'Tube'.

        Règles :
        - On ne touche jamais aux lignes d'en-tête (Nom de la manip, Date, etc.).
        - Si une ligne interne d'en-tête "Nom du spectre" / "Tube" existe, on met à jour uniquement
        les lignes situées après cette ligne.
        - Si cette ligne n'existe pas, on met à jour uniquement les lignes qui ressemblent à un nom
        de spectre (pattern *_NN).
        - Si df_map n'existe pas encore, on crée un mapping minimal avec tubes par défaut.
        """
        manip_name = (manip_name or "").strip()
        if not manip_name:
            return

        # --- Création minimaliste si df_map n'existe pas encore ---
        if self.df_map is None or not isinstance(self.df_map, pd.DataFrame) or self.df_map.empty:
            tube_labels_default = self._default_tube_labels_for_mapping()
            n = len(tube_labels_default)
            names = [f"{manip_name}_{i:02d}" for i in range(start_index, start_index + n)]
            self.df_map = pd.DataFrame(
                {
                    "Nom du spectre": ["Nom du spectre"] + names,
                    "Tube": ["Tube"] + tube_labels_default,
                }
            )
            return

        df = self.df_map.copy()
        if not {"Nom du spectre", "Tube"}.issubset(df.columns):
            return

        # Cherche la ligne d'en-tête interne "Nom du spectre" / "Tube"
        hdr_idx = None
        for ridx, row in df.iterrows():
            left = str(row.get("Nom du spectre", "")).strip()
            right = str(row.get("Tube", "")).strip()
            if left == "Nom du spectre" and right == "Tube":
                hdr_idx = ridx
                break

        # Déterminer les indices des lignes à mettre à jour
        if hdr_idx is not None:
            mapping_idx = [i for i in df.index if i > hdr_idx]
        else:
            # Fallback : ne mettre à jour que les lignes qui ressemblent à un nom de spectre (ex: XYZ_00)
            col_ns = df["Nom du spectre"].astype(str).str.strip()
            mask = col_ns.str.match(r"^.+_\d+$", na=False)
            mapping_idx = df.index[mask].tolist()

        if not mapping_idx:
            return

        names = [f"{manip_name}_{i:02d}" for i in range(start_index, start_index + len(mapping_idx))]
        for k, ridx in enumerate(mapping_idx):
            df.at[ridx, "Nom du spectre"] = names[k]

        # Important : ne jamais toucher à la colonne Tube
        self.df_map = df
        
    def _on_manip_name_changed(self, *args) -> None:
        """Met à jour les noms de spectres générés automatiquement.

        Cette méthode est connectée à la fois à :
        - QLineEdit.textChanged (str)
        - QSpinBox.valueChanged (int)

        Donc on IGNORE l'argument du signal et on lit directement les widgets.
        """
        manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        start_index = int(self.spin_spec_start.value()) if hasattr(self, "spin_spec_start") else 0

        if not manip_name:
            return

        self._apply_spectrum_names_to_df_map(manip_name, start_index)

        self.map_dirty = True
        self._refresh_button_states()

    def _on_validate_volume_preset_clicked(self) -> None:
        """Applique le modèle sélectionné au tableau des volumes (df_comp)."""
        name = self.combo_volume_preset.currentText().strip()
        if not name or name not in self.VOLUME_PRESETS:
            QMessageBox.warning(self, "Modèle introuvable", "Le modèle sélectionné est introuvable.")
            return

        self.df_comp = self.VOLUME_PRESETS[name].copy()
        self.lbl_status_comp.setText(
            f"Tableau des volumes : {self.df_comp.shape[0]} ligne(s), {self.df_comp.shape[1]} colonne(s)"
        )

        QMessageBox.information(
            self,
            "Modèle appliqué",
            f"Le modèle '{name}' a été appliqué au tableau des volumes.\n\n"
            "Vous pouvez maintenant ouvrir 'Créer / éditer le tableau des volumes' pour l'ajuster.",
        )
        self._sync_df_map_with_df_comp()

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

    def _on_generate_volumes_clicked(self) -> None:
        """Ouvre le dialogue de génération des volumes selon une distribution gaussienne."""

        dlg = GaussianVolumesDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return

        params = dlg.params()
        fixed_stock = dlg.fixed_solution_stocks()

        try:
            df_gen, meta = sers_gaussian_volumes(**params)
        except Exception as e:
            QMessageBox.critical(self, "Erreur génération", f"Impossible de générer les volumes :\n{e}")
            return

        # Toujours définir tube_cols une seule fois
        n_tubes = int(df_gen["Tube"].max())
        tube_cols = [f"Tube {i}" for i in range(1, n_tubes + 1)]

        # --- Initialiser df_comp si aucun modèle n'a été validé ---
        if self.df_comp is None or not isinstance(self.df_comp, pd.DataFrame) or self.df_comp.empty:
            cC, uC = fixed_stock.get("Solution C", (2.0, "µM"))
            cD, uD = fixed_stock.get("Solution D", (0.5, "% en masse"))
            cE, uE = fixed_stock.get("Solution E", (1.0, "mM"))
            cF, uF = fixed_stock.get("Solution F", (1.0, "mM"))
            base_rows = [
                {"Réactif": "Solution A", "Concentration": 4.0, "Unité": "mM"},
                {"Réactif": "Solution B", "Concentration": float(params["C_titrant_uM"]), "Unité": "µM"},
                {"Réactif": "Solution C", "Concentration": float(cC), "Unité": str(uC)},
                {"Réactif": "Solution D", "Concentration": float(cD), "Unité": str(uD)},
                {"Réactif": "Solution E", "Concentration": float(cE), "Unité": str(uE)},
                {"Réactif": "Solution F", "Concentration": float(cF), "Unité": str(uF)},
                {"Réactif": "Echantillon", "Concentration": float(params["C0_nM"]), "Unité": "nM"},
            ]
            df = pd.DataFrame(base_rows)
            for c in tube_cols:
                df[c] = 0.0

            # Supprimer les colonnes Tube au-delà de n_tubes (pour que la taille corresponde exactement)
            existing = self._get_tube_columns(df)
            extra = [c for c in existing if c not in tube_cols]
            if extra:
                df = df.drop(columns=extra)
            self.df_comp = df

        # Copie de travail
        df = self.df_comp.copy()

        # Harmoniser les colonnes de concentration/unité si besoin
        if "Concentration" not in df.columns and "Conc" in df.columns:
            df = df.rename(columns={"Conc": "Concentration"})
        if "Unité" not in df.columns and "Unit" in df.columns:
            df = df.rename(columns={"Unit": "Unité"})
        for c in ("Concentration", "Unité"):
            if c not in df.columns:
                df[c] = pd.NA

        # Vérif structure minimale
        if "Réactif" not in df.columns:
            QMessageBox.critical(self, "Erreur", "La colonne 'Réactif' est introuvable dans le tableau des volumes.")
            return

        # Aligner les colonnes Tube exactement sur le nombre demandé
        existing_tube_cols = self._get_tube_columns(df)
        extra = [c for c in existing_tube_cols if c not in tube_cols]
        if extra:
            df = df.drop(columns=extra)
        for c in tube_cols:
            if c not in df.columns:
                df[c] = 0.0

        def ensure_row(reactif_name: str) -> int:
            mask = df["Réactif"].astype(str).str.strip().str.lower() == reactif_name.strip().lower()
            if mask.any():
                return int(df.index[mask][0])
            new = {col: pd.NA for col in df.columns}
            new["Réactif"] = reactif_name
            df.loc[len(df)] = new
            return int(df.index[-1])

        # A (tampon) et B (titrant)
        idx_A = ensure_row("Solution A")
        idx_B = ensure_row("Solution B")
        idx_sample = ensure_row("Echantillon")
        idx_C = ensure_row("Solution C")
        idx_D = ensure_row("Solution D")
        idx_E = ensure_row("Solution E")
        idx_F = ensure_row("Solution F")

        # Renseigner automatiquement les concentrations stock (saisie utilisateur dans le dialogue)
        df.at[idx_sample, "Concentration"] = float(params["C0_nM"])
        df.at[idx_sample, "Unité"] = "nM"

        df.at[idx_B, "Concentration"] = float(params["C_titrant_uM"])
        df.at[idx_B, "Unité"] = "µM"

        # Solutions fixes : concentration stock + unité (saisie dans le dialogue)
        cC, uC = fixed_stock.get("Solution C", (2.0, "µM"))
        cD, uD = fixed_stock.get("Solution D", (0.5, "% en masse"))
        cE, uE = fixed_stock.get("Solution E", (1.0, "mM"))
        cF, uF = fixed_stock.get("Solution F", (1.0, "mM"))

        df.at[idx_C, "Concentration"] = float(cC)
        df.at[idx_C, "Unité"] = str(uC)

        df.at[idx_D, "Concentration"] = float(cD)
        df.at[idx_D, "Unité"] = str(uD)

        df.at[idx_E, "Concentration"] = float(cE)
        df.at[idx_E, "Unité"] = str(uE)

        df.at[idx_F, "Concentration"] = float(cF)
        df.at[idx_F, "Unité"] = str(uF)

        for _, r in df_gen.iterrows():
            col = f"Tube {int(r['Tube'])}"
            df.at[idx_B, col] = float(r["V_B_titrant_uL"])
            df.at[idx_A, col] = float(r["V_A_tampon_uL"])

        # Fixes : ne remplit que si les lignes existent
        fixed_map = {
            "Echantillon": params["V_echant_uL"],
            "Solution C": params["V_Solution_C_uL"],
            "Solution D": params["V_Solution_D_uL"],
            "Solution E": params["V_Solution_E_uL"],
            "Solution F": params["V_Solution_F_uL"],
        }

        for reactif, val in fixed_map.items():
            mask = df["Réactif"].astype(str).str.strip().str.lower() == reactif.strip().lower()
            if not mask.any():
                continue
            ridx = int(df.index[mask][0])
            for c in tube_cols:
                df.at[ridx, c] = float(val)

        self.df_comp = df

        QMessageBox.information(
            self,
            "Volumes générés",
            "Les volumes ont été générés et appliqués au tableau des volumes.\n\n"
            f"Veq ≈ {meta.get('Veq_uL', 0.0):.2f} µL\n\n"
            f"V_AB = {meta.get('V_AB_uL', 0.0):.2f} µL\n",
        )

        self._sync_df_map_with_df_comp(duplicate_last_override=bool(params.get("duplicate_last", True)))

    def _get_tube_columns(self, df: pd.DataFrame) -> list[str]:
        """Retourne les colonnes 'Tube N' triées par N."""
        if not isinstance(df, pd.DataFrame) or df.empty:
            return []
        tube_cols = [c for c in df.columns if isinstance(c, str) and c.strip().lower().startswith("tube ")]

        def _tube_key(name: str) -> int:
            m = re.search(r"(\d+)", str(name))
            return int(m.group(1)) if m else 10**9

        tube_cols = sorted(tube_cols, key=_tube_key)
        return [str(c).strip() for c in tube_cols]

    def _infer_n_tubes_for_mapping(self) -> int:
        """Déduit le nombre de tubes à afficher dans la correspondance spectres↔tubes.

        Priorité : le tableau des volumes (df_comp) s'il existe.
        Fallback : 10 tubes.
        """
        fallback = 10
        if self.df_comp is None or not isinstance(self.df_comp, pd.DataFrame) or self.df_comp.empty:
            return fallback

        tube_cols = self._get_tube_columns(self.df_comp)
        if not tube_cols:
            return fallback

        max_n = 0
        for c in tube_cols:
            m = re.search(r"(\d+)", str(c))
            if m:
                max_n = max(max_n, int(m.group(1)))

        return max(1, max_n or len(tube_cols) or fallback)

    def _sync_df_map_with_df_comp(self, duplicate_last_override: bool | None = None) -> None:
        """Synchronise automatiquement la correspondance spectres↔tubes avec le tableau des volumes.

        Objectifs :
        - le nombre de lignes de correspondance suit le nombre de colonnes 'Tube N'
        - le dernier tube est 'Contrôle' uniquement si un duplicat est détecté
        - pas de doublons de noms de spectres (renumérotation si nécessaire)

        Cette méthode n'ouvre pas de dialogue : elle met à jour self.df_map en place.
        """
        default_cols = ["Nom du spectre", "Tube"]

        # Récupération des infos d'en-tête depuis l'UI
        try:
            manip_name = self.edit_manip.text().strip() if hasattr(self, "edit_manip") else ""
        except Exception:
            manip_name = ""

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

        # Tubes attendus d'après df_comp
        n_tubes = self._infer_n_tubes_for_mapping()
        tube_labels_default = self._default_tube_labels_for_mapping(
            n_tubes,
            duplicate_last_override=duplicate_last_override,
        )
        start_index = int(self.spin_spec_start.value()) if hasattr(self, "spin_spec_start") else 0

        def _make_mapping_rows(labels: list[str]) -> pd.DataFrame:
            prefix = manip_name if manip_name else "Spectrum"
            names = [f"{prefix}_{i:02d}" for i in range(start_index, start_index + len(labels))]
            return pd.DataFrame({"Nom du spectre": names, "Tube": labels}, columns=default_cols)

        # Extraire la partie "mapping" existante si possible
        mapping_rows = None
        if self.df_map is not None and isinstance(self.df_map, pd.DataFrame) and not self.df_map.empty:
            existing = self.df_map.copy()
            if {"Nom du spectre", "Tube"}.issubset(existing.columns):
                try:
                    col_ns = existing["Nom du spectre"].astype(str).str.strip()
                    col_t = existing["Tube"].astype(str).str.strip()
                    hdr_mask = (col_ns == "Nom du spectre") & (col_t == "Tube")
                    if hdr_mask.any():
                        last_hdr_idx = hdr_mask[hdr_mask].index[-1]
                        mapping_rows = existing.loc[existing.index > last_hdr_idx, ["Nom du spectre", "Tube"]].copy()
                    else:
                        mapping_rows = existing[["Nom du spectre", "Tube"]].copy()
                    mapping_rows = mapping_rows.reset_index(drop=True)
                except Exception:
                    mapping_rows = None

        if mapping_rows is None or mapping_rows.empty:
            mapping_rows = _make_mapping_rows(tube_labels_default)
        else:
            def _clean_text(v) -> str:
                if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
                    return ""
                s = str(v).strip()
                return "" if s == "<NA>" else s

            existing_by_tube: dict[str, dict[str, str]] = {}
            desired_keys = {self._tube_merge_key_for_mapping(t, n_tubes) for t in tube_labels_default}

            for _, row in mapping_rows.iterrows():
                tube_raw = _clean_text(row.get("Tube", ""))
                name_raw = _clean_text(row.get("Nom du spectre", ""))
                tube_key = self._tube_merge_key_for_mapping(tube_raw, n_tubes)
                if not tube_key:
                    continue
                if tube_key in desired_keys:
                    if tube_key not in existing_by_tube:
                        existing_by_tube[tube_key] = {"Nom du spectre": name_raw, "Tube": tube_raw}
                    else:
                        if name_raw and not existing_by_tube[tube_key].get("Nom du spectre"):
                            existing_by_tube[tube_key]["Nom du spectre"] = name_raw

            aligned_rows: list[dict] = []
            prefix = manip_name if manip_name else "Spectrum"
            for i, tube_label in enumerate(tube_labels_default):
                tube_key = self._tube_merge_key_for_mapping(tube_label, n_tubes)
                existing = existing_by_tube.get(tube_key, {})
                name = existing.get("Nom du spectre") or f"{prefix}_{start_index + i:02d}"
                existing_tube = existing.get("Tube") or ""

                if self._is_control_label(tube_label):
                    tube_val = "Contrôle"
                else:
                    tube_val = tube_label if self._is_control_label(existing_tube) else (existing_tube or tube_label)
                aligned_rows.append({"Nom du spectre": name, "Tube": tube_val})

            # Renumérote automatiquement si doublons / mismatch (même comportement que l'ouverture du dialogue)
            names_seen: set[str] = set()
            has_dup = False
            for r in aligned_rows:
                nm = str(r.get("Nom du spectre", "")).strip()
                if not nm:
                    continue
                if nm in names_seen:
                    has_dup = True
                    break
                names_seen.add(nm)

            need_renumber = has_dup or (len(existing_by_tube) != len(desired_keys))
            if need_renumber:
                for i in range(len(aligned_rows)):
                    aligned_rows[i]["Nom du spectre"] = f"{prefix}_{start_index + i:02d}"

            mapping_rows = pd.DataFrame(aligned_rows, columns=default_cols)

        self.df_map = pd.concat([df_header, blank_row, internal_header, mapping_rows], ignore_index=True)
        self.map_dirty = False

        if hasattr(self, "lbl_status_map"):
            self.lbl_status_map.setText(
                f"Tableau de correspondance : {self.df_map.shape[0]} ligne(s), {self.df_map.shape[1]} colonne(s)"
            )

        self._refresh_button_states()

    def _infer_duplicate_last_for_mapping(self, n_tubes: int | None = None) -> bool:
        """Devine si le dernier tube est un duplicat (contrôle).

        Heuristique : si, dans le tableau des volumes (df_comp), la ligne "Solution B"
        a les deux derniers volumes identiques (Tube N-1 == Tube N), alors on considère
        que le dernier tube est un réplicat.

        Si l'information n'est pas disponible, on suppose True (comportement historique).
        """
        if n_tubes is None:
            n = self._infer_n_tubes_for_mapping()
        else:
            n = max(1, int(n_tubes))

        if n < 2:
            return True

        if self.df_comp is None or not isinstance(self.df_comp, pd.DataFrame) or self.df_comp.empty:
            return True

        df = self.df_comp
        if "Réactif" not in df.columns:
            return True

        # Trouver la ligne Solution B (titrant)
        reactif = df["Réactif"].astype(str).str.strip().str.lower()
        mask = reactif == "solution b"
        if not mask.any():
            mask = reactif.str.startswith("solution b")
        if not mask.any():
            return True
        ridx = int(mask[mask].index[0])

        # Colonnes des 2 derniers tubes
        col_prev = f"Tube {n - 1}"
        col_last = f"Tube {n}"
        if col_prev in df.columns and col_last in df.columns:
            cols = [col_prev, col_last]
        else:
            tube_cols = self._get_tube_columns(df)
            if len(tube_cols) < 2:
                return True
            cols = tube_cols[-2:]

        v_prev = self._to_float(df.at[ridx, cols[0]])
        v_last = self._to_float(df.at[ridx, cols[1]])
        if v_prev is None or v_last is None:
            return True

        return abs(float(v_last) - float(v_prev)) <= 1e-9

    @staticmethod
    def _norm_text_key(s: str) -> str:
        return (
            ("" if s is None else str(s))
            .strip()
            .lower()
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
        )

    def _is_control_label(self, label: str) -> bool:
        """Vrai si le label correspond au tube 'Contrôle' (réplicat)."""
        key = self._norm_text_key(label)
        return key in {"controle", "control"}

    def _tube_merge_key_for_mapping(self, label: str, n_tubes: int | None = None) -> str:
        """Retourne la clé de tube utilisée pour relier la correspondance au tableau des volumes.

        - "A13" / "13" / "Tube 13" -> "Tube 13"
        - "Contrôle"              -> "Tube N" (le dernier tube du tableau des volumes)
        - autres                  -> label normalisé (trim)
        """
        n = self._infer_n_tubes_for_mapping() if n_tubes is None else max(1, int(n_tubes))
        norm = self._normalize_tube_label(label)
        if self._is_control_label(norm):
            return f"Tube {n}"
        return norm

    def _default_tube_labels_for_mapping(
        self,
        n_tubes: int | None = None,
        *,
        duplicate_last_override: bool | None = None,
    ) -> list[str]:
        """Liste de tubes par défaut pour la correspondance.

        Convention :
        - 1er spectre : "Contrôle BRB"
        - puis tubes expérimentaux :
            - si duplicat actif : Tube 1..Tube (N-1) + "Contrôle" (réplicat)
            - sinon             : Tube 1..Tube N
        """
        n = self._infer_n_tubes_for_mapping() if n_tubes is None else max(1, int(n_tubes))
        duplicate_last = (
            bool(duplicate_last_override)
            if duplicate_last_override is not None
            else self._infer_duplicate_last_for_mapping(n)
        )
        if n <= 1:
            tubes = ["Contrôle" if duplicate_last else "Tube 1"]
        elif duplicate_last:
            tubes = [f"Tube {i}" for i in range(1, n)] + ["Contrôle"]
        else:
            tubes = [f"Tube {i}" for i in range(1, n + 1)]
        return ["Contrôle BRB"] + tubes
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
            # nM -> M
            if u in ("nm", "nmol/l", "nmol.l-1"):
                return v * 1e-9
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

        # IMPORTANT : on conserve des valeurs numériques (float) pour permettre
        # l'édition + la rétro-propagation vers les concentrations stock.
        # L'affichage à 2 décimales est géré dans TableEditorDialog._load_from_dataframe.
        return conc_df
    
    def _on_show_conc_clicked(self) -> None:
        """Affiche le tableau des concentrations dans les cuvettes (information uniquement)."""

        if self.df_comp is None or self.df_comp.empty:
            QMessageBox.warning(self, "Erreur", "Le tableau des volumes n'est pas défini.")
            return

        conc_df = self._compute_concentration_table()

        dlg = TableEditorDialog(
            "Concentration dans les cuvettes (info)",
            list(conc_df.columns),
            conc_df,
            self,
        )

        # Lecture seule
        dlg.table.setEditTriggers(QTableWidget.NoEditTriggers)

        dlg.exec()

    # ------------------------------------------------------------------
    # Méthode utilitaire optionnelle
    # ------------------------------------------------------------------
    def build_merged_metadata(self) -> pd.DataFrame:
        """
        Construit le DataFrame de métadonnées final utilisé pour :
        - tracés des spectres
        - PCA
        - analyses de pics

        Sources :
        - df_map  : correspondance spectre ↔ tube
        - df_conc : concentrations finales par tube (PRIORITAIRE si défini)
        - df_comp : tableau des volumes (fallback)

        Sortie :
        DataFrame contenant au minimum :
            - Spectrum name
            - Tube
            - Sample description
        + colonnes de concentrations (une par réactif si disponibles)
        + colonnes dérivées pour le titrant ([titrant] (M), [titrant] (mM))
        """

        # ------------------------------------------------------------------
        # 1) Vérifications de base
        # ------------------------------------------------------------------
        if self.df_map is None or self.df_map.empty:
            raise ValueError("Le tableau de correspondance spectres ↔ tubes n'est pas défini.")

        df_map = self.df_map.copy()

        if not {"Nom du spectre", "Tube"}.issubset(df_map.columns):
            raise ValueError(
                "Le tableau de correspondance doit contenir les colonnes "
                "'Nom du spectre' et 'Tube'."
            )

        # Nettoyage texte
        df_map["Nom du spectre"] = df_map["Nom du spectre"].astype(str).str.strip()
        df_map["Tube"] = df_map["Tube"].astype(str).str.strip()

        # Supprimer une éventuelle ligne d'en-tête interne
        header_mask = df_map["Nom du spectre"] == "Nom du spectre"
        if header_mask.any():
            last_header_idx = header_mask[header_mask].index[-1]
            df_map = df_map.loc[df_map.index > last_header_idx].copy()

        # Retirer lignes vides
        df_map = df_map[
            (df_map["Nom du spectre"] != "") &
            (df_map["Tube"] != "")
        ].copy()

        # ------------------------------------------------------------------
        # 2) Base des métadonnées : spectre ↔ tube
        # ------------------------------------------------------------------
        meta = df_map.rename(columns={"Nom du spectre": "Spectrum name"})

        # Clé de jointure vers le tableau des volumes (tubes numérotés)
        n_tubes = self._infer_n_tubes_for_mapping()
        meta["Tube_merge"] = meta["Tube"].apply(lambda s: self._tube_merge_key_for_mapping(s, n_tubes))

        # Description lisible des échantillons
        tube_to_desc = {
            "Contrôle BRB": "Contrôle BRB",
            "Contrôle": "Contrôle",
        }
        meta["Sample description"] = (
            meta["Tube"].map(tube_to_desc).fillna(meta["Tube"])
        )

        # ------------------------------------------------------------------
        # 3) Choix de la source des concentrations
        # ------------------------------------------------------------------
        conc_df = self._compute_concentration_table()

        # ------------------------------------------------------------------
        # 4) Injection des concentrations dans les métadonnées
        # ------------------------------------------------------------------
        if "Réactif" in conc_df.columns:
            tube_cols = [c for c in conc_df.columns if c != "Réactif"]

            if tube_cols:
                # Passage au format long : (Tube, Réactif, C_M)
                conc_long = conc_df.melt(
                    id_vars="Réactif",
                    value_vars=tube_cols,
                    var_name="Tube",
                    value_name="C_M",
                )

                # Pivot : une colonne par réactif
                conc_wide = conc_long.pivot(
                    index="Tube",
                    columns="Réactif",
                    values="C_M",
                )

                # Sécuriser les types numériques
                conc_wide = conc_wide.apply(pd.to_numeric, errors="coerce")

                # Jointure par tube (en utilisant Tube_merge pour gérer "Contrôle" = dernier Tube N)
                conc_wide = conc_wide.reset_index().rename(columns={"Tube": "Tube_merge"})
                meta = meta.merge(conc_wide, on="Tube_merge", how="left")
                meta = meta.drop(columns=["Tube_merge"])

                # ----------------------------------------------------------
                # 5) Détection automatique du titrant
                # ----------------------------------------------------------
                titrant_candidates = []
                for col in conc_wide.columns:
                    name_low = str(col).lower()
                    if (
                        "titrant" in name_low
                        or "egta" in name_low
                        or name_low in {"solution b", "solutionb", "sol b", "b"}
                    ):
                        titrant_candidates.append(col)

                if titrant_candidates:
                    titrant = titrant_candidates[0]
                    meta[titrant] = pd.to_numeric(meta[titrant], errors="coerce")
                    meta["[titrant] (M)"] = meta[titrant]
                    meta["[titrant] (mM)"] = meta[titrant] * 1e3

        # ------------------------------------------------------------------
        # 6) Nettoyage final
        # ------------------------------------------------------------------
        # Exclure explicitement les contrôles BRB des analyses
        meta = meta[~meta["Tube"].isin(["Tube BRB"])].copy()

        # Nettoyage : ne pas exposer la clé interne de jointure si elle existe
        if "Tube_merge" in meta.columns:
            meta = meta.drop(columns=["Tube_merge"])

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
        """Enregistre df_comp et df_map dans un fichier Excel."""
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
        """Charge df_comp et df_map à partir d'un fichier Excel précédemment sauvegardé."""
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
    def _on_edit_comp_clicked(self) -> None:
        """Ouvre le dialogue d'édition du tableau des volumes.

        Règles :
        - Si un tableau existe déjà (self.df_comp), on le ré-ouvre tel quel.
        - Sinon, on propose un modèle par défaut ou pré-rempli selon le preset sélectionné.
        - À la validation, on enregistre dans self.df_comp.
        - Si l'utilisateur n'a PAS édité les concentrations (conc_user_overridden=False),
          on recalcule self.df_conc depuis les volumes.
        """

        # Colonnes de base
        base_cols = ["Réactif", "Concentration", "Unité"]

        # 1) Si un tableau existe déjà, on le ré-ouvre tel quel
        if self.df_comp is not None and not self.df_comp.empty:
            initial_df = self.df_comp.copy()

        # 2) Sinon, on charge le modèle sélectionné
        else:
            preset_name = self.combo_volume_preset.currentText()
            preset_df = self.VOLUME_PRESETS.get(preset_name)
            if preset_df is None:
                QMessageBox.warning(self, "Modèle inconnu", "Le modèle sélectionné est introuvable.")
                return
            initial_df = preset_df.copy()

        # Harmonise les alias fréquents pour les colonnes standard (Conc/Unit -> Concentration/Unité)
        def _norm_col(col: str) -> str:
            return (
                str(col)
                .strip()
                .lower()
                .replace("é", "e")
                .replace("è", "e")
                .replace("ê", "e")
            )

        # Renommer si les colonnes standard manquent
        rename_map: dict[str, str] = {}
        has_conc = "Concentration" in initial_df.columns
        has_unit = "Unité" in initial_df.columns

        for col in list(initial_df.columns):
            n = _norm_col(col)
            if n == "conc" and not has_conc:
                rename_map[col] = "Concentration"
                has_conc = True
            if n in ("unit", "units", "unite") and not has_unit:
                rename_map[col] = "Unité"
                has_unit = True

        if rename_map:
            initial_df = initial_df.rename(columns=rename_map)

        # Fusionner d'éventuelles colonnes alias supplémentaires avec les colonnes standard
        conc_aliases = [c for c in initial_df.columns if _norm_col(c) == "conc" and c != "Concentration"]
        for alias in conc_aliases:
            if "Concentration" not in initial_df.columns:
                initial_df = initial_df.rename(columns={alias: "Concentration"})
            else:
                initial_df["Concentration"] = initial_df["Concentration"].fillna(initial_df[alias])
                initial_df = initial_df.drop(columns=[alias])

        unit_aliases = [
            c for c in initial_df.columns if _norm_col(c) in ("unit", "units", "unite") and c != "Unité"
        ]
        for alias in unit_aliases:
            if "Unité" not in initial_df.columns:
                initial_df = initial_df.rename(columns={alias: "Unité"})
            else:
                initial_df["Unité"] = initial_df["Unité"].fillna(initial_df[alias])
                initial_df = initial_df.drop(columns=[alias])

        # Colonnes de tubes dynamiques : on reprend tout ce qui existe déjà pour ne rien perdre
        tube_cols = self._get_tube_columns(initial_df)
        if not tube_cols:
            tube_cols = [f"Tube {i}" for i in range(1, 11)]

        default_cols = base_cols + tube_cols

        # Sécurité : garantir les colonnes standard
        for c in default_cols:
            if c not in initial_df.columns:
                initial_df[c] = pd.NA

        # Conserver d'éventuelles colonnes supplémentaires ajoutées par l'utilisateur
        extra_cols = [c for c in initial_df.columns if c not in default_cols]
        initial_df = initial_df[default_cols + extra_cols]

        dlg = TableEditorDialog(
            title="Tableau des volumes",
            columns=list(initial_df.columns),
            initial_df=initial_df,
            parent=self,
        )
        if dlg.exec() != QDialog.Accepted:
            return

        df = dlg.to_dataframe()
        if df.empty:
            QMessageBox.warning(self, "Tableau vide", "Le tableau des volumes est vide.")
            return

        # Nettoyage léger
        if "Réactif" in df.columns:
            df["Réactif"] = df["Réactif"].astype(str).str.strip()

        self.df_comp = df
        self.lbl_status_comp.setText(
            f"Tableau des volumes : {df.shape[0]} ligne(s), {df.shape[1]} colonne(s)"
        )

        self._sync_df_map_with_df_comp()
