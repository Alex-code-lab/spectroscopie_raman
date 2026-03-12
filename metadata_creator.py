

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
    QAbstractItemView,
    QInputDialog,
    QMenu,
)

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QBrush

import metadata_model as mm


# Affichage numérique: précision et suppression de l'écriture scientifique.
NUM_DISPLAY_PRECISION = 12
DEFAULT_VTOT_UL = 3090.0
SUM_TARGET_TOL = 1e-3
MOLAR_UNIT_FACTORS = {
    "M": 1.0,
    "mM": 1e-3,
    "µM": 1e-6,
    "nM": 1e-9,
}


# Affichage numérique: précision et suppression de l'écriture scientifique.
NUM_DISPLAY_PRECISION = 12


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
        sum_row: bool = False,
        sum_row_label: str = "Total",
        sum_target: float | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)

        self._columns = list(columns)
        # self.conc_user_overridden = False
        self._reset_df = reset_df.copy() if reset_df is not None else None
        self._reset_button_text = reset_button_text
        self._sum_row_enabled = bool(sum_row)
        self._sum_row_label = str(sum_row_label)
        self._sum_target = float(sum_target) if sum_target is not None else None
        self._sum_row_idx: int | None = None
        self._sum_updating = False
        self._active_tube_col: int | None = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(title))

        # --- Table d'édition ---
        self.table = QTableWidget(self)
        self.table.setColumnCount(len(self._columns))
        self.table.setHorizontalHeaderLabels(self._columns)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.itemChanged.connect(self._on_item_changed)

        # Remplissage initial : soit à partir du DataFrame, soit quelques lignes vides
        if initial_df is not None and not initial_df.empty:
            self._load_from_dataframe(initial_df)
        else:
            self.table.setRowCount(8)  # quelques lignes vides par défaut
            if self._sum_row_enabled:
                self._append_sum_row()

        layout.addWidget(self.table, 1)

        # --- Boutons de modification du tableau ---
        btns_edit = QHBoxLayout()
        self.btn_add_row = QPushButton("+ Ligne", self)
        self.btn_del_row = QPushButton("− Ligne", self)
        self.btn_add_col = QPushButton("Ajouter un tube", self)
        self.btn_del_col = QPushButton("Supprimer le tube sélectionné", self)
        self.btn_dup_col = QPushButton("Dupliquer le tube sélectionné", self)
        self.btn_rename_col = QPushButton("Renommer colonne…", self)
        self.btn_del_col.setToolTip("Cliquez d'abord sur l'en-tête du tube (ligne de titres) pour le sélectionner, puis cliquez ce bouton.")
        self.btn_dup_col.setToolTip("Cliquez d'abord sur l'en-tête du tube (ligne de titres) pour le sélectionner, puis cliquez ce bouton.")
        btns_edit.addWidget(self.btn_add_row)
        btns_edit.addWidget(self.btn_del_row)
        btns_edit.addSpacing(20)
        btns_edit.addWidget(self.btn_add_col)
        btns_edit.addWidget(self.btn_del_col)
        btns_edit.addWidget(self.btn_dup_col)
        btns_edit.addWidget(self.btn_rename_col)

        btns_edit.addSpacing(20)
        self.btn_reset = QPushButton(self._reset_button_text, self)
        self.btn_reset.setEnabled(self._reset_df is not None)
        btns_edit.addWidget(self.btn_reset)

        btns_edit.addStretch(1)
        layout.addLayout(btns_edit)

        # --- Option saisie manuelle (désactive la sync A↔B) ---
        mode_row = QHBoxLayout()
        self.chk_manual_mode = QCheckBox("Saisie manuelle (désactiver la synchronisation Solution A ↔ Solution B)", self)
        self.chk_manual_mode.setToolTip(
            "Coché : vous pouvez saisir librement les volumes de Solution A et Solution B "
            "sans qu'ils se recalculent mutuellement.\n"
            "Décoché (défaut) : modifier A ajuste B automatiquement pour conserver la somme constante."
        )
        mode_row.addWidget(self.chk_manual_mode)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        self.btn_add_row.clicked.connect(self._add_row)
        self.btn_del_row.clicked.connect(self._del_row)
        self.btn_add_col.clicked.connect(self._add_col)
        self.btn_del_col.clicked.connect(self._del_col)
        self.btn_dup_col.clicked.connect(self._duplicate_col)
        self.btn_rename_col.clicked.connect(self._rename_col)
        self.btn_reset.clicked.connect(self._reset_table)

        # Clic simple sur un en-tête → sélectionner la colonne
        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)
        # Double-clic sur un en-tête → renommer
        self.table.horizontalHeader().sectionDoubleClicked.connect(self._rename_col)
        # Clic droit sur un en-tête → menu contextuel
        self.table.horizontalHeader().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.horizontalHeader().customContextMenuRequested.connect(self._on_header_context_menu)

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

        self.table.setRowCount(len(df) + (1 if self._sum_row_enabled else 0))
        self.table.setColumnCount(len(self._columns))
        self.table.setHorizontalHeaderLabels(self._columns)

        for i in range(len(df)):
            for j, col in enumerate(self._columns):
                val = df.iloc[i, j]
                if pd.isna(val):
                    txt = ""
                else:
                    # Affichage lisible : pas d'écriture scientifique, précision configurable.
                    if isinstance(val, (float, int)):
                        txt = self._format_number(float(val))
                    else:
                        txt = str(val)
                item = QTableWidgetItem(txt)
                self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()
        if self._sum_row_enabled:
            self._sum_row_idx = len(df)
            self._ensure_sum_row_items()
            self._set_sum_row_label()
            self._refresh_sum_row()

    def _format_number(self, val: float) -> str:
        try:
            return np.format_float_positional(
                float(val),
                precision=NUM_DISPLAY_PRECISION,
                unique=False,
                trim="-",
            )
        except Exception:
            return str(val)

    def _parse_float(self, txt: str | None) -> float | None:
        return mm.to_float(txt)

    def _get_tube_column_indices(self) -> list[int]:
        idxs: list[int] = []
        for j in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(j)
            name = header.text().strip() if header is not None else ""
            if name.lower().startswith("tube "):
                idxs.append(j)
        return idxs

    def _ensure_sum_row_items(self) -> None:
        if self._sum_row_idx is None:
            return
        for j in range(self.table.columnCount()):
            item = self.table.item(self._sum_row_idx, j)
            if item is None:
                item = QTableWidgetItem("")
                self.table.setItem(self._sum_row_idx, j, item)
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

    def _set_sum_row_label(self) -> None:
        if self._sum_row_idx is None or self.table.columnCount() == 0:
            return
        label_col = 0
        for j in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(j)
            name = header.text().strip().lower() if header is not None else ""
            if name in ("réactif", "reactif"):
                label_col = j
                break
        item = self.table.item(self._sum_row_idx, label_col)
        if item is None:
            item = QTableWidgetItem("")
            self.table.setItem(self._sum_row_idx, label_col, item)
        item.setText(self._sum_row_label)
        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

    def _append_sum_row(self) -> None:
        self._sum_row_idx = self.table.rowCount()
        self.table.insertRow(self._sum_row_idx)
        self._ensure_sum_row_items()
        self._set_sum_row_label()
        self._refresh_sum_row()

    def _refresh_sum_row(self) -> None:
        if not self._sum_row_enabled or self._sum_row_idx is None:
            return
        if self._sum_updating:
            return
        self._sum_updating = True
        try:
            self.table.blockSignals(True)
            self._ensure_sum_row_items()
            self._set_sum_row_label()
            tube_cols = self._get_tube_column_indices()
            # Indice de la colonne "Pas (µL)" pour les lignes compte-goutte
            pas_col_idx: int | None = None
            for j in range(self.table.columnCount()):
                hdr = self.table.horizontalHeaderItem(j)
                if hdr is not None and hdr.text().strip() == "Pas (µL)":
                    pas_col_idx = j
                    break
            # Nettoie les fonds
            for j in range(self.table.columnCount()):
                item = self.table.item(self._sum_row_idx, j)
                if item is not None:
                    item.setBackground(QBrush())
            for j in tube_cols:
                total = 0.0
                for r in range(self._sum_row_idx):
                    item = self.table.item(r, j)
                    val = self._parse_float(item.text() if item is not None else "")
                    if val is None:
                        continue
                    # Si ligne compte-goutte, multiplier les n_gouttes par le pas
                    if pas_col_idx is not None:
                        pas_item = self.table.item(r, pas_col_idx)
                        pas = self._parse_float(pas_item.text() if pas_item is not None else "") or 0.0
                        if pas > 0:
                            val = val * pas
                    total += val
                item = self.table.item(self._sum_row_idx, j)
                if item is None:
                    item = QTableWidgetItem("")
                    self.table.setItem(self._sum_row_idx, j, item)
                item.setText(self._format_number(total))
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                if self._sum_target is not None and abs(total - self._sum_target) <= SUM_TARGET_TOL:
                    item.setBackground(QBrush(QColor(198, 239, 206)))
        finally:
            self.table.blockSignals(False)
            self._sum_updating = False

    def _auto_resize_columns(self) -> None:
        """Ajuste automatiquement la largeur des colonnes au contenu."""
        self.table.resizeColumnsToContents()

    def _on_header_clicked(self, col: int) -> None:
        """Sélectionne visuellement la colonne cliquée et la mémorise."""
        self._active_tube_col = col
        self.table.selectColumn(col)

    def _on_header_context_menu(self, pos) -> None:
        """Menu contextuel sur un en-tête de colonne Tube : dupliquer / supprimer."""
        col = self.table.horizontalHeader().logicalIndexAt(pos)
        if col < 0:
            return
        header = self.table.horizontalHeaderItem(col)
        col_name = header.text().strip() if header is not None else ""
        if not col_name.lower().startswith("tube "):
            return
        self._active_tube_col = col
        self.table.selectColumn(col)
        menu = QMenu(self)
        act_dup = menu.addAction(f"Dupliquer « {col_name} »")
        act_del = menu.addAction(f"Supprimer « {col_name} »")
        action = menu.exec(self.table.horizontalHeader().mapToGlobal(pos))
        if action == act_dup:
            self._duplicate_col()
        elif action == act_del:
            self._del_col()

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        self._auto_resize_columns()
        if not self._sum_row_enabled or self._sum_row_idx is None:
            return
        if self._sum_updating:
            return
        if item.row() == self._sum_row_idx:
            return
        self._refresh_sum_row()

    def _add_row(self) -> None:
        if self._sum_row_enabled and self._sum_row_idx is not None:
            self.table.insertRow(self._sum_row_idx)
            self._sum_row_idx += 1
        else:
            self.table.insertRow(self.table.rowCount())
        if self._sum_row_enabled:
            self._refresh_sum_row()

    def _del_row(self) -> None:
        selected_rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not selected_rows:
            row = self.table.currentRow()
            if row < 0:
                return
            selected_rows = [row]

        for row in selected_rows:
            if self._sum_row_enabled and self._sum_row_idx is not None:
                if row == self._sum_row_idx:
                    continue
                if row < self._sum_row_idx:
                    self._sum_row_idx -= 1
            self.table.removeRow(row)
        if self._sum_row_enabled:
            self._refresh_sum_row()

    def _add_col(self) -> None:
        # Trouver les colonnes "Tube N" existantes et le dernier numéro
        tube_indices = self._get_tube_column_indices()
        if tube_indices:
            last_tube_col = tube_indices[-1]
            last_header = self.table.horizontalHeaderItem(last_tube_col)
            last_name = last_header.text().strip() if last_header else ""
            # Extraire le numéro du dernier tube (ex: "Tube 5" → 5)
            import re as _re
            m = _re.search(r"\d+$", last_name)
            next_num = int(m.group()) + 1 if m else len(tube_indices) + 1
            new_name = f"Tube {next_num}"
        else:
            last_tube_col = None
            new_name = "Tube 1"

        # Insérer la nouvelle colonne
        col_count = self.table.columnCount()
        self.table.insertColumn(col_count)
        self._columns.append(new_name)
        self.table.setHorizontalHeaderLabels(self._columns)

        # Pré-remplir depuis le dernier tube (hors ligne de somme)
        # Les signaux sont bloqués pour éviter que sync_A_B se déclenche
        # au milieu de la copie avec un target_sum incomplet/erroné.
        if last_tube_col is not None:
            n_data_rows = self._sum_row_idx if (self._sum_row_enabled and self._sum_row_idx is not None) else self.table.rowCount()
            self.table.blockSignals(True)
            try:
                for r in range(n_data_rows):
                    src = self.table.item(r, last_tube_col)
                    txt = src.text() if src is not None else ""
                    self.table.setItem(r, col_count, QTableWidgetItem(txt))
            finally:
                self.table.blockSignals(False)

        if self._sum_row_enabled:
            self._ensure_sum_row_items()
            self._refresh_sum_row()

    def _resolve_tube_col(self, action_label: str) -> int | None:
        """Retourne l'indice de la colonne Tube active, ou None avec un message d'erreur."""
        col = self._active_tube_col
        if col is None or col < 0 or col >= self.table.columnCount():
            QMessageBox.information(
                self, action_label,
                "Cliquez d'abord sur l'en-tête d'un tube (la ligne de titres) pour le sélectionner."
            )
            return None
        header = self.table.horizontalHeaderItem(col)
        col_name = header.text().strip() if header is not None else ""
        if not col_name.lower().startswith("tube "):
            QMessageBox.information(
                self, action_label,
                f"La colonne « {col_name} » n'est pas une colonne Tube.\n"
                "Cliquez sur l'en-tête d'un tube pour le sélectionner."
            )
            return None
        return col

    def _del_col(self) -> None:
        col = self._resolve_tube_col("Supprimer un tube")
        if col is None:
            return
        if 0 <= col < len(self._columns):
            self._columns.pop(col)
        self.table.removeColumn(col)
        self.table.setHorizontalHeaderLabels(self._columns)
        self._active_tube_col = None
        if self._sum_row_enabled:
            self._refresh_sum_row()

    def _duplicate_col(self) -> None:
        """Duplique le tube de la colonne actuellement sélectionnée."""
        col = self._resolve_tube_col("Dupliquer un tube")
        if col is None:
            return

        # Numéro du nouveau tube = max existant + 1
        tube_indices = self._get_tube_column_indices()
        last_header = self.table.horizontalHeaderItem(tube_indices[-1])
        last_name = last_header.text().strip() if last_header else ""
        import re as _re
        m = _re.search(r"\d+$", last_name)
        next_num = int(m.group()) + 1 if m else len(tube_indices) + 1
        new_name = f"Tube {next_num}"

        # Insérer juste après la colonne sélectionnée
        new_col = col + 1
        self.table.insertColumn(new_col)
        self._columns.insert(new_col, new_name)
        self.table.setHorizontalHeaderLabels(self._columns)

        # Copier les valeurs (hors ligne de somme)
        n_data_rows = (
            self._sum_row_idx
            if (self._sum_row_enabled and self._sum_row_idx is not None)
            else self.table.rowCount()
        )
        self.table.blockSignals(True)
        try:
            for r in range(n_data_rows):
                src = self.table.item(r, col)
                txt = src.text() if src is not None else ""
                self.table.setItem(r, new_col, QTableWidgetItem(txt))
        finally:
            self.table.blockSignals(False)

        if self._sum_row_enabled:
            self._ensure_sum_row_items()
            self._refresh_sum_row()

    def _rename_col(self, col: int = -1) -> None:
        """Renomme la colonne `col` (ou la colonne actuellement sélectionnée si col=-1)."""
        if col < 0:
            col = self.table.currentColumn()
        if col < 0 or col >= self.table.columnCount():
            QMessageBox.information(self, "Renommer", "Cliquez d'abord sur une colonne pour la sélectionner.")
            return
        header = self.table.horizontalHeaderItem(col)
        current_name = header.text() if header is not None else ""
        new_name, ok = QInputDialog.getText(
            self, "Renommer la colonne", f"Nouveau nom pour « {current_name} » :", text=current_name
        )
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        if 0 <= col < len(self._columns):
            self._columns[col] = new_name
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
        if self._sum_row_enabled and self._sum_row_idx is not None and self._sum_row_idx in df.index:
            df = df.drop(index=self._sum_row_idx)
        # Normalise quelques colonnes si présentes
        if "Tube" in df.columns:
            df["Tube"] = df["Tube"].astype(str).str.strip()
            df = df[df["Tube"] != ""]
        if "Spectrum name" in df.columns:
            df["Spectrum name"] = df["Spectrum name"].astype(str).str.strip()
        return df.reset_index(drop=True)


class MolesDialog(QDialog):
    """Dialogue d'édition des quantités de matière (nmol)."""

    def __init__(self, entries: list[dict], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Quantités de matière")
        self._entries = entries
        self._spins: dict[str, QDoubleSpinBox] = {}

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Éditez les quantités de matière (nmol) calculées à partir des "
                "concentrations et volumes actuels."
            )
        )

        form = QFormLayout()
        for entry in self._entries:
            key = str(entry.get("key", ""))
            label = str(entry.get("label", ""))
            value = entry.get("value", None)
            editable = bool(entry.get("editable", True))
            note = str(entry.get("note", "")).strip()

            row_label = f"{label} (nmol)"
            if editable and value is not None:
                spin = QDoubleSpinBox(self)
                spin.setDecimals(6)
                spin.setRange(0.0, 1e12)
                spin.setValue(float(value))
                if note:
                    spin.setToolTip(note)
                form.addRow(row_label, spin)
                self._spins[key] = spin
            else:
                txt = "N/A"
                if value is not None:
                    txt = str(value)
                le = QLineEdit(txt, self)
                le.setReadOnly(True)
                if note:
                    le.setToolTip(note)
                form.addRow(row_label, le)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.resize(520, 320)

    def values(self) -> dict[str, float]:
        return {k: float(spin.value()) for k, spin in self._spins.items()}



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

        # Case à cocher mode compte‑goutte
        self.chk_dropper = QCheckBox("Utilisation d'un compte‑goutte (pas = 40 µL, marge = 0)", self)
        self.chk_dropper.setChecked(False)
        layout.addWidget(self.chk_dropper)

        box = QGroupBox("Paramètres")
        form = QFormLayout(box)

        self.spin_n = QSpinBox(self); self.spin_n.setRange(2, 200); self.spin_n.setValue(11)
        self.spin_C0 = QDoubleSpinBox(self); self.spin_C0.setDecimals(3); self.spin_C0.setRange(0.0, 1e9); self.spin_C0.setValue(1000.0)
        self.spin_Vech = QDoubleSpinBox(self); self.spin_Vech.setDecimals(1); self.spin_Vech.setRange(0.0, 1e9); self.spin_Vech.setValue(1000.0)
        self.spin_Ctit = QDoubleSpinBox(self); self.spin_Ctit.setDecimals(3); self.spin_Ctit.setRange(0.0, 1e9); self.spin_Ctit.setValue(5.0)
        self.spin_Vtot = QDoubleSpinBox(self); self.spin_Vtot.setDecimals(1); self.spin_Vtot.setRange(0.0, 1e9); self.spin_Vtot.setValue(DEFAULT_VTOT_UL)

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

        # Tous les champs V et C sont indépendants — aucune liaison automatique.
        # La case compte-goutte sert uniquement de preset pour pré-remplir les valeurs.
        self._updating_fixed = False  # conservé pour _apply_params

        layout.addWidget(box2)

        box3 = QGroupBox("Contraintes")
        form3 = QFormLayout(box3)

        self.spin_margin = QDoubleSpinBox(self); self.spin_margin.setDecimals(1); self.spin_margin.setRange(0.0, 1e9); self.spin_margin.setValue(20.0)
        self.spin_step = QDoubleSpinBox(self); self.spin_step.setDecimals(3); self.spin_step.setRange(0.001, 1e6); self.spin_step.setValue(1.0)

        # Cocher compte-goutte : applique un preset de valeurs, tout reste modifiable ensuite
        def _update_dropper_mode(state):
            if state:
                self.spin_margin.setValue(0.0)
                self.spin_step.setValue(40.0)
                self.spin_VC.setValue(40.0)
                self.spin_VF.setValue(40.0)

        self.chk_dropper.toggled.connect(_update_dropper_mode)

        self.chk_zero = QCheckBox("Inclure un tube à 0 µL de titrant", self); self.chk_zero.setChecked(True)
        self.chk_dup = QCheckBox("Dupliquer le dernier volume pour contrôle", self); self.chk_dup.setChecked(False)

        form3.addRow("Marge (µL)", self.spin_margin)
        form3.addRow("Pas pipette (µL)", self.spin_step)
        form3.addRow(self.chk_zero)
        form3.addRow(self.chk_dup)

        layout.addWidget(box3)

        self.btn_moles = QPushButton("Quantités de matière…", self)
        self.btn_moles.clicked.connect(self._on_edit_moles_clicked)
        layout.addWidget(self.btn_moles)

        self.btn_reset_defaults = QPushButton("Valeurs par défaut", self)
        self.btn_reset_defaults.setToolTip("Remet tous les paramètres aux valeurs d'usine.")
        self.btn_reset_defaults.clicked.connect(self._reset_to_defaults)
        layout.addWidget(self.btn_reset_defaults)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    # Valeurs par défaut de référence (pour le bouton "Réinitialiser")
    _DEFAULTS = {
        "n_tubes": 11, "C0_nM": 1000.0, "V_echant_uL": 1000.0,
        "C_titrant_uM": 5.0, "Vtot_uL": DEFAULT_VTOT_UL,
        "margin_uL": 20.0, "pipette_step_uL": 1.0,
        "include_zero": True, "duplicate_last": False, "compte_goutte": False,
        "_V_C": 30.0, "_V_D": 750.0, "_V_E": 750.0, "_V_F": 60.0,
    }
    _DEFAULTS_FIXED_STOCK = {
        "Solution C": (2.0, "µM"), "Solution D": (0.5, "% en masse"),
        "Solution E": (1.0, "mM"), "Solution F": (1.0, "mM"),
    }

    def _apply_params(self, params: dict, fixed_stock: dict | None = None) -> None:
        """Remplit les widgets avec les paramètres fournis.

        Le signal toggled de chk_dropper est temporairement déconnecté pour éviter
        que cocher la case écrase les volumes sauvegardés.
        """
        # Déconnecter le preset compte-goutte le temps du chargement
        try:
            self.chk_dropper.toggled.disconnect()
        except RuntimeError:
            pass

        self.spin_n.setValue(int(params.get("n_tubes", self.spin_n.value())))
        self.spin_C0.setValue(float(params.get("C0_nM", self.spin_C0.value())))
        self.spin_Vech.setValue(float(params.get("V_echant_uL", self.spin_Vech.value())))
        self.spin_Ctit.setValue(float(params.get("C_titrant_uM", self.spin_Ctit.value())))
        self.spin_Vtot.setValue(float(params.get("Vtot_uL", self.spin_Vtot.value())))
        self.spin_margin.setValue(float(params.get("margin_uL", self.spin_margin.value())))
        self.spin_step.setValue(float(params.get("pipette_step_uL", self.spin_step.value())))
        self.chk_zero.setChecked(bool(params.get("include_zero", self.chk_zero.isChecked())))
        self.chk_dup.setChecked(bool(params.get("duplicate_last", self.chk_dup.isChecked())))
        self.chk_dropper.setChecked(bool(params.get("compte_goutte", self.chk_dropper.isChecked())))
        if "_V_C" in params:
            self.spin_VC.setValue(float(params["_V_C"]))
        if "_V_D" in params:
            self.spin_VD.setValue(float(params["_V_D"]))
        if "_V_E" in params:
            self.spin_VE.setValue(float(params["_V_E"]))
        if "_V_F" in params:
            self.spin_VF.setValue(float(params["_V_F"]))
        if fixed_stock is not None:
            for letter, spin_conc, combo in [
                ("Solution C", self.spin_conc_C, self.combo_unit_C),
                ("Solution D", self.spin_conc_D, self.combo_unit_D),
                ("Solution E", self.spin_conc_E, self.combo_unit_E),
                ("Solution F", self.spin_conc_F, self.combo_unit_F),
            ]:
                if letter in fixed_stock:
                    c_val, u_val = fixed_stock[letter]
                    spin_conc.setValue(float(c_val))
                    combo.setCurrentText(str(u_val))

        # Reconnecter le preset compte-goutte
        def _update_dropper_mode(state):
            if state:
                self.spin_margin.setValue(0.0)
                self.spin_step.setValue(40.0)
                self.spin_VC.setValue(40.0)
                self.spin_VF.setValue(40.0)

        self.chk_dropper.toggled.connect(_update_dropper_mode)

    def _reset_to_defaults(self) -> None:
        """Remet tous les paramètres aux valeurs d'usine."""
        self._apply_params(self._DEFAULTS, self._DEFAULTS_FIXED_STOCK)

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
            "compte_goutte": bool(self.chk_dropper.isChecked()),
        }

    def fixed_solution_stocks(self) -> dict[str, tuple[float, str]]:
        return {
            "Solution C": (float(self.spin_conc_C.value()), str(self.combo_unit_C.currentText()).strip()),
            "Solution D": (float(self.spin_conc_D.value()), str(self.combo_unit_D.currentText()).strip()),
            "Solution E": (float(self.spin_conc_E.value()), str(self.combo_unit_E.currentText()).strip()),
            "Solution F": (float(self.spin_conc_F.value()), str(self.combo_unit_F.currentText()).strip()),
        }

    def _conc_to_n_nmol(self, conc_val: float, unit: str, vol_uL: float) -> float | None:
        if vol_uL <= 0:
            return None
        factor = MOLAR_UNIT_FACTORS.get(str(unit).strip())
        if factor is None:
            return None
        c_m = float(conc_val) * factor
        return c_m * float(vol_uL) * 1e3

    def _n_nmol_to_conc(self, n_nmol: float, unit: str, vol_uL: float) -> float | None:
        if vol_uL <= 0:
            return None
        factor = MOLAR_UNIT_FACTORS.get(str(unit).strip())
        if factor is None:
            return None
        c_m = float(n_nmol) / (float(vol_uL) * 1e3)
        return c_m / factor

    def _on_edit_moles_clicked(self) -> None:
        """Ouvre un dialogue pour fixer les quantités de matière (nmol)."""

        entries: list[dict] = []

        v_echant = float(self.spin_Vech.value())
        n_cu = self._conc_to_n_nmol(float(self.spin_C0.value()), "nM", v_echant)
        entries.append(
            {
                "key": "Echantillon",
                "label": "Echantillon (Cu)",
                "value": n_cu if n_cu is not None else None,
                "editable": n_cu is not None,
                "note": "n = C0 × V échantillon",
            }
        )

        def _add_solution(
            key: str,
            label: str,
            conc_spin: QDoubleSpinBox,
            unit_combo: QComboBox,
            vol_spin: QDoubleSpinBox,
        ) -> None:
            unit = str(unit_combo.currentText()).strip()
            vol = float(vol_spin.value())
            conc = float(conc_spin.value())
            n = self._conc_to_n_nmol(conc, unit, vol)
            note = ""
            editable = n is not None
            if n is None:
                if vol <= 0:
                    note = "Volume nul"
                else:
                    note = "Unité non molaire"
            entries.append(
                {
                    "key": key,
                    "label": label,
                    "value": n if n is not None else None,
                    "editable": editable,
                    "note": note,
                }
            )

        _add_solution("Solution C", "Solution C", self.spin_conc_C, self.combo_unit_C, self.spin_VC)
        _add_solution("Solution D", "Solution D", self.spin_conc_D, self.combo_unit_D, self.spin_VD)
        _add_solution("Solution E", "Solution E", self.spin_conc_E, self.combo_unit_E, self.spin_VE)
        _add_solution("Solution F", "Solution F", self.spin_conc_F, self.combo_unit_F, self.spin_VF)

        dlg = MolesDialog(entries, parent=self)
        if dlg.exec() != QDialog.Accepted:
            return

        values = dlg.values()
        if "Echantillon" in values and v_echant > 0:
            new_c0 = self._n_nmol_to_conc(values["Echantillon"], "nM", v_echant)
            if new_c0 is not None:
                self.spin_C0.setValue(float(new_c0))

        def _apply_solution(
            key: str,
            conc_spin: QDoubleSpinBox,
            unit_combo: QComboBox,
            vol_spin: QDoubleSpinBox,
        ) -> None:
            if key not in values:
                return
            unit = str(unit_combo.currentText()).strip()
            vol = float(vol_spin.value())
            new_conc = self._n_nmol_to_conc(values[key], unit, vol)
            if new_conc is not None:
                conc_spin.setValue(float(new_conc))

        _apply_solution("Solution C", self.spin_conc_C, self.combo_unit_C, self.spin_VC)
        _apply_solution("Solution D", self.spin_conc_D, self.combo_unit_D, self.spin_VD)
        _apply_solution("Solution E", self.spin_conc_E, self.combo_unit_E, self.spin_VE)
        _apply_solution("Solution F", self.spin_conc_F, self.combo_unit_F, self.spin_VF)
    
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
        # Indique que la correspondance spectres↔tubes doit être revue (champs d'en-tête ou volumes modifiés)
        self.map_dirty: bool = False
        self._protocol_states: dict = {}   # état des cases du protocole de paillasse
        # Options d'affichage pour la correspondance spectres ↔ tubes
        self._map_include_brb: bool = True
        self._map_include_control: bool = True
        # Derniers paramètres utilisés dans le dialogue gaussien (mémorisés entre ouvertures)
        self._last_gaussian_params: dict | None = None
        self._last_gaussian_fixed_stock: dict | None = None
        # Cible de contrôle pour la somme des volumes (µL) dans le tableau
        self._sum_target_uL: float = DEFAULT_VTOT_UL

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
        self.btn_save_meta     = QPushButton("Enregistrer les métadonnées…", self)
        self.btn_load_meta     = QPushButton("Charger des métadonnées…", self)
        self.btn_export_proto  = QPushButton("Feuille de protocole…", self)

        self.btn_edit_comp.clicked.connect(self._on_edit_comp_clicked)
        self.btn_edit_map.clicked.connect(self._on_edit_map_clicked)
        self.btn_show_conc.clicked.connect(self._on_show_conc_clicked)
        self.btn_save_meta.clicked.connect(self._on_save_metadata_clicked)
        self.btn_load_meta.clicked.connect(self._on_load_metadata_clicked)
        self.btn_export_proto.clicked.connect(self._on_export_protocol_clicked)

        # Première ligne
        btns.addWidget(self.btn_edit_comp)
        btns.addWidget(self.btn_edit_map)
        btns.addWidget(self.btn_show_conc)

        layout.addLayout(btns)

        # Seconde ligne
        btns2 = QHBoxLayout()
        btns2.addWidget(self.btn_save_meta)
        btns2.addWidget(self.btn_load_meta)
        btns2.addWidget(self.btn_export_proto)
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
        return mm.to_float(x)

    @staticmethod
    def _normalize_tube_label(label: str) -> str:
        return mm.normalize_tube_label(label)
    
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

        def _extract_mapping_rows(existing: pd.DataFrame | None) -> pd.DataFrame | None:
            if existing is None or not isinstance(existing, pd.DataFrame) or existing.empty:
                return None
            if not {"Nom du spectre", "Tube"}.issubset(existing.columns):
                return None
            hdr_idx = None
            for idx, row in existing.iterrows():
                left = str(row.get("Nom du spectre", "")).strip()
                right = str(row.get("Tube", "")).strip()
                if left == "Nom du spectre" and right == "Tube":
                    hdr_idx = idx
                    break
            if hdr_idx is not None:
                return existing.loc[hdr_idx + 1 :].reset_index(drop=True)
            return existing.reset_index(drop=True)

        def _is_brb_label(label: str) -> bool:
            return "brb" in self._norm_text_key(label)

        def _build_initial_df(
            include_brb: bool,
            include_control: bool,
            existing_df: pd.DataFrame | None,
        ) -> tuple[pd.DataFrame, pd.DataFrame]:
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
            tube_labels_default = self._default_tube_labels_for_mapping(
                n_tubes,
                include_brb_override=include_brb,
                include_control_override=include_control,
            )

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
            mapping_rows = _extract_mapping_rows(existing_df)

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
                    if (not include_brb) and _is_brb_label(tube_raw):
                        continue
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
                    # Forcer l'affichage du dernier tube comme "Contrôle" si demandé,
                    # et inversement forcer "Tube N" si le contrôle n'est pas demandé.
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
            return initial_df, reset_df

        # Valeurs par défaut des options BRB / Contrôle
        include_brb = self._map_include_brb
        include_control = self._map_include_control
        current_mapping = _extract_mapping_rows(self.df_map)
        if current_mapping is not None and not current_mapping.empty:
            tubes = current_mapping["Tube"].astype(str).str.strip()
            include_brb = any(_is_brb_label(t) for t in tubes)
            include_control = any(self._is_control_label(t) for t in tubes)

        initial_df, reset_df = _build_initial_df(include_brb, include_control, self.df_map)

        # 5) Ouverture du dialogue d'édition
        dlg = TableEditorDialog(
            title="Correspondance spectres ↔ tubes",
            columns=default_cols,
            initial_df=initial_df,
            parent=self,
            reset_df=reset_df,
            reset_button_text="Reset (valeurs par défaut)",
        )
        # Options : inclure Contrôle BRB / Contrôle (dernier tube)
        opts_layout = QHBoxLayout()
        chk_brb = QCheckBox('Inclure "Contrôle BRB"', dlg)
        chk_ctrl = QCheckBox('Inclure "Contrôle" (dernier tube)', dlg)
        chk_brb.setChecked(bool(include_brb))
        chk_ctrl.setChecked(bool(include_control))
        opts_layout.addWidget(chk_brb)
        opts_layout.addWidget(chk_ctrl)
        opts_layout.addStretch(1)
        try:
            dlg.layout().insertLayout(1, opts_layout)
        except Exception:
            pass

        def _rebuild_mapping_table() -> None:
            inc_brb = chk_brb.isChecked()
            inc_ctrl = chk_ctrl.isChecked()
            try:
                current_df = dlg.to_dataframe()
            except Exception:
                current_df = self.df_map
            new_df, new_reset = _build_initial_df(inc_brb, inc_ctrl, current_df)
            dlg._reset_df = new_reset
            try:
                dlg._load_from_dataframe(new_df)
            except Exception:
                pass

        chk_brb.toggled.connect(_rebuild_mapping_table)
        chk_ctrl.toggled.connect(_rebuild_mapping_table)
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

        # Mémoriser les options BRB / Contrôle
        self._map_include_brb = bool(chk_brb.isChecked())
        self._map_include_control = bool(chk_ctrl.isChecked())

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
        # Restaurer les derniers paramètres utilisés si disponibles
        if self._last_gaussian_params is not None:
            dlg._apply_params(self._last_gaussian_params, self._last_gaussian_fixed_stock)
        if dlg.exec() != QDialog.Accepted:
            return

        params = dlg.params()
        fixed_stock = dlg.fixed_solution_stocks()
        # Mémoriser les paramètres pour la prochaine ouverture
        self._last_gaussian_params = dict(params)
        self._last_gaussian_params["_V_C"] = dlg.spin_VC.value()
        self._last_gaussian_params["_V_D"] = dlg.spin_VD.value()
        self._last_gaussian_params["_V_E"] = dlg.spin_VE.value()
        self._last_gaussian_params["_V_F"] = dlg.spin_VF.value()
        self._last_gaussian_fixed_stock = dict(fixed_stock)
        self._sum_target_uL = float(params.get("Vtot_uL", DEFAULT_VTOT_UL))
        compte_goutte = bool(params.pop("compte_goutte", False))

        try:
            df_gen, meta = mm.sers_gaussian_volumes(**params)
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
                {"Réactif": "Contrôle", "Concentration": float(params["C0_nM"]), "Unité": "nM"},
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
        idx_control = ensure_row("Contrôle")
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
            "Contrôle": 0.0,
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

        # Le dernier tube est le contrôle : on y met l'échantillon dans la ligne "Contrôle"
        # et on enlève l'échantillon de la ligne "Echantillon"
        last_col = tube_cols[-1]

        if "Contrôle" in df["Réactif"].values:
            ridx_ctrl = int(df.index[df["Réactif"] == "Contrôle"][0])
            df.at[ridx_ctrl, last_col] = float(params["V_echant_uL"])

        if "Echantillon" in df["Réactif"].values:
            ridx_samp = int(df.index[df["Réactif"] == "Echantillon"][0])
            df.at[ridx_samp, last_col] = 0.0

        # ── Mode compte-goutte : convertir les volumes B, C, F en nombre de gouttes ──
        # La colonne "Pas (µL)" mémorise le pas pour que compute_concentration_table
        # puisse retrouver les µL réels (vol_µL = n_gouttes × pas).
        pas = params.get("pipette_step_uL", 1.0)
        if "Pas (µL)" not in df.columns:
            df["Pas (µL)"] = 0.0
        if compte_goutte and pas > 0:
            for reactif_name in ("Solution B", "Solution C", "Solution F"):
                mask = df["Réactif"].astype(str).str.strip().str.lower() == reactif_name.strip().lower()
                if not mask.any():
                    continue
                ridx = int(df.index[mask][0])
                df.at[ridx, "Pas (µL)"] = pas
                for c in tube_cols:
                    vol_uL = df.at[ridx, c]
                    try:
                        n = float(vol_uL) / pas
                        df.at[ridx, c] = round(n)
                    except (ValueError, TypeError, ZeroDivisionError):
                        pass

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
        return mm.get_tube_columns(df)

    def _infer_n_tubes_for_mapping(self) -> int:
        return mm.infer_n_tubes_for_mapping(self.df_comp)

    def _sync_df_map_with_df_comp(
        self,
        duplicate_last_override: bool | None = None,
        include_brb_override: bool | None = None,
        include_control_override: bool | None = None,
        *,
        mark_dirty: bool = True,
    ) -> None:
        """Synchronise automatiquement la correspondance spectres↔tubes avec le tableau des volumes.

        Objectifs :
        - le nombre de lignes de correspondance suit le nombre de colonnes 'Tube N'
        - le dernier tube est 'Contrôle' uniquement si un duplicat est détecté
        - pas de doublons de noms de spectres (renumérotation si nécessaire)

        Cette méthode n'ouvre pas de dialogue : elle met à jour self.df_map en place.
        Par défaut, elle marque la correspondance comme à revoir (pas de validation utilisateur).
        """
        default_cols = ["Nom du spectre", "Tube"]
        include_brb = self._map_include_brb if include_brb_override is None else bool(include_brb_override)
        include_control = (
            self._map_include_control if include_control_override is None else bool(include_control_override)
        )

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
            include_brb_override=include_brb,
            include_control_override=include_control,
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
                if (not include_brb) and ("brb" in self._norm_text_key(tube_raw)):
                    continue
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
        self.map_dirty = bool(mark_dirty)

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
        return mm.norm_text_key(s)

    def _is_control_label(self, label: str) -> bool:
        return mm.is_control_label(label)

    def _tube_merge_key_for_mapping(self, label: str, n_tubes: int | None = None) -> str:
        return mm.tube_merge_key_for_mapping(label, n_tubes, self.df_comp)

    def _default_tube_labels_for_mapping(
        self,
        n_tubes: int | None = None,
        *,
        duplicate_last_override: bool | None = None,
        include_brb_override: bool | None = None,
        include_control_override: bool | None = None,
    ) -> list[str]:
        """Liste de tubes par défaut pour la correspondance.

        Convention :
        - 1er spectre : "Contrôle BRB" (optionnel)
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
        include_brb = self._map_include_brb if include_brb_override is None else bool(include_brb_override)
        include_control = (
            self._map_include_control if include_control_override is None else bool(include_control_override)
        )
        if n <= 1:
            tubes = ["Contrôle" if (duplicate_last and include_control) else "Tube 1"]
        elif duplicate_last and include_control:
            tubes = [f"Tube {i}" for i in range(1, n)] + ["Contrôle"]
        else:
            tubes = [f"Tube {i}" for i in range(1, n + 1)]
        return (["Contrôle BRB"] if include_brb else []) + tubes
    # ------------------------------------------------------------------
    # Calcul du tableau des concentrations à partir des volumes
    # ------------------------------------------------------------------
    def _compute_concentration_table(self) -> pd.DataFrame:
        return mm.compute_concentration_table(self.df_comp)
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
        return mm.build_merged_metadata(self.df_comp, self.df_map)
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

        def _value_for(*labels: str) -> str:
            """Cherche la valeur pour n'importe laquelle des variantes de label."""
            for label in labels:
                mask = col_ns == label
                if mask.any():
                    return col_tube[mask].iloc[0].strip()
            return ""

        # Nom de la manip
        name_val = _value_for("Nom de la manip :", "Nom de la manip")
        if name_val and hasattr(self, "edit_manip"):
            self.edit_manip.setText(name_val)

        # Date
        date_val = _value_for("Date :", "Date")
        if date_val and hasattr(self, "edit_date"):
            qd = QDate.fromString(date_val, "dd/MM/yyyy")
            if qd.isValid():
                self.edit_date.setDate(qd)

        # Lieu
        loc_val = _value_for("Lieu :", "Lieu")
        if loc_val and hasattr(self, "edit_location"):
            self.edit_location.setText(loc_val)

        # Coordinateur
        coord_val = _value_for("Coordinateur :", "Coordinateur")
        if coord_val and hasattr(self, "edit_coordinator"):
            self.edit_coordinator.setText(coord_val)

        # Opérateur
        oper_val = _value_for("Opérateur :", "Opérateur")
        if oper_val and hasattr(self, "edit_operator"):
            self.edit_operator.setText(oper_val)

    # ------------------------------------------------------------------
    # Export feuille de protocole
    # ------------------------------------------------------------------
    def _on_export_protocol_clicked(self) -> None:
        """Ouvre le dialogue de suivi interactif du protocole de paillasse."""
        if self.df_comp is None or self.df_comp.empty:
            QMessageBox.warning(
                self,
                "Tableau des volumes manquant",
                "Veuillez d'abord définir le tableau des volumes avant d'ouvrir le protocole.",
            )
            return

        meta = {
            "nom_manip":    self.edit_manip.text().strip()                   if hasattr(self, "edit_manip")       else "",
            "date":         self.edit_date.date().toString("dd/MM/yyyy")     if hasattr(self, "edit_date")        else "",
            "lieu":         self.edit_location.text().strip()                if hasattr(self, "edit_location")    else "",
            "coordinateur": self.edit_coordinator.text().strip()             if hasattr(self, "edit_coordinator") else "",
            "operateur":    self.edit_operator.text().strip()                if hasattr(self, "edit_operator")    else "",
        }

        from protocol_dialog import ProtocolDialog
        dlg = ProtocolDialog(self.df_comp, meta,
                             initial_states=self._protocol_states,
                             df_map=self.df_map, parent=self)
        dlg.exec()
        # Mémoriser l'état des cases pour la prochaine ouverture / sauvegarde
        self._protocol_states = dlg.get_states()

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
                # État du protocole de paillasse
                if self._protocol_states:
                    rows = []
                    for key, checked in self._protocol_states.items():
                        if key[0] == "verif":
                            rows.append({"type": "verif", "r_idx": key[1], "t_idx": -1, "checked": int(checked)})
                        else:
                            rows.append({"type": key[0], "r_idx": key[1], "t_idx": key[2], "checked": int(checked)})
                    pd.DataFrame(rows).to_excel(writer, sheet_name="EtatProtocole", index=False)
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
            # État du protocole de paillasse
            self._protocol_states = {}
            if "EtatProtocole" in xls.sheet_names:
                df_proto = pd.read_excel(xls, sheet_name="EtatProtocole")
                for _, r in df_proto.iterrows():
                    t   = str(r.get("type", ""))
                    ri  = int(r.get("r_idx", 0))
                    ti  = int(r.get("t_idx", -1))
                    chk = bool(r.get("checked", 0))
                    if t == "verif":
                        self._protocol_states[("verif", ri)] = chk
                    elif t in ("coord", "oper"):
                        self._protocol_states[(t, ri, ti)] = chk
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
            sum_row=True,
            sum_target=self._sum_target_uL,
        )
        # ---------------------------------------------------------
        # Synchronisation Solution A + Solution B = constante
        # ---------------------------------------------------------
        # On impose :
        #     Solution A + Solution B = constante (par tube)
        # La constante est définie à l'ouverture du tableau.
        # Si l'utilisateur modifie A → B s'ajuste, et inversement.
        table = dlg.table

        def _norm_name(txt: str) -> str:
            return (
                str(txt)
                .strip()
                .lower()
                .replace("é", "e")
                .replace("è", "e")
                .replace("ê", "e")
            )

        def _reactif_col_index() -> int:
            for j in range(table.columnCount()):
                header = table.horizontalHeaderItem(j)
                if header is None:
                    continue
                name = _norm_name(header.text())
                if name in ("reactif", "reactifs"):
                    return j
            return 0

        def _is_tube_col(col_idx: int) -> bool:
            header = table.horizontalHeaderItem(col_idx)
            if header is None:
                return False
            return header.text().strip().lower().startswith("tube")

        def _find_rows_ab() -> tuple[int | None, int | None, int]:
            col_reactif = _reactif_col_index()
            row_a = None
            row_b = None
            sum_row = getattr(dlg, "_sum_row_idx", None)
            for r in range(table.rowCount()):
                if sum_row is not None and r == sum_row:
                    continue
                item = table.item(r, col_reactif)
                if item is None:
                    continue
                name = _norm_name(item.text())
                if name == "solution a":
                    row_a = r
                elif name == "solution b":
                    row_b = r
            return row_a, row_b, col_reactif

        def _get_pas_B(row_b: int) -> float:
            """Retourne le pas (µL/goutte) de Solution B (0 si pas de mode goutte)."""
            for j in range(table.columnCount()):
                hdr = table.horizontalHeaderItem(j)
                if hdr is not None and hdr.text().strip() == "Pas (µL)":
                    pas_item = table.item(row_b, j)
                    return dlg._parse_float(pas_item.text() if pas_item is not None else "") or 0.0
            return 0.0

        # Somme cible A+B (en µL) pour chaque tube — initialisée à l'ouverture
        target_sum: dict[int, float] = {}
        row_A, row_B, _ = _find_rows_ab()
        if row_A is not None and row_B is not None:
            pas_B_init = _get_pas_B(row_B)
            for c in range(table.columnCount()):
                if not _is_tube_col(c):
                    continue
                itemA = table.item(row_A, c)
                itemB = table.item(row_B, c)
                vA = dlg._parse_float(itemA.text() if itemA is not None else "") or 0.0
                vB_raw = dlg._parse_float(itemB.text() if itemB is not None else "") or 0.0
                # En mode goutte, vB_raw = nombre de gouttes → convertir en µL
                vB_uL = vB_raw * pas_B_init if pas_B_init > 0 else vB_raw
                target_sum[c] = vA + vB_uL

        def sync_A_B(item: QTableWidgetItem) -> None:
            if item is None:
                return
            # Saisie manuelle activée → ne rien synchroniser
            if dlg.chk_manual_mode.isChecked():
                return
            r = item.row()
            c = item.column()
            if not _is_tube_col(c):
                return
            sum_row = getattr(dlg, "_sum_row_idx", None)
            if sum_row is not None and r == sum_row:
                return

            row_A, row_B, _ = _find_rows_ab()
            if row_A is None or row_B is None:
                return
            if r not in (row_A, row_B):
                return

            val = dlg._parse_float(item.text())
            if val is None:
                return

            pas_B = _get_pas_B(row_B)

            if c not in target_sum:
                itemA = table.item(row_A, c)
                itemB = table.item(row_B, c)
                vA = dlg._parse_float(itemA.text() if itemA is not None else "") or 0.0
                vB_raw = dlg._parse_float(itemB.text() if itemB is not None else "") or 0.0
                vB_uL = vB_raw * pas_B if pas_B > 0 else vB_raw
                target_sum[c] = vA + vB_uL

            total = target_sum[c]  # toujours en µL

            if r == row_B:
                # L'utilisateur a modifié Solution B
                # val = nombre de gouttes (si compte-goutte) ou µL sinon
                val_uL = val * pas_B if pas_B > 0 else val
                new_val_A = total - val_uL   # Solution A en µL
                other_row, new_val = row_A, new_val_A
            else:
                # L'utilisateur a modifié Solution A (toujours en µL)
                remaining_uL = total - val
                if pas_B > 0:
                    # Convertir µL restants en nombre de gouttes (arrondi)
                    new_val_B = round(remaining_uL / pas_B)
                else:
                    new_val_B = remaining_uL
                other_row, new_val = row_B, new_val_B

            table.blockSignals(True)
            try:
                other_item = table.item(other_row, c)
                if other_item is None:
                    other_item = QTableWidgetItem("")
                    table.setItem(other_row, c, other_item)
                other_item.setText(dlg._format_number(new_val))
            finally:
                table.blockSignals(False)

            # Met à jour la ligne "Total" après ajustement automatique
            try:
                dlg._refresh_sum_row()
            except Exception:
                pass

        table.itemChanged.connect(sync_A_B)
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
