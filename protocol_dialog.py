"""Dialogue de suivi de protocole de paillasse.

Affiche le tableau des volumes avec, pour chaque tube, des cases à cocher
cliquables (Coordinateur + Opérateur).

Mode guidé : seule l'étape courante est éclairée ; tout le reste est ombré
et désactivé.  Séquence par ligne (réactif) :
  Vérif. solution → Tube 1 Coord → Tube 1 Opér → Tube 2 Coord → … → Tube N Opér
Puis passage au réactif suivant.

Un bouton "Exporter Excel" génère la feuille stylée avec l'état courant.
"""

from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush, QFont
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QCheckBox, QWidget,
    QHeaderView, QLabel, QFileDialog, QMessageBox,
    QSizePolicy,
)

import metadata_model as mm
import protocol_export

# ── Couleurs actives (identique à protocol_export) ───────────────────────────
_C_TUBE     = QColor("#2E75B6")
_C_COORD    = QColor("#5B9BD5")
_C_OPER     = QColor("#9DC3E6")
_C_REACT    = QColor("#ED7D31")
_C_VERIF    = QColor("#70AD47")
_C_VERIF_L  = QColor("#E2EFDA")
_C_META     = QColor("#D6DCE4")
_C_ROW_A    = QColor("#DEEAF1")
_C_ROW_B    = QColor("#FFFFFF")
_C_GRP_HDR  = QColor("#1F4E79")
_C_VOL      = QColor("#BDD7EE")
_C_COORD_A  = QColor("#DAEEF3")
_C_COORD_B  = QColor("#EBF5FB")
_C_OPER_A   = QColor("#E8F4F8")
_C_OPER_B   = QColor("#F5FBFD")
_C_WHITE    = QColor("#FFFFFF")

# ── Couleurs ombrées ──────────────────────────────────────────────────────────
_C_DIM_BG   = QColor("#2C2C2C")   # fond ombré (thème sombre)
_C_DIM_FG   = QColor("#555555")   # texte ombré
_SS_DIM     = "background: #2C2C2C; color: #555555;"
_SS_ACTIVE  = "background: transparent;"

_FIXED_COLS = 4   # Réactif | Conc. | Unité | Vérif.
_SUBCOLS    = 3   # Vol. | Coord. | Opér.  par tube

# Indices relatifs des sous-colonnes tube dans QTableWidget (0-based)
_SUB_VOL   = 0
_SUB_COORD = 1
_SUB_OPER  = 2


# ── Helpers ───────────────────────────────────────────────────────────────────

def _centered_checkbox(bg: str = "transparent") -> tuple[QWidget, QCheckBox]:
    w   = QWidget()
    cb  = QCheckBox()
    lay = QHBoxLayout(w)
    lay.addWidget(cb)
    lay.setAlignment(Qt.AlignCenter)
    lay.setContentsMargins(0, 0, 0, 0)
    w.setStyleSheet(f"background: {bg};")
    return w, cb


def _item(text: str = "", *, bg: QColor | None = None,
          fg: QColor | None = None, bold: bool = False,
          center: bool = True) -> QTableWidgetItem:
    it = QTableWidgetItem(text)
    if bg: it.setBackground(QBrush(bg))
    if fg: it.setForeground(QBrush(fg))
    if bold:
        f = it.font(); f.setBold(True); it.setFont(f)
    it.setTextAlignment(
        (Qt.AlignCenter if center else Qt.AlignLeft) | Qt.AlignVCenter
    )
    it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
    return it


# ── Calcul des couleurs normales ──────────────────────────────────────────────

def _normal_item_bg(r_idx: int, col: int, n_fixed: int, n_subcols: int) -> QColor:
    """Retourne la couleur de fond normale d'une cellule (hors ombrée)."""
    if col == 0:   return _C_REACT
    if col == 1 or col == 2: return _C_META
    if col == 3:   return _C_VERIF_L  # verif
    rel = (col - n_fixed) % n_subcols
    if rel == _SUB_VOL:   return _C_VOL
    if rel == _SUB_COORD: return _C_COORD_A if r_idx % 2 == 0 else _C_COORD_B
    return _C_OPER_A if r_idx % 2 == 0 else _C_OPER_B


# ── Dialogue ──────────────────────────────────────────────────────────────────

class ProtocolDialog(QDialog):
    """Dialogue de suivi interactif du protocole de paillasse (mode guidé)."""

    def __init__(self, df_comp: pd.DataFrame, meta: dict,
                 initial_states: dict | None = None,
                 df_map=None, parent=None):
        super().__init__(parent)
        self._df            = df_comp.copy()
        self._df_map        = df_map
        self._meta          = meta
        self._initial_states = initial_states or {}
        self._tubes         = mm.get_tube_columns(df_comp)
        self._n_tubes       = len(self._tubes)
        self._n_rows        = len(df_comp)

        # Registres des widgets
        self._checks:  dict[tuple, QCheckBox] = {}   # clé → QCheckBox
        self._c_widgets: dict[tuple[int,int], QWidget] = {}  # (tr,col) → QWidget conteneur

        self.setWindowTitle("Suivi de protocole — Paillasse")
        parent_width = parent.width() if parent is not None else 0
        default_width = parent_width if parent_width > 0 else max(900, 300 + self._n_tubes * 120)
        self.resize(default_width, 560)
        self.setSizeGripEnabled(True)

        root = QVBoxLayout(self)
        root.setSpacing(6)

        self._build_info_bar(root, meta)
        self._build_legend(root)

        # Table
        n_cols = _FIXED_COLS + self._n_tubes * _SUBCOLS
        self._table = QTableWidget(1 + self._n_rows, n_cols, self)
        self._table.verticalHeader().setVisible(False)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionMode(QTableWidget.NoSelection)
        self._table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._table.setShowGrid(True)
        self._build_headers()
        self._build_data()
        self._apply_initial_states()
        root.addWidget(self._table)

        self._build_buttons(root)

        # Premier affichage
        self._refresh_highlights()

    # ── Barres d'info et légende ──────────────────────────────────────────────

    def _build_info_bar(self, root, meta):
        row = QHBoxLayout()
        for label, val in [
            ("Manip :",        meta.get("nom_manip", "—")),
            ("Date :",         meta.get("date", "—")),
            ("Lieu :",         meta.get("lieu", "—")),
            ("Coordinateur :", meta.get("coordinateur", "—")),
            ("Opérateur :",    meta.get("operateur", "—")),
        ]:
            lbl = QLabel(f"<b>{label}</b> {val}")
            lbl.setStyleSheet("font-size: 11px; color: #d4d4d4;")
            row.addWidget(lbl)
            row.addSpacing(12)
        row.addStretch()
        root.addLayout(row)

    def _build_legend(self, root):
        row = QHBoxLayout()
        self._lbl_step = QLabel("")
        self._lbl_step.setStyleSheet(
            "font-size: 12px; font-weight: bold; color: #46ba48; padding: 2px 6px;"
        )
        row.addWidget(self._lbl_step)
        row.addStretch()
        for color_hex, label in [
            ("#2E75B6", "Volume"),
            ("#5B9BD5", "Coordinateur"),
            ("#9DC3E6", "Opérateur"),
            ("#70AD47", "Vérif. solution"),
        ]:
            dot = QLabel("●")
            dot.setStyleSheet(f"color: {color_hex}; font-size: 14px;")
            txt = QLabel(label)
            txt.setStyleSheet("font-size: 11px; color: #aaa; margin-right: 10px;")
            row.addWidget(dot); row.addWidget(txt)
        root.addLayout(row)

    # ── Construction de la table ──────────────────────────────────────────────

    def _build_headers(self):
        tbl = self._table
        labels = ["Réactif", "Conc.", "Unité", "Vérif."]
        for _ in self._tubes:
            labels += ["Vol. (µL)", "Coord.", "Opér."]
        tbl.setHorizontalHeaderLabels(labels)
        tbl.horizontalHeader().setDefaultSectionSize(60)
        tbl.setColumnWidth(0, 140)
        tbl.setColumnWidth(1, 70)
        tbl.setColumnWidth(2, 60)
        tbl.setColumnWidth(3, 55)
        for t in range(self._n_tubes):
            base = _FIXED_COLS + t * _SUBCOLS
            tbl.setColumnWidth(base,     70)
            tbl.setColumnWidth(base + 1, 55)
            tbl.setColumnWidth(base + 2, 55)

        # Ligne 0 : en-têtes de groupe
        tbl.setRowHeight(0, 28)
        for col, bg, label in [
            (0, _C_REACT, "Réactif"),
            (1, _C_META,  ""),
            (2, _C_META,  ""),
            (3, _C_VERIF, "Vérif.\nsolution"),
        ]:
            tbl.setItem(0, col, _item(label, bg=bg, fg=_C_WHITE, bold=True))
        for t_idx, tube_name in enumerate(self._tubes):
            base = _FIXED_COLS + t_idx * _SUBCOLS
            tbl.setSpan(0, base, 1, _SUBCOLS)
            tbl.setItem(0, base, _item(tube_name, bg=_C_GRP_HDR, fg=_C_WHITE, bold=True))

    def _build_data(self):
        tbl = self._table
        for r_idx, (_, row) in enumerate(self._df.iterrows()):
            tr = r_idx + 1
            tbl.setRowHeight(tr, 26)

            tbl.setItem(tr, 0, _item(str(row.get("Réactif", "")),
                                     bg=_C_ROW_A if r_idx % 2 == 0 else _C_ROW_B,
                                     center=False))
            raw = row.get("Concentration", "")
            try:    conc_str = str(float(raw))
            except: conc_str = str(raw) if raw not in (None, "") else ""
            rbg = _C_ROW_A if r_idx % 2 == 0 else _C_ROW_B
            tbl.setItem(tr, 1, _item(conc_str, bg=rbg))
            tbl.setItem(tr, 2, _item(str(row.get("Unité", "")), bg=rbg))

            # Vérification solution
            w, cb = _centered_checkbox(_C_VERIF_L.name())
            self._checks[("verif", r_idx)] = cb
            self._c_widgets[(tr, 3)] = w
            cb.stateChanged.connect(self._refresh_highlights)
            tbl.setCellWidget(tr, 3, w)

            # Mode compte-goutte : pas > 0 → valeurs stockées en n_gouttes
            try:
                pas = float(row.get("Pas (µL)", 0) or 0)
            except (ValueError, TypeError):
                pas = 0.0

            # Tubes
            for t_idx, tube_name in enumerate(self._tubes):
                base = _FIXED_COLS + t_idx * _SUBCOLS
                raw_vol = row.get(tube_name, "")
                try:
                    v = float(raw_vol)
                    if pas > 0:
                        # Afficher le nombre de gouttes
                        n = int(round(v))
                        vol_str = f"{n} gouttes"
                    else:
                        vol_str = str(int(v)) if v == int(v) else str(v)
                except (ValueError, TypeError):
                    vol_str = str(raw_vol) if raw_vol not in (None, "") else ""

                tbl.setItem(tr, base, _item(vol_str, bg=_C_VOL))

                bg_c = (_C_COORD_A if r_idx % 2 == 0 else _C_COORD_B).name()
                w_c, cb_c = _centered_checkbox(bg_c)
                self._checks[("coord", r_idx, t_idx)] = cb_c
                self._c_widgets[(tr, base + 1)] = w_c
                cb_c.stateChanged.connect(self._refresh_highlights)
                tbl.setCellWidget(tr, base + 1, w_c)

                bg_o = (_C_OPER_A if r_idx % 2 == 0 else _C_OPER_B).name()
                w_o, cb_o = _centered_checkbox(bg_o)
                self._checks[("oper", r_idx, t_idx)] = cb_o
                self._c_widgets[(tr, base + 2)] = w_o
                cb_o.stateChanged.connect(self._refresh_highlights)
                tbl.setCellWidget(tr, base + 2, w_o)

    def _apply_initial_states(self):
        """Pré-coche les cases selon les états chargés depuis le fichier."""
        for key, checked in self._initial_states.items():
            cb = self._checks.get(key)
            if cb is not None and checked:
                cb.blockSignals(True)
                cb.setEnabled(True)
                cb.setChecked(True)
                cb.blockSignals(False)

    def _build_buttons(self, root):
        row = QHBoxLayout()
        btn_coord  = QPushButton("Tout cocher — Coord.")
        btn_oper   = QPushButton("Tout cocher — Opér.")
        btn_reset  = QPushButton("Tout décocher")
        self._btn_export = QPushButton("Exporter Excel…")
        self._btn_close  = QPushButton("Fermer")

        btn_coord.clicked.connect(lambda: self._check_all("coord", True))
        btn_oper.clicked.connect(lambda:  self._check_all("oper",  True))
        btn_reset.clicked.connect(lambda: self._check_all(None,    False))
        self._btn_export.clicked.connect(self._on_export)
        self._btn_close.clicked.connect(self.accept)

        for btn in (btn_coord, btn_oper, btn_reset):
            row.addWidget(btn)
        row.addStretch()
        row.addWidget(self._btn_export)
        row.addWidget(self._btn_close)
        root.addLayout(row)

    def _check_all(self, role: str | None, state: bool):
        """Coche ou décoche toutes les cases d'un rôle (ou toutes si role=None)."""
        # Bloquer les signaux pour éviter N rafraîchissements intermédiaires
        for key, cb in self._checks.items():
            if role is None or key[0] == role:
                cb.blockSignals(True)
                cb.setEnabled(True)   # temporairement pour pouvoir cocher
                cb.setChecked(state)
                cb.blockSignals(False)
        self._refresh_highlights()

    # ── Logique de guidage ────────────────────────────────────────────────────

    def _find_active(self) -> tuple | None:
        """Retourne la prochaine étape à compléter.

        Retourne :
          ("verif", r_idx)           — vérification solution non cochée
          ("tube",  r_idx, t_idx)    — coord ou opér non coché pour ce tube
          None                       — tout est terminé
        """
        for r_idx in range(self._n_rows):
            if not self._checks[("verif", r_idx)].isChecked():
                return ("verif", r_idx)
            for t_idx in range(self._n_tubes):
                cb_c = self._checks[("coord", r_idx, t_idx)]
                cb_o = self._checks[("oper",  r_idx, t_idx)]
                if not cb_c.isChecked() or not cb_o.isChecked():
                    return ("tube", r_idx, t_idx)
        return None

    def _active_cols(self, active: tuple) -> set[int]:
        """Colonnes éclairées (0-based) pour l'étape active."""
        if active[0] == "verif":
            return {3}   # colonne Vérif.
        t_idx = active[2]
        base = _FIXED_COLS + t_idx * _SUBCOLS
        # Volume inclus pour montrer la quantité à pipeter
        return {base + _SUB_VOL, base + _SUB_COORD, base + _SUB_OPER}

    def _refresh_highlights(self):
        """Met à jour l'ombrageage de toute la table selon l'étape courante."""
        active   = self._find_active()
        n_cols   = self._table.columnCount()
        tbl      = self._table

        if active is None:
            # Tout terminé : tout éclairé
            self._lbl_step.setText("✓ Protocole complété !")
            for tr in range(1, 1 + self._n_rows):
                r_idx = tr - 1
                for col in range(n_cols):
                    self._set_cell(tr, col, r_idx, dim=False, enabled=True)
            return

        active_r_idx = active[1]
        active_cols  = self._active_cols(active)

        # Texte de guidage
        if active[0] == "verif":
            react = self._df.iloc[active_r_idx].get("Réactif", f"Ligne {active_r_idx+1}")
            self._lbl_step.setText(f"→ Vérifier la solution : {react}")
        else:
            react = self._df.iloc[active_r_idx].get("Réactif", f"Ligne {active_r_idx+1}")
            tube  = self._tubes[active[2]]
            self._lbl_step.setText(f"→ {react}  ·  {tube}")

        for tr in range(1, 1 + self._n_rows):
            r_idx     = tr - 1
            same_row  = (r_idx == active_r_idx)

            for col in range(n_cols):
                same_col = (col in active_cols)
                # Éclairé seulement si même ligne ET même colonne active
                lit     = same_row and same_col
                enabled = lit
                dim     = not lit
                self._set_cell(tr, col, r_idx, dim=dim, enabled=enabled)

    def _set_cell(self, tr: int, col: int, r_idx: int,
                  dim: bool, enabled: bool):
        """Applique ou retire l'ombrageage sur une cellule."""
        tbl = self._table

        # ── Cellule item ──────────────────────────────────────────────────────
        it = tbl.item(tr, col)
        if it is not None:
            if dim:
                it.setBackground(QBrush(_C_DIM_BG))
                it.setForeground(QBrush(_C_DIM_FG))
            else:
                normal_bg = _normal_item_bg(r_idx, col, _FIXED_COLS, _SUBCOLS)
                it.setBackground(QBrush(normal_bg))
                it.setForeground(QBrush(QColor("#000000")))

        # ── Cellule widget (checkbox) ─────────────────────────────────────────
        w = self._c_widgets.get((tr, col))
        if w is not None:
            if dim:
                w.setStyleSheet(_SS_DIM)
            else:
                normal_bg = _normal_item_bg(r_idx, col, _FIXED_COLS, _SUBCOLS)
                w.setStyleSheet(f"background: {normal_bg.name()};")
            # Activer / désactiver la case
            cb = w.findChild(QCheckBox)
            if cb is not None:
                cb.setEnabled(enabled or cb.isChecked())

    # ── Export ────────────────────────────────────────────────────────────────

    def get_states(self) -> dict:
        return {k: cb.isChecked() for k, cb in self._checks.items()}

    def _on_export(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter la feuille de protocole", "",
            "Fichier Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            protocol_export.export_protocol_excel(
                self._df, self._meta, path,
                states=self.get_states(), df_map=self._df_map
            )
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'export",
                                 f"Impossible de générer la feuille :\n{e}")
            return
        QMessageBox.information(self, "Export réussi",
                                f"Feuille enregistrée dans :\n{path}")
