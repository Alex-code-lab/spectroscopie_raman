"""Génération de la feuille de protocole Excel (feuille de paillasse).

Pour chaque réactif (Solution A–F, Échantillon, Eau de contrôle) et chaque tube,
le tableau comporte :
  - la valeur de volume (µL)
  - une case "Coordinateur" (à cocher lors de l'annonce de la quantité)
  - une case "Opérateur"    (à cocher lors de l'ajout effectif)

La mise en forme suit le modèle CitizenSers :
  en-têtes pivotés à 90°, couleurs par rôle (volume / coordinateur / opérateur),
  colonne "Réactifs" verticale orange, cases à cocher ☐.
"""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import metadata_model as mm

# ── Palette ───────────────────────────────────────────────────────────────────
_TUBE_HDR    = PatternFill("solid", fgColor="2E75B6")  # bleu foncé  — volume
_COORD_HDR   = PatternFill("solid", fgColor="5B9BD5")  # bleu moyen  — coordinateur
_OPER_HDR    = PatternFill("solid", fgColor="9DC3E6")  # bleu clair  — opérateur
_REACT_SIDE  = PatternFill("solid", fgColor="ED7D31")  # orange      — label Réactifs
_VERIF_HDR   = PatternFill("solid", fgColor="70AD47")  # vert        — vérification
_VERIF_CELL  = PatternFill("solid", fgColor="E2EFDA")  # vert clair
_META_HDR    = PatternFill("solid", fgColor="D6DCE4")  # gris        — nom/conc/unité
_ROW_A       = PatternFill("solid", fgColor="DEEAF1")  # bleu très clair (lignes paires)
_ROW_B       = PatternFill("solid", fgColor="FFFFFF")  # blanc       (lignes impaires)

# ── Polices ───────────────────────────────────────────────────────────────────
_WHITE_BOLD  = Font(color="FFFFFF", bold=True, size=9)
_HDR_FONT    = Font(color="FFFFFF", bold=True, size=8)
_DARK        = Font(color="000000", size=9)
_BOLD        = Font(bold=True, size=9)
_CHECK_FONT  = Font(size=13)

# ── Bordures ──────────────────────────────────────────────────────────────────
_T  = Side(style="thin",   color="BFBFBF")
_M  = Side(style="medium", color="808080")
_THIN_BORDER = Border(left=_T, right=_T, top=_T, bottom=_T)
_MED_BORDER  = Border(left=_M, right=_M, top=_M, bottom=_M)

# ── Alignements ───────────────────────────────────────────────────────────────
_ROT    = Alignment(text_rotation=90, horizontal="center", vertical="bottom",  wrap_text=True)
_VROT   = Alignment(text_rotation=90, horizontal="center", vertical="center",  wrap_text=True)
_CTR    = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ── Symbole case à cocher ─────────────────────────────────────────────────────
_BOX = "☐"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _s(ws, row: int, col: int, value=None, *, fill=None, font=None,
       align=None, border=None):
    """Écriture rapide d'une cellule avec styles optionnels."""
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if fill   is not None: c.fill      = fill
    if font   is not None: c.font      = font
    if align  is not None: c.alignment = align
    if border is not None: c.border    = border
    return c


def _col_letter(n: int) -> str:
    return get_column_letter(n)


# ── Fonction principale ───────────────────────────────────────────────────────

_CHECK_ON  = "✓"   # case cochée dans l'export


def _write_df_to_sheet(wb, sheet_name: str, df) -> None:
    """Écrit un DataFrame dans une feuille openpyxl (données brutes, sans style)."""
    import pandas as pd
    ws = wb.create_sheet(sheet_name)
    for c_idx, col in enumerate(df.columns, 1):
        ws.cell(1, c_idx).value = str(col)
    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            # Convertir NaN/NaT en chaîne vide
            if val is None or (isinstance(val, float) and val != val):
                val = ""
            ws.cell(r_idx, c_idx).value = val


def export_protocol_excel(df_comp, meta: dict, path: str,
                           states: dict | None = None,
                           df_map=None) -> None:
    """Génère le fichier Excel complet (protocole + métadonnées).

    Le fichier produit contient 4 feuilles :
      - Protocole      : feuille de paillasse stylée (cases à cocher)
      - Volumes        : df_comp brut (rechargeable comme métadonnées)
      - Correspondance : df_map brut (rechargeable comme métadonnées)
      - EtatProtocole  : états des cases (rechargeable par l'application)

    Parameters
    ----------
    df_comp : pd.DataFrame
        Tableau des volumes (colonnes : Réactif, Concentration, Unité, Tube 1, …)
    meta : dict
        Informations d'en-tête : nom_manip, date, lieu, coordinateur, operateur
    path : str
        Chemin de sortie (.xlsx)
    states : dict | None
        États des cases à cocher issus de ProtocolDialog.get_states().
        Clés : ("verif", r_idx) | ("coord", r_idx, t_idx) | ("oper", r_idx, t_idx).
    df_map : pd.DataFrame | None
        Tableau de correspondance spectres ↔ tubes (optionnel).
    """
    tube_cols = mm.get_tube_columns(df_comp)
    n_tubes   = len(tube_cols)
    states    = states or {}

    def _box(key) -> str:
        return _CHECK_ON if states.get(key, False) else _BOX

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Protocole"

    # ── Indices de colonnes ───────────────────────────────────────────────────
    COL_LABEL  = 1   # "Réactifs" — label vertical orange (fusionné)
    COL_REACT  = 2   # nom du réactif
    COL_CONC   = 3   # concentration
    COL_UNIT   = 4   # unité
    COL_VERIF  = 5   # vérification solution (☐)
    COL_FIRST_TUBE = 6   # début des colonnes tubes

    # ── Indices de lignes ─────────────────────────────────────────────────────
    ROW_META1  = 1   # nom manip / date / lieu
    ROW_META2  = 2   # coordinateur / opérateur
    ROW_SEP    = 3   # ligne vide
    ROW_HDR    = 4   # en-têtes du tableau (pivotés)
    ROW_DATA   = 5   # première ligne de données

    n_data = len(df_comp)
    last_data_row = ROW_DATA + n_data - 1

    # ── En-tête d'information manip ───────────────────────────────────────────
    info_line1 = [
        ("Nom de la manip :", meta.get("nom_manip", "")),
        ("Date :",            meta.get("date", "")),
        ("Lieu :",            meta.get("lieu", "")),
    ]
    info_line2 = [
        ("Coordinateur :", meta.get("coordinateur", "")),
        ("Opérateur :",   meta.get("operateur", "")),
    ]
    for i, (label, val) in enumerate(info_line1):
        ws.cell(ROW_META1, COL_REACT + i * 4).value = label
        ws.cell(ROW_META1, COL_REACT + i * 4).font  = _BOLD
        ws.cell(ROW_META1, COL_REACT + i * 4 + 1).value = val
        ws.cell(ROW_META1, COL_REACT + i * 4 + 1).font  = _DARK
    for i, (label, val) in enumerate(info_line2):
        ws.cell(ROW_META2, COL_REACT + i * 4).value = label
        ws.cell(ROW_META2, COL_REACT + i * 4).font  = _BOLD
        ws.cell(ROW_META2, COL_REACT + i * 4 + 1).value = val
        ws.cell(ROW_META2, COL_REACT + i * 4 + 1).font  = _DARK

    # ── Colonne "Réactifs" — fusionnée verticalement sur toute la zone données ─
    ws.merge_cells(
        start_row=ROW_HDR, start_column=COL_LABEL,
        end_row=last_data_row, end_column=COL_LABEL,
    )
    c = ws.cell(ROW_HDR, COL_LABEL)
    c.value     = "Réactifs"
    c.fill      = _REACT_SIDE
    c.font      = _WHITE_BOLD
    c.alignment = _VROT
    c.border    = _MED_BORDER

    # ── En-têtes des colonnes fixes ───────────────────────────────────────────
    for col, label, fill in [
        (COL_REACT, "Réactif",               _META_HDR),
        (COL_CONC,  "Concentration",          _META_HDR),
        (COL_UNIT,  "Unité",                  _META_HDR),
        (COL_VERIF, "Vérification\nsolution", _VERIF_HDR),
    ]:
        font = _WHITE_BOLD if fill is _VERIF_HDR else _BOLD
        _s(ws, ROW_HDR, col, label, fill=fill, font=font,
           align=_ROT, border=_THIN_BORDER)

    # ── En-têtes des colonnes tubes ───────────────────────────────────────────
    sub_defs = [
        ("Volume (µL)",  _TUBE_HDR,  _HDR_FONT),
        ("Coordinateur", _COORD_HDR, _HDR_FONT),
        ("Opérateur",    _OPER_HDR,  _HDR_FONT),
    ]
    for t_idx, tube_name in enumerate(tube_cols):
        base = COL_FIRST_TUBE + t_idx * 3
        for sub, (sub_label, fill, font) in enumerate(sub_defs):
            label = tube_name if sub == 0 else sub_label
            _s(ws, ROW_HDR, base + sub, label,
               fill=fill, font=font, align=_ROT, border=_THIN_BORDER)

    # ── Données ───────────────────────────────────────────────────────────────
    for r_idx, (_, row) in enumerate(df_comp.iterrows()):
        dr   = ROW_DATA + r_idx
        rfil = _ROW_A if r_idx % 2 == 0 else _ROW_B

        # Réactif
        _s(ws, dr, COL_REACT, str(row.get("Réactif", "")),
           fill=rfil, font=_DARK, align=_LEFT, border=_THIN_BORDER)

        # Concentration
        raw_conc = row.get("Concentration", "")
        try:
            conc_val = float(raw_conc)
        except (ValueError, TypeError):
            conc_val = str(raw_conc) if raw_conc not in (None, "") else ""
        _s(ws, dr, COL_CONC, conc_val,
           fill=rfil, font=_DARK, align=_CTR, border=_THIN_BORDER)

        # Unité
        _s(ws, dr, COL_UNIT, str(row.get("Unité", "")),
           fill=rfil, font=_DARK, align=_CTR, border=_THIN_BORDER)

        # Vérification solution
        _s(ws, dr, COL_VERIF, _box(("verif", r_idx)),
           fill=_VERIF_CELL, font=_CHECK_FONT, align=_CTR, border=_THIN_BORDER)

        # Mode compte-goutte : pas > 0 → valeurs stockées en n_gouttes
        try:
            pas = float(row.get("Pas (µL)", 0) or 0)
        except (ValueError, TypeError):
            pas = 0.0

        # Volumes + cases à cocher par tube
        for t_idx, tube_name in enumerate(tube_cols):
            base = COL_FIRST_TUBE + t_idx * 3
            raw_vol = row.get(tube_name, "")
            try:
                v = float(raw_vol)
                if pas > 0:
                    vol_val = f"{int(round(v))} gouttes"
                else:
                    vol_val = int(v) if v == int(v) else v
            except (ValueError, TypeError):
                vol_val = str(raw_vol) if raw_vol not in (None, "") else ""

            _s(ws, dr, base,     vol_val,                            fill=rfil,      font=_DARK,       align=_CTR, border=_THIN_BORDER)
            _s(ws, dr, base + 1, _box(("coord", r_idx, t_idx)),     fill=rfil,      font=_CHECK_FONT, align=_CTR, border=_THIN_BORDER)
            _s(ws, dr, base + 2, _box(("oper",  r_idx, t_idx)),     fill=rfil,      font=_CHECK_FONT, align=_CTR, border=_THIN_BORDER)

    # ── Dimensions des colonnes et lignes ─────────────────────────────────────
    ws.row_dimensions[ROW_HDR].height = 90
    for r in range(ROW_DATA, last_data_row + 1):
        ws.row_dimensions[r].height = 20

    ws.column_dimensions[_col_letter(COL_LABEL)].width = 4
    ws.column_dimensions[_col_letter(COL_REACT)].width = 20
    ws.column_dimensions[_col_letter(COL_CONC) ].width = 10
    ws.column_dimensions[_col_letter(COL_UNIT) ].width = 8
    ws.column_dimensions[_col_letter(COL_VERIF)].width = 7

    for t_idx in range(n_tubes):
        base = COL_FIRST_TUBE + t_idx * 3
        ws.column_dimensions[_col_letter(base)    ].width = 9   # volume
        ws.column_dimensions[_col_letter(base + 1)].width = 5   # coordinateur
        ws.column_dimensions[_col_letter(base + 2)].width = 5   # opérateur

    # Figer les volets : colonne Réactif visible en scrollant à droite
    ws.freeze_panes = ws.cell(ROW_DATA, COL_FIRST_TUBE)

    # ── Feuilles de données (rechargement comme métadonnées) ──────────────────
    _write_df_to_sheet(wb, "Volumes", df_comp)
    if df_map is not None:
        _write_df_to_sheet(wb, "Correspondance", df_map)
    # EtatProtocole
    if states:
        import pandas as pd
        rows = []
        for key, checked in states.items():
            if key[0] == "verif":
                rows.append({"type": "verif", "r_idx": key[1], "t_idx": -1, "checked": int(checked)})
            else:
                rows.append({"type": key[0], "r_idx": key[1], "t_idx": key[2], "checked": int(checked)})
        _write_df_to_sheet(wb, "EtatProtocole", pd.DataFrame(rows))

    wb.save(path)
