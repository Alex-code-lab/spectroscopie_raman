#!/usr/bin/env python3
"""Réécriture des données Raman d'Angélina en deux tableaux Excel :

1) Tableau de composition des "tubes" (puits de plaque) avec les concentrations,
2) Tableau de correspondance entre noms de spectres et puits.

Entrée :
    - Un fichier CSV du type
      "AN336_and_AN344_laser_532nm_250nM_copper_variation_of_PAN_all_data_corrected_serie_11_points.csv"
      avec séparateur ";".

Sorties (dans le même dossier que le CSV) :
    - *_composition_tubes.xlsx
    - *_correspondance_spectre_tube.xlsx

Usage :
    python reecriture_angelina.py chemin/vers/fichier.csv
"""

from __future__ import annotations

import os
import sys
from typing import Tuple
import pandas as pd
import numpy as np


# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Définir ici le chemin vers le fichier CSV d'Angélina
CSV_PATH = "/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/Data_Angelina/AN336_and_AN344_laser_532nm_250nM_copper_variation_of_PAN_all_data_corrected_serie_11_points.csv"
# Exemple :
# CSV_PATH = "/Users/souchaud/Documents/Travail/.../fichier.csv"
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Paramètres instrumentaux (issus d'un fichier .txt type GC514_103.txt)
# Utilisés pour reconstruire des fichiers texte au format BWSpec-like.
CALIB_COEFS = {
    "a0": 531.016606800964,
    "a1": 0.0821624253029554,
    "a2": -1.81293471677812e-06,
    "a3": -8.76279014177826e-10,
}
LASER_WAVELENGTH_NM = 532.0602  # nm
REFERENCE_LEVEL = 65535.0       # valeur utilisée dans les fichiers GC514


def load_angelina_csv(path: str) -> pd.DataFrame:
    """Charge le CSV d'Angélina avec le bon séparateur et vérifie les colonnes clés."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Fichier introuvable : {path}")

    df = pd.read_csv(path, sep=";")

    required = [
        "Spectrum",
        "Meas_sample",
        "C.Cu (nM)",
        "C.EGTA (nM)",
        "C.PAN (nM)",
        "C.NPs (microM)",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans le CSV : {missing}")

    return df


def build_composition_table(df: pd.DataFrame) -> pd.DataFrame:
    """Construit le tableau de composition des puits (tubes).

    On considère que chaque puits (Meas_sample, ex. A1, B3, ...) est l'équivalent d'un tube.
    On regroupe donc par Meas_sample et on garde une seule ligne par puits.

    Colonnes produites (entre autres) :
        - Tube (ex-A1, B3, ...)
        - C (EGTA) (M)
        - C (Cu) (M)
        - C (PAN) (M)
        - C (NPs) (M)
        - et quelques colonnes contextuelles (NPs_batch, P (%), ...)
    """
    group_cols = [
        "Meas_sample",
        "C.Cu (nM)",
        "C.EGTA (nM)",
        "C.PAN (nM)",
        "C.NPs (microM)",
    ]
    # On ajoute des colonnes optionnelles si elles existent (batch, etc.)
    optional_cols = ["NPs_batch", "P (%)", "t (s)", "n", "excess.titrant"]
    for col in optional_cols:
        if col in df.columns and col not in group_cols:
            group_cols.append(col)

    comp = df[group_cols].drop_duplicates(subset=["Meas_sample"]).copy()

    # Renommer Meas_sample en Tube
    comp = comp.rename(columns={"Meas_sample": "Tube"})

    # Conversion en mol/L (M)
    # nM -> M : 1 nM = 1e-9 M
    comp["C (EGTA) (M)"] = comp["C.EGTA (nM)"] * 1e-9
    comp["C (Cu) (M)"] = comp["C.Cu (nM)"] * 1e-9
    comp["C (PAN) (M)"] = comp["C.PAN (nM)"] * 1e-9
    # microM -> M : 1 µM = 1e-6 M
    comp["C (NPs) (M)"] = comp["C.NPs (microM)"] * 1e-6

    # On place les colonnes principales en tête
    main_cols = [
        "Tube",
        "C (EGTA) (M)",
        "C (Cu) (M)",
        "C (PAN) (M)",
        "C (NPs) (M)",
    ]
    other_cols = [c for c in comp.columns if c not in main_cols]
    comp = comp[main_cols + other_cols]

    return comp


def build_mapping_table(df: pd.DataFrame) -> pd.DataFrame:
    """Construit le tableau de correspondance spectre → tube.

    On garde une seule ligne par spectre (nom de spectre), avec :
        - Spectrum name : nom du spectre (ex. AN336_01)
        - Tube : le puits associé (Meas_sample, ex. A1)
        - Exp, Series, Rep_exp pour garder le contexte expérimental.
    """
    cols = ["Spectrum", "Meas_sample"]
    for extra in ["Exp", "Series", "Rep_exp"]:
        if extra in df.columns:
            cols.append(extra)

    m = df[cols].drop_duplicates(subset=["Spectrum"]).copy()
    m = m.rename(columns={"Spectrum": "Spectrum name", "Meas_sample": "Tube"})

    # On ordonne gentiment les colonnes
    first_cols = ["Spectrum name", "Tube"]
    other_cols = [c for c in m.columns if c not in first_cols]
    m = m[first_cols + other_cols]

    return m


# Détecte les colonnes d'intensité spectrale (hors métadonnées)
def _detect_spectral_columns(df: pd.DataFrame) -> list[str]:
    """Détecte les colonnes contenant les intensités spectrales.

    Hypothèse : chaque ligne du DataFrame correspond à un spectre,
    et les colonnes "métadonnées" sont celles connues (Spectrum, Meas_sample, concentrations...).
    Toutes les autres colonnes numériques sont considérées comme des points du spectre.
    """
    meta_cols = {
        "Spectrum",
        "Meas_sample",
        "C.Cu (nM)",
        "C.EGTA (nM)",
        "C.PAN (nM)",
        "C.NPs (microM)",
        "NPs_batch",
        "P (%)",
        "t (s)",
        "n",
        "excess.titrant",
        "Exp",
        "Series",
        "Rep_exp",
    }
    spectral_cols: list[str] = []
    for col in df.columns:
        if col in meta_cols:
            continue
        # On ne garde que les colonnes numériques
        if pd.api.types.is_numeric_dtype(df[col]):
            spectral_cols.append(col)
    if not spectral_cols:
        raise ValueError(
            "Impossible de détecter les colonnes spectrales dans le CSV "
            "(aucune colonne numérique hors métadonnées)."
        )
    return spectral_cols


def _fmt(value: float, ndigits: int = 4) -> str:
    """Formate un flottant avec une virgule décimale (style fichier GC514)."""
    return f"{value:.{ndigits}f}".replace(".", ",")


def _compute_axes_for_pixel(pixel: int) -> tuple[float, float, float]:
    """Calcule (wavelength_nm, wavenumber_cm1, raman_shift_cm1) pour un pixel donné."""
    a0 = CALIB_COEFS["a0"]
    a1 = CALIB_COEFS["a1"]
    a2 = CALIB_COEFS["a2"]
    a3 = CALIB_COEFS["a3"]
    x = float(pixel)
    wavelength_nm = a0 + a1 * x + a2 * x * x + a3 * x * x * x
    # Conversion en nombre d'onde (cm^-1) : nu = 1 / lambda(cm) = 1e7 / lambda(nm)
    wavenumber_cm1 = 1e7 / wavelength_nm
    laser_wavenumber_cm1 = 1e7 / LASER_WAVELENGTH_NM
    raman_shift_cm1 = laser_wavenumber_cm1 - wavenumber_cm1
    return wavelength_nm, wavenumber_cm1, raman_shift_cm1


def export_spectra_as_txt(df: pd.DataFrame, out_dir: str) -> list[str]:
    """Exporte chaque spectre individuel dans un fichier .txt 'BWSpec-like'
    en regroupant les lignes par 'Spectrum' (format long du fichier d’Angélina).

    Chaque fichier .txt contiendra :
        - les colonnes BWSpec classiques,
        - la colonne 'raman_shift' (Rshift),
        - la colonne 'spectra_corrected' (Spectra_corrected).
    """
    # Vérifications des colonnes indispensables
    if "Spectrum" not in df.columns:
        raise ValueError("La colonne 'Spectrum' est absente du CSV.")
    if "Rshift" not in df.columns or "Spectra_corrected" not in df.columns:
        raise ValueError("Colonnes 'Rshift' et/ou 'Spectra_corrected' absentes du CSV.")

    generated_paths: list[str] = []
    os.makedirs(out_dir, exist_ok=True)

    # Regroupement par nom de spectre
    grouped = df.groupby("Spectrum")

    for spec_name, sub in grouped:
        # Tri des points par décalage Raman
        sub = sub.sort_values("Rshift")
        rshifts = sub["Rshift"].to_numpy(dtype=float)
        intensities = sub["Spectra_corrected"].to_numpy(dtype=float)
        pixel_num = len(intensities)

        # Nettoyage du nom du fichier
        safe_name = "".join(c if c.isalnum() or c in ("_", "-", ".") else "_" for c in str(spec_name))
        out_path = os.path.join(out_dir, f"{safe_name}.txt")

        with open(out_path, "w", encoding="utf-8") as f:
            # Écriture de l'en-tête BWSpec
            f.write("File Version;BWSpec4.02_14\n")
            f.write("Date;\n")
            f.write("title;BWS415-532S\n")
            f.write("model;BTC162E-532S-SYS\n")
            f.write("c code;RLZ\n")
            f.write("operator;\n")
            f.write("port1;0\n")
            f.write("baud1;3\n")
            f.write("pixel_start;0\n")
            f.write(f"pixel_end;{pixel_num - 1}\n")
            f.write("step;1\n")
            f.write("units;3\n")
            f.write("bkcolor;16777215\n")
            f.write("show_mode;3\n")
            f.write("data_mode;0\n")
            f.write("pixel_mode;0\n")
            f.write("intigration times(ms);5000\n")
            f.write("average number;3\n")
            f.write("time_multiply;1\n")
            f.write("spectrometer_type ;14\n")
            f.write("yaxis;0\n")
            f.write("yaxis_min;0\n")
            f.write("yaxis_max;65535\n")
            f.write("xaxis;1\n")
            f.write("xaxis_min;11\n")
            f.write("xaxis_max;1305\n")
            f.write("irrands_DispWLMin;100\n")
            f.write("irrands_DispWLMax;1000\n")
            f.write("yaxis_min_6;0\n")
            f.write("yaxis_max_6;0\n")
            f.write("irradiance_unit;0\n")
            f.write("Color_Data_Flag;0\n")
            f.write("Color_StartWL;532\n")
            f.write("Color_EndWL;684\n")
            f.write("Color_IncWL;10\n")
            f.write("power_unit_index;3\n")
            f.write("photometric_index;0\n")
            f.write("Illuminant_index;3\n")
            f.write("observer_index;0\n")
            f.write("lab_l;0\n")
            f.write("lab_a;0\n")
            f.write("lab_b;0\n")
            f.write("radiometric_flag;0\n")
            f.write(f"coefs_a0;{_fmt(CALIB_COEFS['a0'], 12)}\n")
            f.write(f"coefs_a1;{_fmt(CALIB_COEFS['a1'], 13)}\n")
            f.write("coefs_a2;-1.81293471677812E-06\n")
            f.write("coefs_a3;-8.76279014177826E-10\n")
            f.write("coefs_b0;0\n")
            f.write("coefs_b1;0\n")
            f.write("coefs_b2;0\n")
            f.write("coefs_b3;0\n")
            f.write("all_data_save;1\n")
            f.write("select_option;-1\n")
            f.write("interval_time;5000\n")
            f.write(f"pixel_num;{pixel_num}\n")
            f.write("sel_pixel_start;2016\n")
            f.write("sel_pixel_end;0\n")
            f.write("sel_pixel_delta;1\n")
            f.write("dark_compensate;0\n")
            f.write("dark_compensate_value_1;0\n")
            f.write("dark_compensate_value_2;0\n")
            f.write("dark_compensate_value_3;0\n")
            f.write("monitor pixel_0;0\n")
            f.write("monitor pixel_1;0\n")
            f.write("monitor pixel_2;0\n")
            f.write("monitor pixel_3;0\n")
            f.write("monitor pixel_4;0\n")
            f.write("monitor pixel_5;0\n")
            f.write("vertical_select_flag;0\n")
            f.write("vertical_line3;0\n")
            f.write("vertical_line4;0\n")
            f.write("vertical_line3_wv;1\n")
            f.write("vertical_line4_wv;1\n")
            f.write("vertical_line_flag;0\n")
            f.write("vertical_line_ratio;0\n")
            f.write(f"laser_wavelength;{_fmt(LASER_WAVELENGTH_NM, 4)}\n")
            f.write("laser_powerlevel;70\n")
            f.write("overlay_js;0\n")
            f.write("Relative Intensity Correction Flag;0\n")

            # Ligne d’en-tête des données (avec 2 colonnes supplémentaires)
            f.write(
                "Pixel;Wavelength;Wavenumber;Raman Shift;Dark;Reference;Raw data #1;"
                "Dark Subtracted #1;%TR #1;Absorbance #1;Irradiance (lumen) #1;"
                "raman_shift;spectra_corrected;\n"
            )

            # Écriture des points spectrales
            for pixel_idx, (rs_val, intensity) in enumerate(zip(rshifts, intensities)):
                wl_nm, wn_cm1, rs_cm1 = _compute_axes_for_pixel(pixel_idx)

                f.write(
                    f"{pixel_idx};"
                    f"{_fmt(wl_nm, 2)};"
                    f"{_fmt(wn_cm1, 2)};"
                    f"{_fmt(rs_cm1, 2)};"
                    f"{_fmt(0.0, 4)};"
                    f"{_fmt(REFERENCE_LEVEL, 4)};"
                    f"{_fmt(float(intensity), 4)};"
                    f"{_fmt(float(intensity), 4)};"
                    f"{_fmt(0.0, 4)};"
                    f"{_fmt(0.0, 4)};"
                    f"{_fmt(0.0, 4)};"
                    f"{_fmt(float(rs_val), 4)};"
                    f"{_fmt(float(intensity), 4)};"
                    "\n"
                )

        generated_paths.append(out_path)

    return generated_paths


def rewrite_for_angelina(csv_path: str) -> Tuple[str, str]:
    """Pipeline complet : lit le CSV et écrit les deux fichiers Excel.

    Returns
    -------
    (path_comp, path_map): chemins des deux fichiers générés.
    """
    df = load_angelina_csv(csv_path)

    comp = build_composition_table(df)
    mapping = build_mapping_table(df)

    base_dir = os.path.dirname(csv_path)
    base_name = os.path.splitext(os.path.basename(csv_path))[0]

    comp_path = os.path.join(base_dir, f"{base_name}_composition_tubes.xlsx")
    map_path = os.path.join(base_dir, f"{base_name}_correspondance_spectre_tube.xlsx")
    txt_dir = os.path.join(base_dir, f"{base_name}_spectres_txt")

    comp.to_excel(comp_path, index=False)
    mapping.to_excel(map_path, index=False)
    export_spectra_as_txt(df, txt_dir)

    return comp_path, map_path


def main(argv: list[str] | None = None) -> int:
    csv_path = CSV_PATH
    if not csv_path or not os.path.exists(csv_path):
        print(f"ERREUR : le fichier défini dans CSV_PATH est introuvable : {csv_path}")
        return 1

    try:
        comp_path, map_path = rewrite_for_angelina(csv_path)
    except Exception as e:
        print(f"ERREUR : {e}")
        return 1

    print("Fichiers générés :")
    print(f"  - Tableau de composition : {comp_path}")
    print(f"  - Tableau de correspondance : {map_path}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
