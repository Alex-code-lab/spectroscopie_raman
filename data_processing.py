# data_processing.py

import os
import re
import numpy as np
import pandas as pd
from pybaselines import Baseline


def load_spectrum_file(path: str, poly_order: int = 5) -> pd.DataFrame | None:
    """
    Charge un fichier .txt de spectroscopie Raman, applique une correction de baseline,
    et retourne un DataFrame avec colonnes utiles.

    Args:
        path: chemin vers le fichier .txt
        poly_order: ordre du polynôme pour la baseline

    Returns:
        DataFrame avec colonnes ["Raman Shift", "Dark Subtracted #1", "Intensity_corrected", "file"]
        ou None si erreur
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        header_idx = next(i for i, line in enumerate(lines) if line.strip().startswith("Pixel;"))

        df = pd.read_csv(
            path,
            skiprows=header_idx,
            sep=";",
            decimal=",",
            encoding="utf-8",
            skipinitialspace=True,
            na_values=["", " ", "   ", "\t"],
            keep_default_na=True,
        )

        # supprimer dernière colonne si vide
        if df.columns[-1].startswith("Unnamed"):
            df = df.iloc[:, :-1]

        # conversion numérique
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = pd.to_numeric(df[col], errors="coerce")

        if not {"Raman Shift", "Dark Subtracted #1"}.issubset(df.columns):
            return None

        temp = df[["Raman Shift", "Dark Subtracted #1"]].copy()
        temp = temp.dropna()

        # baseline correction
        x = temp["Raman Shift"].values
        y = temp["Dark Subtracted #1"].values
        baseline_fitter = Baseline(x)
        baseline, _ = baseline_fitter.modpoly(y, poly_order=poly_order)

        temp["Intensity_corrected"] = y - baseline
        temp["file"] = os.path.basename(path)

        return temp.dropna(subset=["Intensity_corrected"])

    except Exception as e:
        print(f"Erreur lors du chargement de {path}: {e}")
        return None


def build_combined_dataframe_from_df(
    txt_files: list[str],
    metadata_df: pd.DataFrame,
    poly_order: int = 5,
    exclude_brb: bool = True,
) -> pd.DataFrame:
    """
    Construit le DataFrame fusionné (spectres + métadonnées) à partir
    d'un DataFrame de métadonnées déjà chargé.

    Args:
        txt_files: liste des chemins vers les fichiers .txt
        metadata_df: DataFrame contenant au moins la colonne 'Spectrum name'
        poly_order: ordre du polynôme pour la baseline
        exclude_brb: si True, supprime les échantillons "Cuvette BRB"

    Returns:
        combined_df: DataFrame fusionné prêt pour l'analyse
    """
    all_data = []
    for path in txt_files:
        temp = load_spectrum_file(path, poly_order=poly_order)
        if temp is not None:
            all_data.append(temp)

    if not all_data:
        raise ValueError("Aucun spectre valide n'a été chargé.")

    spectra_df = pd.concat(all_data, ignore_index=True)
    spectra_df["Spectrum name"] = spectra_df["file"].str.replace(".txt", "", regex=False)

    combined_df = pd.merge(
        spectra_df,
        metadata_df,
        on="Spectrum name",
        how="left"
    )

    if exclude_brb and "Sample description" in combined_df.columns:
        combined_df = combined_df[combined_df["Sample description"] != "Cuvette BRB"].copy()

    return combined_df


def build_combined_dataframe(
    txt_files: list[str],
    metadata_path: str,
    poly_order: int = 5,
    exclude_brb: bool = True,
) -> pd.DataFrame:
    """
    Construit le DataFrame fusionné (spectres + métadonnées) à partir d'un fichier Excel ou CSV.

    Args:
        txt_files: liste des chemins vers les fichiers .txt
        metadata_path: chemin vers le fichier de métadonnées (Excel ou CSV)
        poly_order: ordre du polynôme pour la baseline
        exclude_brb: si True, supprime les échantillons "Cuvette BRB"

    Returns:
        combined_df: DataFrame fusionné prêt pour l'analyse
    """
    # Charger les métadonnées en fonction de l'extension
    if metadata_path.lower().endswith(".csv"):
        metadata_df = pd.read_csv(metadata_path, sep=None, engine="python")
    else:
        metadata_df = pd.read_excel(metadata_path, skiprows=1)

    return build_combined_dataframe_from_df(
        txt_files,
        metadata_df,
        poly_order=poly_order,
        exclude_brb=exclude_brb,
    )