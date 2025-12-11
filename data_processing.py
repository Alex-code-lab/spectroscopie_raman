# data_processing.py

import os
import re
import numpy as np
import pandas as pd
from pybaselines import Baseline


def load_spectrum_file(path: str, poly_order: int = 5, apply_baseline: bool = True) -> pd.DataFrame | None:
    """
    Charge un fichier .txt de spectroscopie Raman, applique une correction de baseline,
    et retourne un DataFrame avec colonnes utiles.

    Args:
        path: chemin vers le fichier .txt
        poly_order: ordre du polynôme pour la baseline

    Returns:
        DataFrame avec colonnes ["Raman Shift", "Dark Subtracted #1", "Intensity_corrected", "file"]
        ou None si erreur / format non reconnu.
    """
    try:
        # --- Lecture brute + repérage de la ligne "Pixel;" ---
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

        # Supprimer dernière colonne si vide / Unnamed
        if df.columns[-1].startswith("Unnamed"):
            df = df.iloc[:, :-1]

        # Nettoyage des noms de colonnes (espaces, insécables, etc.)
        df.columns = [str(c).strip().replace("\xa0", " ") for c in df.columns]

        # Conversion numérique des colonnes texte
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # -------------------------------------------------------
        # 1) Harmonisation des noms de colonnes spectrales
        # -------------------------------------------------------
        cols_lower = {c.lower(): c for c in df.columns}

        # a) Colonne Raman Shift
        if "Raman Shift" not in df.columns:
            # Cherche des variantes possibles
            for cand in ["raman shift", "raman_shift", "rshift", "ramanshift"]:
                if cand in cols_lower:
                    df["Raman Shift"] = df[cols_lower[cand]]
                    break

        # b) Colonne Dark Subtracted #1 (intensité brute utilisée pour baseline)
        if "Dark Subtracted #1" not in df.columns:
            for cand in ["spectra_corrected", "spectra corrected", "intensity", "intensite"]:
                if cand in cols_lower:
                    df["Dark Subtracted #1"] = df[cols_lower[cand]]
                    break

        # Si après harmonisation on n'a toujours pas les colonnes, on abandonne
        if not {"Raman Shift", "Dark Subtracted #1"}.issubset(set(df.columns)):
            # Debug minimal dans la console
            print(
                f"[load_spectrum_file] Colonnes requises absentes pour {path}. "
                f"Colonnes présentes : {list(df.columns)}"
            )
            return None

        # -------------------------------------------------------
        # 2) Construction du DataFrame de travail
        # -------------------------------------------------------
        temp = df[["Raman Shift", "Dark Subtracted #1"]].copy()
        temp = temp.dropna(subset=["Raman Shift", "Dark Subtracted #1"])

        if temp.empty:
            print(f"[load_spectrum_file] Données vides après dropna pour {path}")
            return None

        x = temp["Raman Shift"].values
        y = temp["Dark Subtracted #1"].values

        # Tri par shift croissant
        order = np.argsort(x)
        x = x[order]
        y = y[order]

        # Baseline correction (optionnelle)
        if apply_baseline:
            try:
                baseline_fitter = Baseline(x)
                baseline, _ = baseline_fitter.modpoly(y, poly_order=poly_order)
                ycorr = y - baseline
            except Exception as be:
                print(f"[load_spectrum_file] Échec baseline pour {path} : {be}")
                baseline = np.zeros_like(y)
                ycorr = y
        else:
            # Pas de baseline : on garde le signal brut
            baseline = np.zeros_like(y)
            ycorr = y

        out = pd.DataFrame(
            {
                "Raman Shift": x,
                "Dark Subtracted #1": y,
                "Intensity_corrected": ycorr,
                "file": os.path.basename(path),
            }
        )

        return out.dropna(subset=["Intensity_corrected"])

    except Exception as e:
        print(f"Erreur lors du chargement de {path}: {e}")
        return None

def build_combined_dataframe_from_df(
    txt_files: list[str],
    metadata_df: pd.DataFrame,
    poly_order: int = 5,
    exclude_brb: bool = True,
    apply_baseline: bool = True,
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
        temp = load_spectrum_file(path, poly_order=poly_order, apply_baseline=apply_baseline)
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
        combined_df = combined_df[
            ~combined_df["Sample description"].isin(["Tube MQW", "Tube BRB", "BRB",
                                                     "Cuvette BRB", "MQW", "Cuvette MQW"])
        ].copy()
    return combined_df


def build_combined_dataframe(
    txt_files: list[str],
    metadata_path: str,
    poly_order: int = 5,
    exclude_brb: bool = True,
    apply_baseline: bool = True,
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
        apply_baseline=apply_baseline,
    )