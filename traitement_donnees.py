import os
import re
import pandas as pd
from pybaselines import Baseline
import itertools
import plotly.express as px





EXPERIMENT_NAME = "AS004"
WAVELENGTH = "532nm"
EXPERIMENT_FOLDER = EXPERIMENT_NAME + "_" +WAVELENGTH

data_dir = os.path.join('/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/', EXPERIMENT_FOLDER)


spectrocopy_files = [f for f in os.listdir(data_dir)
                    if f.endswith('.txt')]

# trie dans l'ordre croissant en extrayant la partie numérique
spectrocopy_files = sorted(
    spectrocopy_files,
    key=lambda x: int(re.search(r'_(\d+)\.txt$', x).group(1)) 
)

print(spectrocopy_files)



all_data = []

for fname in spectrocopy_files:   # ta liste triée
    file_path = os.path.join(data_dir, fname)

    # repérer la ligne de l'en-tête
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    header_idx = next(i for i, line in enumerate(lines)
                      if line.strip().startswith('Pixel;'))

    # lire le fichier
    df = pd.read_csv(file_path,
                     skiprows=header_idx,
                     sep=";",
                     decimal=",",
                     encoding="utf-8",
                     skipinitialspace=True,
                     na_values=["", " ", "   ", "\t"],
                     keep_default_na=True)

    # supprimer la dernière colonne si vide
    if df.columns[-1].startswith("Unnamed"):
        df = df.iloc[:, :-1]

    # convertir en numérique
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # garder seulement les colonnes utiles
    temp = df[["Wavenumber", "Wavelength", "Raman Shift", "Dark Subtracted #1"]].copy()
    temp = temp.dropna()

    # baseline correction avec modpolyfit
    x = temp["Raman Shift"].values
    y = temp["Dark Subtracted #1"].values
    baseline_fitter = Baseline(x)
    baseline, _ = baseline_fitter.modpoly(y, poly_order=5)

    # ajouter la colonne corrigée
    temp["Intensity_corrected"] = y - baseline
    temp = temp.dropna(subset=["Intensity_corrected"])

    # ajouter une colonne "fichier"
    temp["file"] = fname

    # stocker dans la liste
    all_data.append(temp)

# fusionner tous les fichiers
spectra_df = pd.concat(all_data, ignore_index=True)

spectra_df.head()
# print(spectra_df["file"].unique())


# Lecture du fichier Excel de métadonnées :
metadata_df = pd.read_excel("/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/AS004_metadata.xlsx", skiprows=1)
metadata_df.head()

#  Création d'une clé de jointure dans spectra_df 
# (on enlève ".txt" pour matcher le nom type"AS00X_Y" dans l'Excel
spectra_df["Spectrum name"] = spectra_df["file"].str.replace(".txt", "", regex=False)

# Fusion des deux tables : 
combined_df = pd.merge(
    spectra_df,
    metadata_df,
    on="Spectrum name",
    how="left"   # garde toutes les lignes de spectra_df
)

# Print du résultat
print(combined_df.head())
print(combined_df.columns)

# Sauvegarde
# combined_df.to_excel("Spectres_fusionnés.xlsx", index=False)

from plotnine import ggplot, aes, geom_line, theme_bw, labs, scale_color_brewer

p = (
    ggplot(combined_df, aes(x="Raman Shift", y="Intensity_corrected", color="file"))
    + geom_line(size=0.1)
    + theme_bw()
    + labs(
        title="Spectres Raman",
        x="Raman Shift (cm⁻¹)",
        y="Intensité (a.u.)"
    )
    + scale_color_brewer(type='qual', palette='Paired')  # ou 'Dark2', 'Paired', etc.
)

p



fig = px.line(combined_df, 
              x="Raman Shift", 
              y="Intensity_corrected", 
              color="file",
              title="Spectres Raman interactifs")

fig.update_layout(
    xaxis_title="Raman Shift (cm⁻¹)",
    yaxis_title="Intensité (a.u.)",
    width=1900,   # largeur en pixels
    height=1000   # hauteur en pixels
)

fig.show()