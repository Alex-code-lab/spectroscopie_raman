import os
import re
import itertools
import numpy as np
import pandas as pd
from pybaselines import Baseline
from plotnine import ggplot, aes, geom_point, geom_line, theme_bw, labs, scale_color_brewer, theme, scale_x_continuous, ylim, xlim
import matplotlib.pyplot as plt
import plotly.express as px

EXPERIMENT_NAME = "AS004"
WAVELENGTH = "532nm"
EXPERIMENT_FOLDER = EXPERIMENT_NAME + "_" +WAVELENGTH
data_dir = os.path.join('/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/', EXPERIMENT_FOLDER)


# Lecture du fichier Excel de métadonnées :
metadata_path = "/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/"
metadata_df = pd.read_excel(metadata_path + EXPERIMENT_NAME + "_metadata.xlsx", skiprows=1)
metadata_df.head()


spectrocopy_files = [f for f in os.listdir(data_dir)
                    if f.endswith('.txt')]

# trie dans l'ordre croissant en extrayant la partie numérique
spectrocopy_files = sorted(
    spectrocopy_files,
    key=lambda x: int(re.search(r'_(\d+)\.txt$', x).group(1)) 
)

spectrocopy_files


# Exemple : on prend un seul spectre
file_path = os.path.join(data_dir, spectrocopy_files[1])
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

# Initialiser le modèle
baseline_fitter = Baseline(x)

# Méthode polynomial modpolyfit (comme en R)
baseline, params = baseline_fitter.modpoly(y, poly_order=5)

# Corriger le spectre
y_corrected = y - baseline

# Affichage
plt.figure(figsize=(8, 5))
plt.plot(x, y, label="Spectre brut", alpha=0.7)
plt.plot(x, baseline, label="Baseline estimée", linestyle="--")
plt.plot(x, y_corrected, label="Spectre corrigé", linewidth=1)
plt.xlabel("Raman Shift (cm⁻¹)")
plt.ylabel("Intensité (a.u.)")
plt.legend()
plt.title("Baseline correction Raman (modpolyfit)")
plt.show()


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
# cuvette BRB
cuvette_BRB = combined_df[combined_df["Sample description"] == "Cuvette BRB"]
# Supprimer les lignes avec "Cuvette BRB"
combined_df = combined_df[combined_df["Sample description"] != "Cuvette BRB"].copy()
# Print du résultat
combined_df.head(10)

# Sauvegarde
# combined_df.to_excel("Spectres_fusionnés.xlsx", index=False)


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


# Spectres des cuvettes BRB

fig = px.line(cuvette_BRB, 
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


# Spectres
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


# Liste des pics d’intérêt (en cm⁻¹)
peaks = [1231, 1327, 1342, 1358, 1450] # 532nm
peaks = [412, 444, 471, 547, 1561] # 785nm


# Largeur de la fenêtre de recherche autour de chaque pic
tolerance = 2  # ±5 cm⁻¹

results = []

for fname, group in combined_df.groupby("file"):
    spectrum = group.sort_values("Raman Shift")

    record = {"file": fname}

    for target in peaks:
        # on sélectionne la fenêtre autour du Raman Shift cible
        window = spectrum[
            (spectrum["Raman Shift"] >= target - tolerance) &
            (spectrum["Raman Shift"] <= target + tolerance)
        ]

        if not window.empty:
            # on prend le maximum corrigé dans cette zone
            record[f"I_{target}"] = window["Intensity_corrected"].max()
        else:
            record[f"I_{target}"] = np.nan

    results.append(record)

peak_intensities = pd.DataFrame(results)
peak_intensities.head()


for (target_a, target_b) in itertools.combinations(peaks, 2):
    peak_intensities[f"ratio_I_{target_a}_I_{target_b}"] = (
        peak_intensities[f"I_{target_a}"] / peak_intensities[f"I_{target_b}"]
    )
peak_intensities.head(10)


# Fusion des intensités avec les métadonnées
peak_intensities["Spectrum name"] = peak_intensities["file"].str.replace(".txt", "", regex=False)
merged = peak_intensities.merge(metadata_df, on="Spectrum name", how="left")
merged.head(10)

ratio_cols = [c for c in merged.columns if c.startswith("ratio_I_")]


# Restructurer les données en format long
df_ratios = merged.melt(
    id_vars=["n(EGTA) (mol)", "file"],
    value_vars=ratio_cols,
    var_name="Ratio",
    value_name="Value"
)


# Tracé
p = (
    ggplot(df_ratios, aes(x="n(EGTA) (mol)", y="Value", color="Ratio"))
    + geom_point(size=3)
    + geom_line(aes(group="Ratio"))
    + theme_bw()
    + labs(
        title="Rapports d’intensité Raman selon [EGTA]",
        x="Quantité EGTA (mol)",
        y="Rapport d’intensité (a.u.)"
    )
    + scale_color_brewer(type='qual', palette='Dark2')
    + scale_x_continuous(labels=lambda l: [f"{x:.0e}" for x in l])
    + theme(figure_size=(8, 5))
    # + ylim(0.4, 1.8)
    #  + xlim(0, 1.5e-9)
)

p