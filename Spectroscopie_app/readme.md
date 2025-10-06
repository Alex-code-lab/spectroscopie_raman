# Ramanalyze — Application d’analyse de spectres Raman

> **PySide6 · pandas · plotly · pybaselines**  
> Outil graphique pour assembler, visualiser et analyser des spectres Raman à partir de fichiers `.txt` et de métadonnées **Excel/CSV**.

---

## Table des matières

- [Aperçu](#aperçu)
- [Fonctionnalités](#fonctionnalités)
- [Structure du dépôt](#structure-du-dépôt)
- [Installation](#installation)
- [Lancement de l’application](#lancement-de-lapplication)
- [How-To (pas à pas)](#how-to-pas-à-pas)
  - [1) Onglet Fichiers](#1-onglet-fichiers)
  - [2) Onglet Métadonnées](#2-onglet-métadonnées)
  - [3) Onglet Spectres](#3-onglet-spectres)
  - [4) Onglet Analyse](#4-onglet-analyse)
- [Formats de fichiers attendus](#formats-de-fichiers-attendus)
- [Détails techniques (modules)](#détails-techniques-modules)
- [Export & sauvegardes](#export--sauvegardes)
- [Bonnes pratiques & conseils](#bonnes-pratiques--conseils)
- [Dépannage (FAQ)](#dépannage-faq)
- [Crédits](#crédits)
- [Licence](#licence)

---

## Aperçu

**Ramanalyze** est une application de bureau (Qt / PySide6) qui guide l’utilisateur :
1. **Sélection** des spectres `.txt`.
2. **Association** avec un fichier de **métadonnées** (Excel/CSV).
3. **Assemblage** et **correction de baseline**.
4. **Visualisation** des spectres corrigés.
5. **Analyse** des **pics** (intensités locales) et **ratios**.
6. **Export** des résultats.

---

## Fonctionnalités

- ✅ Sélecteur de fichiers **.txt** avec navigation, ajout/retrait, liste d’éléments sélectionnés.
- ✅ Lecture **Excel (.xlsx/.xls)** **et** **CSV (.csv)** pour les métadonnées.
- ✅ **Détection automatique** du séparateur CSV.
- ✅ **Correction de baseline** (algorithme *modpoly*, `pybaselines`), ordre 5 par défaut.
- ✅ **Assemblage** des données (spectres + métadonnées) et **prévisualisation**.
- ✅ **Sauvegarde** du fichier combiné **en CSV ou Excel (au choix)**.
- ✅ **Tracé interactif** des spectres combinés (Plotly).
- ✅ **Analyse des pics** (intensité max autour de pics cibles, tolérance paramétrable).
- ✅ **Ratios** automatiques pour toutes les paires de pics.
- ✅ **Export** des résultats d’analyse en **Excel** (feuilles `intensites` et `ratios`).

---

## Structure du dépôt

```
Spectroscopie_app/
├─ main.py                # Point d’entrée de l’application (fenêtre, onglets, aide intégrée)
├─ file_picker.py         # Onglet « Fichiers » : sélection et liste de .txt
├─ metadata_picker.py     # Onglet « Métadonnées » : choix Excel/CSV, assemblage, enregistrement combiné
├─ spectra_plot.py        # Onglet « Spectres » : tracé interactif (Plotly) depuis le combiné
├─ analysis_tab.py        # Onglet « Analyse » : intensités autour des pics, ratios, export
└─ data_processing.py     # Chargement .txt, baseline modpoly, fusion avec métadonnées
```

---

## Installation

### Prérequis
- **Python 3.9+**
- OS : macOS / Linux / Windows

### Environnement & dépendances

```bash
# Optionnel mais recommandé : créer un venv
python -m venv .venv
source .venv/bin/activate        # macOS/Linux
# .venv\Scripts\activate         # Windows (PowerShell/CMD)

# Installer les bibliothèques nécessaires
pip install PySide6 pandas numpy plotly pybaselines openpyxl
```

> **Note** : `openpyxl` est utilisé pour lire/écrire les fichiers Excel.  
> `PySide6` nécessite parfois des runtimes additionnels côté OS (Qt WebEngine).

---

## Lancement de l’application

Depuis la racine du projet (ou le dossier parent), exécuter :

```bash
python Spectroscopie_app/main.py
```

Une fenêtre **Ramanalyze** s’ouvre avec 4 onglets : *Présentation*, *Fichiers*, *Métadonnées*, *Spectres*, *Analyse*.

---

## How-To (pas à pas)

### 1) Onglet Fichiers

- Naviguez dans vos dossiers (double-clic pour entrer).
- Sélectionnez un ou plusieurs **fichiers `.txt`** de spectroscopie Raman.
- Cliquez **Ajouter la sélection** pour les placer dans la liste de droite.
- Vous pouvez **Retirer** des éléments et **Vider la liste** si nécessaire.
- Le compteur en bas indique combien de fichiers sont prêts.

> ℹ️ Seuls les **fichiers `.txt`** sont listés (les dossiers restent navigables).

---

### 2) Onglet Métadonnées

1. **Sélectionner** votre fichier de métadonnées :
   - **Excel** : `.xlsx` / `.xls`
   - **CSV** : `.csv` (le séparateur est détecté automatiquement)
2. Un **aperçu** des métadonnées s’affiche (table triable).
3. Cliquez **Assembler (données au format txt + métadonnées)** :
   - Lecture des `.txt`
   - **Correction de baseline** (modpoly, ordre 5)
   - Fusion avec les métadonnées sur la clé **`Spectrum name`**
   - (Optionnel) Exclusion des échantillons `Cuvette BRB` si présents
4. Vérifiez le message de succès (lignes/colonnes du combiné).
5. Cliquez **Enregistrer le fichier** :
   - Choisissez **CSV (.csv)** **ou** **Excel (.xlsx)**
   - Le dernier chemin utilisé est mémorisé pour accélérer les sauvegardes.

> 💡 Si vos métadonnées sont en **CSV**, l’application crée un **Excel temporaire** en interne pour l’étape d’assemblage (compatibilité), mais vous pouvez **sauvegarder** ensuite **en CSV ou Excel** selon votre préférence.

---

### 3) Onglet Spectres

- Utilise le **fichier combiné** validé dans *Métadonnées*.
- Filtre automatiquement sur les fichiers sélectionnés (`file`).
- La **légende** privilégie `Sample description` si disponible, sinon le nom de fichier.
- Cliquez **Tracer avec baseline** pour afficher les courbes (Plotly, zoom/pan disponibles).

---

### 4) Onglet Analyse

1. Cliquez **Recharger le fichier combiné depuis Métadonnées** (si nécessaire).
2. Choisissez un **jeu de pics** :
   - **532 nm** : `1231, 1327, 1342, 1358, 1450`
   - **785 nm** : `412, 444, 471, 547, 1561`
3. Réglez la **Tolérance (cm⁻¹)** (par défaut 2.00).
4. Cliquez **Analyser les pics** :
   - Recherche de **l’intensité max** autour de chaque pic, **par fichier**.
   - Calcul **de tous les ratios** possibles entre pics (paires).
   - Jointure avec les métadonnées (si colonnes disponibles).
   - Affichage d’un **tableau triable** + **graphique interactif** des ratios vs `n(EGTA) (mol)`.
5. **Exporter résultats (Excel)…** pour sauvegarder :
   - Feuille `intensites` (intensités par pic)
   - Feuille `ratios` (ratios long-format)

---

## Formats de fichiers attendus

### Fichiers spectres `.txt`
- Doivent contenir un en-tête avec la ligne commençant par **`Pixel;`** (séparateur `;`, décimales `,`).
- Colonnes nécessaires après lecture :
  - **`Raman Shift`**
  - **`Dark Subtracted #1`**
- Le pipeline convertit en numérique, trie sur `Raman Shift`, calcule :
  - **`Intensity_corrected`** = `Dark Subtracted #1` – baseline(modpoly)
  - **`file`** = nom de fichier `.txt`
  - **`Spectrum name`** = `file` sans l’extension `.txt`

### Fichier métadonnées (Excel/CSV)
- Excel : `.xlsx` / `.xls` (lecture via `openpyxl`)  
- CSV : `.csv` (séparateur **auto-détecté**)
- La fusion attend la présence de **`Spectrum name`** (clé de jointure).
- Colonnes utiles courantes :
  - `Spectrum name`, `Sample description`, `n(EGTA) (mol)`, … (selon vos besoins)

---

## Détails techniques (modules)

### `data_processing.py`
- `load_spectrum_file(path, poly_order=5)` :
  - lit `.txt` (après la ligne `Pixel;`), parse `;` et décimales `,`
  - convertit en numérique, corrige baseline (**`pybaselines.Baseline.modpoly`**)
  - retourne un `DataFrame` avec `Raman Shift`, `Dark Subtracted #1`, `Intensity_corrected`, `file`
- `build_combined_dataframe(txt_files, metadata_path, poly_order=5, exclude_brb=True)` :
  - lit **Excel** (ou Excel temporaire si source CSV)
  - concatène les spectres valides, ajoute `Spectrum name`
  - fusion gauche sur `Spectrum name` avec les métadonnées
  - retire `Sample description == "Cuvette BRB"` si `exclude_brb=True`

### `file_picker.py`
- Navigation dans l’arborescence, filtre `.txt`, sélection multiple, liste de fichiers choisis.
- Méthodes :
  - `get_selected_files()` → `list[str]` des chemins `.txt` sélectionnés.

### `metadata_picker.py`
- Sélection **Excel/CSV**, aperçu tabulaire.
- **Assemblage** avec les `.txt` sélectionnés (via `build_combined_dataframe`).
- **Enregistrement** du combiné (dialogue **CSV/Excel**, mémorisation du dernier chemin).
- Détails :
  - Conversion **CSV → Excel temporaire** si nécessaire pour l’assembleur.
  - Détection auto du **séparateur CSV**.
  - `self.combined_df` conserve le DataFrame combiné en mémoire pour les autres onglets.

### `spectra_plot.py`
- Tracé **Plotly** depuis `self.combined_df` (via l’onglet Métadonnées).
- Filtrage sur les fichiers sélectionnés (`file`).
- Légende = `Sample description` si présent.

### `analysis_tab.py`
- Recharge `combined_df`, propose jeux de pics **532 nm** / **785 nm**.
- Tolérance réglable, calcul **intensités max** par fenêtre [pic ± tolérance].
- **Ratios** pour toutes les paires de pics.
- Graphique interactif des **ratios** vs **`n(EGTA) (mol)`**.
- Export **Excel** des résultats (feuilles `intensites`, `ratios`).

### `main.py`
- Démarre l’app Qt, crée les onglets, affiche une page **Présentation** (mode d’emploi synthétique).
- Lien **Sources** (dialogue d’information).

---

## Export & sauvegardes

- **Fichier combiné** (issue de l’assemblage) :
  - **CSV (.csv)** ou **Excel (.xlsx)** — au **choix** dans la boîte de dialogue.
- **Résultats d’analyse** :
  - Toujours **Excel (.xlsx)**, avec **`intensites`** et **`ratios`**.

---

## Bonnes pratiques & conseils

- Assurez-vous que les métadonnées contiennent **`Spectrum name`** correspondant aux `.txt` (sans extension).
- Évitez d’ouvrir trop de fichiers simultanément pour préserver la fluidité.
- Utilisez des noms clairs pour vos échantillons → meilleures légendes/rapports.
- Si vous travaillez en CSV, visez une cohérence d’encodage (UTF-8) et laissez la **détection de séparateur** active.

---

## Dépannage (FAQ)

**Q : L’assemblage échoue avec un CSV de métadonnées.**  
R : Le CSV est **converti en Excel temporaire** pour l’assembleur. Vérifiez :
- que le CSV contient bien **`Spectrum name`**,
- que le séparateur est standard (`,` `;` `\t`) — la détection est automatique,
- que l’encodage est lisible (UTF-8 recommandé).

**Q : Mon combiné n’a pas `Intensity_corrected`.**  
R : Les `.txt` doivent contenir `Raman Shift` et `Dark Subtracted #1` après la ligne `Pixel;`. Vérifiez le format (séparateur `;`, décimales `,`) et que les colonnes existent bien.

**Q : Le tracé n’affiche rien.**  
R : Vérifiez que :
- vous avez **assemblé** puis **enregistré** le combiné,
- les fichiers sélectionnés (onglet *Fichiers*) existent dans la colonne `file` du combiné,
- les colonnes `Raman Shift` et `Intensity_corrected` sont présentes.

**Q : L’analyse indique des colonnes manquantes.**  
R : Le combiné doit au minimum contenir `Raman Shift`, `Intensity_corrected`, `file`. Pour le graphique des ratios vs `n(EGTA) (mol)`, cette colonne doit exister dans les métadonnées.

**Q : Je veux toujours enregistrer par défaut en Excel.**  
R : À l’enregistrement, choisissez le filtre **Excel (.xlsx)**. L’application mémorise le **dernier chemin** utilisé. Une mémorisation du dernier **format** peut être ajoutée facilement si besoin.

---

## Crédits

- Développement : **Alexandre Souchaud**  
- Baseline : [`pybaselines`](https://github.com/derb12/pybaselines)  
- UI : **PySide6 / Qt**, graphiques **Plotly**

---

## Licence

Projet académique / interne.