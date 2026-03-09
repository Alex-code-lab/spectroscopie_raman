# Ramanalyze — Application d’analyse de spectres Raman

> **PySide6 · pandas · plotly · pybaselines · scikit-learn**
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
  - [5) Onglet Exploration (PCA)](#5-onglet-exploration-pca)
- [Comprendre la PCA en spectroscopie Raman](#comprendre-la-pca-en-spectroscopie-raman)
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

**Données & assemblage**
- ✅ Sélecteur de fichiers `.txt` avec navigation arborescente, ajout/retrait.
- ✅ Lecture métadonnées **Excel (.xlsx/.xls)** et **CSV** (séparateur auto-détecté).
- ✅ Correction de baseline (*modpoly*, `pybaselines`), ordre 5 par défaut.
- ✅ Validation automatique des spectres : plage physique (−200…10 000 cm⁻¹), déduplication des doubons de Raman Shift.
- ✅ Assemblage spectres + métadonnées et prévisualisation.
- ✅ Sauvegarde du fichier combiné **en CSV ou Excel** (au choix, dernier chemin mémorisé).

**Métadonnées**
- ✅ Tableau des volumes (réactifs × tubes) entièrement éditable.
- ✅ Génération automatique de volumes de titrant (**distribution gaussienne** centrée sur l’équivalence chimique).
- ✅ Concentrations finales dans chaque tube calculées automatiquement.
- ✅ Correspondance spectres ↔ tubes avec synchronisation automatique.
- ✅ Sauvegarde / chargement des métadonnées au format `.xlsx`.

**Visualisation**
- ✅ Tracé interactif Plotly (zoom, pan, export PNG, légende cliquable).

**Analyse quantitative**
- ✅ Intensité maximale autour de pics cibles (tolérance paramétrable).
- ✅ Ratios automatiques pour toutes les paires de pics.
- ✅ Graphique interactif des ratios vs variable de métadonnées.
- ✅ Export Excel (feuilles `intensites` et `ratios`).

**Exploration multivariée (PCA)**
- ✅ PCA corrélée sur les intensités aux pics sélectionnés.
- ✅ PCA libre sur les spectres complets (matrice normalisée, 2–15 composantes).
- ✅ Sous-onglets : Scores, Loadings, Reconstruction, Scatter matrix.
- ✅ Coloration des scores par n’importe quelle variable de métadonnées.

---

## Structure du dépôt

```
spectroscopie_raman/
├─ main.py                # Point d’entrée de l’application (fenêtre, onglets)
├─ file_picker.py         # Onglet « Fichiers » : sélection et liste de .txt
├─ metadata_creator.py    # Onglet « Métadonnées » : création/édition volumes + correspondance (UI Qt)
├─ metadata_model.py      # Couche données pure : calculs concentrations, fusion métadonnées, PCA (sans Qt)
├─ spectra_plot.py        # Onglet « Spectres » : tracé interactif (Plotly)
├─ analysis_tab.py        # Onglet « Analyse » : intensités autour des pics, ratios, export
├─ peak_selector.py       # Onglet « Exploration » : sélection pics + PCA corrélée + PCA libre
├─ data_processing.py     # Chargement .txt, baseline modpoly, fusion avec métadonnées (depuis l’UI)
└─ legacy/                # Anciens prototypes / scripts (non utilisés par main.py)
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
pip install PySide6 pandas numpy plotly pybaselines openpyxl scikit-learn scipy
```

> **Note** : `openpyxl` est nécessaire pour les fichiers Excel. `PySide6` peut nécessiter des composants Qt WebEngine supplémentaires selon le système.

---

## Lancement de l’application

```bash
python main.py
```

Une fenêtre **Ramanalyze** s’ouvre avec les onglets : *Présentation*, *Fichiers*, *Métadonnées*, *Spectres*, *Analyse*, *Exploration*.

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

#### Tableau des volumes

Remplissez le tableau réactifs × tubes (valeurs en µL). Les solutions conventionnelles sont :

| Solution | Rôle |
|---|---|
| **A** | Tampon (volume variable, complémentaire à B) |
| **B** | Titrant (**paramètre variable**) |
| **C** | Indicateur |
| **D** | PEG |
| **E** | Nanoparticules |
| **F** | Crosslinker |

Le bouton **Générer volumes (gaussien)** calcule automatiquement les volumes A et B pour une titration centrée sur l’équivalence chimique (distribution gaussienne déformée).

#### Correspondance spectres ↔ tubes

Associez chaque fichier spectre à un tube. Le label `Contrôle` est automatiquement mappé au dernier tube (réplicat du tube final).

#### Assemblage

1. Cliquez **Assembler** :
   - Lecture des `.txt`, correction de baseline (modpoly, ordre 5)
   - Fusion avec les métadonnées sur la clé **`Spectrum name`**
   - Exclusion des contrôles BRB si présents
2. Vérifiez le message de succès (lignes/colonnes du fichier combiné).
3. Cliquez **Enregistrer** → **CSV** ou **Excel** au choix (dernier chemin mémorisé).

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

### 5) Onglet Exploration (PCA)

L’onglet **Exploration** regroupe la sélection de pics et deux modes de PCA accessibles via des sous-onglets.

#### Sous-onglets disponibles

| Sous-onglet | Description |
|---|---|
| Spectres normalisés | Superposition des spectres après normalisation vecteur |
| Spectres superposés | Superposition brute des spectres |
| PCA corrélée | PCA calculée sur les intensités aux pics sélectionnés |
| PCA — Scores | Scores des spectres sur les composantes de la PCA libre |
| PCA — Loadings | Loadings (poids) des composantes sur le Raman Shift |
| PCA — Reconstruction | Reconstruction d’un spectre à partir de N composantes |
| Scatter matrix | Matrice de corrélation entre pics |

#### Paramètres

- **Composantes PCA corrélée** : nombre de PC pour la PCA sur les pics.
- **Composantes PCA libre** : nombre de PC pour la PCA sur les spectres complets (2–15).
- **Colorier par** : variable de métadonnées utilisée pour colorer les points (scores, spectres…).

#### Workflow typique

1. Rechargez les données.
2. Définissez vos pics cibles et la tolérance.
3. Lancez l’analyse.
4. Explorez les sous-onglets **PCA — Scores** et **PCA — Loadings** pour interpréter les résultats (voir section [Comprendre la PCA](#comprendre-la-pca-en-spectroscopie-raman)).
5. Utilisez **PCA — Reconstruction** pour visualiser ce qu’une composante capture réellement.

---

## Comprendre la PCA en spectroscopie Raman

La PCA (Analyse en Composantes Principales) est l'outil central de l'onglet Exploration. Cette section explique ce qu'elle calcule, comment l'interpréter, et ce que signifient concrètement PC1, PC2, PC3… dans le contexte des spectres Raman.

---

### Pourquoi la PCA ?

Chaque spectre Raman est un vecteur de ~1 000 points (un pour chaque nombre d'onde). Si tu as 15 spectres, tu travailles dans un espace à 1 000 dimensions — impossible à visualiser directement. La PCA projette cet espace sur 2 ou 3 axes qui **résument l'essentiel de la variabilité** entre spectres.

---

### Calcul pas à pas

#### Étape 1 — La matrice de données

On empile les spectres en une matrice **X** de taille (n_spectres × n_points) :

```
              cm⁻¹_1  cm⁻¹_2  ...  cm⁻¹_1000
Spectre 1  [  120      45    ...    8      ]
Spectre 2  [  118      47    ...    9      ]
...
Spectre 15 [  130      41    ...    7      ]
```

#### Étape 2 — Centrage

Pour chaque colonne (chaque nombre d'onde), on soustrait la moyenne sur tous les spectres :

```
X_centré = X − moyenne(X, par colonne)
```

Chaque colonne a maintenant une moyenne de 0. On travaille sur les **écarts** par rapport au spectre moyen — c'est ce qui varie d'un spectre à l'autre qui est intéressant.

#### Étape 3 — Matrice de covariance

On calcule **C = Xᵀ · X / (n−1)**, une matrice 1 000 × 1 000.

Chaque cellule C[i, j] mesure : *"est-ce que quand l'intensité à cm⁻¹_i est élevée, celle à cm⁻¹_j l'est aussi ?"*. Elle encode toutes les corrélations entre bandes Raman.

#### Étape 4 — Décomposition en valeurs propres

On résout l'équation :

```
C · v = λ · v
```

Cela donne 1 000 paires (λ, v) — des **valeurs propres** et des **vecteurs propres**.

- **v** (vecteur de 1 000 coefficients) → c'est le **loading** d'une PC
- **λ** (un scalaire positif) → c'est la variance expliquée par cette PC

On les trie par λ décroissant. Le plus grand λ → PC1, le 2ᵉ → PC2, etc.

#### Étape 5 — Les scores

Le score d'un spectre sur PC1 est un produit scalaire :

```
score_PC1[spectre_i] = X_centré[spectre_i, :] · loading_PC1
```

On "projette" le spectre sur la direction PC1. Résultat : **un seul nombre** par spectre.

---

### Ce que signifient PC1, PC2, PC3…

#### PC1

La direction dans l'espace des spectres qui **explique le plus de variance**. Ce n'est pas forcément la plus chimiquement intéressante — elle peut capturer un artefact (ligne de base fluorescente, variation d'intensité laser…) si c'est la plus grande source de variabilité.

#### PC2

Perpendiculaire à PC1, elle capture la **deuxième source de variabilité**. C'est souvent là qu'on trouve l'effet chimique recherché si PC1 est dominée par un artefact.

#### PC3 et suivantes

Capturent des effets de plus en plus mineurs. Au-delà de PC3–PC5, c'est généralement du bruit de mesure.

> En pratique, dans une titration SERS, si la concentration en Cu²⁺ est la 2ᵉ source de variabilité après la ligne de base, l'effet de la concentration apparaîtra sur **PC2**, pas sur PC1.

---

### Comment interpréter un loading

Le **loading** de PC1 est lui-même un "spectre fantôme" : il a la même longueur que tes spectres (un coefficient par nombre d'onde).

```
Loading PC1 = [0.02, -0.01, 0.0, 0.15, ..., 0.08]
               cm⁻¹_1  cm⁻¹_2       cm⁻¹_4        cm⁻¹_1000
```

- Un coefficient **grand positif** à 1 200 cm⁻¹ → PC1 est très sensible à la bande à 1 200 cm⁻¹
- Un coefficient **proche de zéro** à 800 cm⁻¹ → PC1 ignore cette région
- Un coefficient **négatif** → les spectres avec un **score PC1 négatif** ont une intensité élevée là

**En pratique** : trace le loading comme un spectre et identifie à quels nombres d'onde il a des pics, exactement comme tu le ferais pour un spectre expérimental.

---

### Comment interpréter un score

Le **score** d'un spectre sur PC1 est sa "coordonnée" sur cet axe :

```
Scores PC1 = [-3.2, -2.8, -1.1, 0.0, 0.5, 1.2, ..., 4.1]
               S1    S2    S3   S4   S5   S6       S15
```

- Score **très positif** → ce spectre ressemble fortement au pattern du loading PC1
- Score **proche de 0** → ce spectre est neutre par rapport à PC1
- Score **très négatif** → ce spectre est opposé au pattern de PC1

---

### Graphe Scores PC1 vs PC2 : la clé de l'interprétation

Le graphe des scores, coloré par une variable de métadonnées (ex : concentration en titrant), répond à la question :

> *"Est-ce que ma variable chimique explique la séparation entre spectres ?"*

- Points qui **se séparent le long de PC1** selon la concentration → **PC1 capture l'effet du titrant**
- Points **regroupés aléatoirement** → la variabilité capturée n'est pas liée à ta variable d'intérêt
- Points qui **se séparent en PC2** → le signal chimique est la 2ᵉ source de variance (PC1 = artefact)

```
         PC2
          ↑   · · (faible conc.)
          |  · · ·
          | · · · ·
          |· · ·
          +----------→ PC1    (forte conc.)
```

Chaque point = un spectre. La distance sur un axe = différence chimique/physique capturée par cet axe.

---

### Reconstruction et variance expliquée

L'onglet **PCA — Reconstruction** permet de reconstruire un spectre à partir de N composantes seulement :

```
spectre_reconstruit = spectre_moyen + score_PC1 × loading_PC1
                                    + score_PC2 × loading_PC2
                                    + ...
                                    + score_PCn × loading_PCn
```

- Avec **N = 1** : on voit seulement ce que PC1 a capturé
- Avec **N = 3** : on voit les 3 premières sources de variabilité
- Avec **N = toutes** : on retrouve le spectre original

La **variance expliquée** (ex : 72 % pour PC1) dit que PC1 résume 72 % de toute la variabilité entre spectres. Les PC suivantes se partagent les 28 % restants.

---

### Résumé pratique

| Ce qu'on regarde | Ce que ça dit |
|---|---|
| **Loading PC_n** tracé comme un spectre | Quelles bandes Raman "portent" cette composante |
| **Score PC_n** d'un spectre | À quel point ce spectre ressemble au pattern de PC_n |
| **Graphe scores PC1 vs PC2**, coloré par concentration | Est-ce que la PCA sépare les échantillons selon la chimie |
| **Variance expliquée** (ex : 72 %) | PC1 résume 72 % de toute la variabilité entre spectres |
| **Reconstruction avec N composantes** | Ce que les N premières PC ont réellement capturé |

---

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
- `build_combined_dataframe_from_df(txt_files, metadata_df, poly_order=5, exclude_brb=True)` :
  - concatène les spectres valides, ajoute `Spectrum name`
  - fusion gauche sur `Spectrum name` avec les métadonnées (DataFrame)
  - retire les contrôles BRB si `exclude_brb=True`
- `build_combined_dataframe_from_ui(txt_files, metadata_creator, poly_order=5, exclude_brb=True)` :
  - construit les métadonnées via l’onglet Métadonnées (`metadata_creator`)
  - appelle `build_combined_dataframe_from_df`

### `file_picker.py`
- Navigation dans l’arborescence, filtre `.txt`, sélection multiple, liste de fichiers choisis.
- Méthodes :
  - `get_selected_files()` → `list[str]` des chemins `.txt` sélectionnés.

### `metadata_creator.py`
- Création des métadonnées directement dans l’app :
  - tableau des volumes (réactifs × tubes)
  - correspondance spectres ↔ tubes
  - génération de volumes (distribution gaussienne)
- Sauvegarde / chargement des métadonnées (Excel `.xlsx`).
- L’ancien sélecteur de fichiers de métadonnées est conservé dans `legacy/metadata_picker.py`.

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

### `peak_selector.py`
- Onglet **Exploration** : normalisation des spectres, sélection de pics, scatter matrix.
- **PCA corrélée** : PCA sur les intensités aux pics sélectionnés uniquement.
- **PCA libre** : PCA sur les spectres complets (matrice normalisée).
  - Sous-onglet **Scores** : position de chaque spectre dans l’espace PCA.
  - Sous-onglet **Loadings** : poids de chaque nombre d’onde pour chaque composante.
  - Sous-onglet **Reconstruction** : reconstruction d’un spectre à partir de N composantes.

### `metadata_model.py`
- Couche **données pure** (aucun import Qt) — testable sans interface graphique.
- Contient : `to_float`, `conc_to_M`, `normalize_tube_label`, `norm_text_key`, `is_control_label`, `get_tube_columns`, `infer_n_tubes_for_mapping`, `tube_merge_key_for_mapping`, `compute_concentration_table`, `build_merged_metadata`, `sers_gaussian_volumes`.

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

**Q : L’assemblage échoue.**
R : Vérifiez que :
- la colonne **`Spectrum name`** est présente dans les métadonnées et correspond exactement aux noms de fichiers `.txt` (sans extension, casse exacte) ;
- le séparateur CSV est standard (`,` `;` `\t`) — la détection est automatique ;
- l’encodage est UTF-8.

**Q : Mon fichier combiné n’a pas `Intensity_corrected`.**
R : Les `.txt` doivent contenir `Raman Shift` et `Dark Subtracted #1` après la ligne `Pixel;`. Vérifiez le séparateur (`;`) et les décimales (`,`).

**Q : Le tracé n’affiche rien dans l’onglet Spectres.**
R : Vérifiez que vous avez assemblé le fichier combiné, que les fichiers sélectionnés (onglet *Fichiers*) existent dans la colonne `file` du combiné, et que les colonnes `Raman Shift` et `Intensity_corrected` sont présentes.

**Q : L’analyse indique des colonnes manquantes.**
R : Le combiné doit contenir au minimum `Raman Shift`, `Intensity_corrected`, `file`. Pour le graphique des ratios, la variable de métadonnées choisie (ex. `n(EGTA) (mol)`) doit exister.

**Q : La PCA n’affiche rien ou plante.**
R : Il faut au moins 2 spectres. Le nombre de composantes demandé ne doit pas dépasser le nombre de spectres (ni le nombre de pics pour la PCA corrélée). Vérifiez que l’assemblage a réussi.

**Q : Les spectres Contrôle n’apparaissent pas dans les analyses.**
R : Le label `Contrôle` dans la correspondance spectres ↔ tubes doit être orthographié avec majuscule initiale. Il est automatiquement mappé au dernier tube du tableau des volumes.

**Q : Des doublons de Raman Shift sont signalés dans la console.**
R : Ce n’est pas bloquant — l’application moyenne automatiquement les points dupliqués. Vérifiez néanmoins la qualité de votre fichier `.txt` source.

**Q : La PCA libre et la PCA corrélée donnent des résultats très différents.**
R : C’est normal et attendu. La PCA corrélée n’utilise que les intensités aux pics choisis (5 variables) ; la PCA libre utilise tout le spectre (~1 000 variables). Si un effet chimique important se manifeste en dehors des pics sélectionnés, seule la PCA libre le détectera.

---

## Bonnes pratiques & conseils

- Vérifiez que `Spectrum name` dans les métadonnées correspond **exactement** aux noms de fichiers `.txt` (sans extension, casse exacte).
- Utilisez des noms d'échantillons clairs dans `Sample description` → meilleures légendes dans les graphiques.
- Travaillez en **encodage UTF-8** pour les fichiers CSV.
- Vérifiez la correction de baseline dans l'onglet Spectres **avant** toute analyse quantitative ou PCA.
- En PCA, regardez **plusieurs composantes** (PC1, PC2, PC3) : l'effet chimique n'est pas toujours sur PC1.
- Utilisez la **Reconstruction** (PCA libre) avec N=1, 2, 3 composantes pour comprendre intuitivement ce que chaque axe capture.
- Après modification des fichiers ou des métadonnées, pensez à **recharger** le fichier combiné dans les onglets Analyse et Exploration.

---

## Crédits

- Développement : **Alexandre Souchaud** — CitizenSers
- Correction de baseline : [`pybaselines`](https://github.com/derb12/pybaselines)
- Interface graphique : **PySide6 / Qt**
- Graphiques interactifs : **Plotly**
- PCA : **scikit-learn** (`sklearn.decomposition.PCA`)

---

## Licence

Projet académique / interne — CitizenSers. Tous droits réservés.
