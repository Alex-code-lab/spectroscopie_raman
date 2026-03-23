# Ramanalyze — Application d'analyse de spectres Raman

> **PySide6 · pandas · plotly · pybaselines · scikit-learn**
> Outil graphique pour assembler, visualiser et analyser des spectres Raman à partir de fichiers `.txt` et de métadonnées expérimentales.

---

## Table des matières

- [Aperçu](#aperçu)
- [Structure du dépôt](#structure-du-dépôt)
- [Installation](#installation)
- [Lancement](#lancement)
- [Guide pas à pas](#guide-pas-à-pas)
  - [Onglet Métadonnées](#onglet-métadonnées)
  - [Onglet Fichiers Raman](#onglet-fichiers-raman)
  - [Onglet Spectres](#onglet-spectres)
  - [Onglet Analyse](#onglet-analyse)
  - [Onglet Exploration (PCA)](#onglet-exploration-pca)
- [Comprendre la PCA en spectroscopie Raman](#comprendre-la-pca-en-spectroscopie-raman)
- [Formats de fichiers attendus](#formats-de-fichiers-attendus)
- [Détails techniques (modules)](#détails-techniques-modules)
- [Export et sauvegardes](#export-et-sauvegardes)
- [Bonnes pratiques](#bonnes-pratiques)
- [Dépannage](#dépannage)
- [Crédits](#crédits)

---

## Aperçu

**Ramanalyze** est une application de bureau développée sous Qt (PySide6). Elle guide l'utilisateur à travers un flux de travail structuré en six étapes :

1. Saisie des **métadonnées expérimentales** (volumes, concentrations, correspondance tubes).
2. Sélection des **fichiers spectres** `.txt`.
3. Visualisation et contrôle qualité des **spectres corrigés**.
4. **Analyse quantitative** des pics et calcul des rapports d'intensité.
5. **Exploration multivariée** par PCA.
6. **Export** des résultats et des figures.

---

## Structure du dépôt

```
spectroscopie_raman/
├── main.py               Point d'entrée de l'application (fenêtre principale, onglets, page Présentation)
├── file_picker.py         Onglet "Fichiers Raman" : navigation arborescente et sélection de .txt
├── metadata_creator.py    Onglet "Métadonnées" : tableau des volumes, génération gaussienne, protocole de paillasse
├── metadata_model.py      Couche données pure (sans Qt) : calculs concentrations, génération de volumes, fusion
├── spectra_plot.py        Onglet "Spectres" : tracé interactif Plotly et export haute résolution
├── analysis_tab.py        Onglet "Analyse" : intensités, ratios, fit sigmoïde, export
├── peak_selector.py       Onglet "Exploration" : sélection de pics, PCA corrélée, PCA libre
├── data_processing.py     Chargement .txt, correction baseline modpoly, fusion avec métadonnées
├── plotly_downloads.py    Utilitaires d'export graphique (JS Plotly, printToPdf)
└── legacy/                Anciens scripts et prototypes (non utilisés par main.py)
```

---

## Installation

### Prérequis

- Python 3.9 ou supérieur
- macOS, Linux ou Windows

### Environnement virtuel et dépendances

```bash
# Créer et activer un environnement virtuel (recommandé)
python -m venv .venv
source .venv/bin/activate        # macOS / Linux
# .venv\Scripts\activate         # Windows

# Installer les dépendances
pip install PySide6 pandas numpy plotly pybaselines openpyxl scikit-learn scipy
```

`openpyxl` est nécessaire pour la lecture et l'écriture des fichiers Excel. PySide6 inclut QtWebEngine, qui est utilisé pour le rendu des graphiques Plotly interactifs.

---

## Lancement

```bash
python main.py
```

La fenêtre **Ramanalyze** s'ouvre avec les onglets : *Présentation*, *Métadonnées*, *Fichiers Raman*, *Spectres*, *Analyse*, *Exploration*.

L'onglet **Présentation** reprend un guide d'utilisation synthétique accessible à tout moment depuis l'application.

---

## Guide pas à pas

### Onglet Métadonnées

C'est le point de départ de toute session. Cet onglet centralise la saisie des informations expérimentales et la gestion du protocole de paillasse.

#### Tableau des volumes

Le tableau représente la composition de chaque tube de l'expérience. Les lignes correspondent aux réactifs (Solutions A à F), les colonnes aux tubes (Tube 1, Tube 2, etc.). Toutes les valeurs sont en µL.

| Solution | Rôle |
|---|---|
| **Solution A** | Tampon (volume variable, complémentaire à B pour atteindre le volume total) |
| **Solution B** | Titrant (paramètre variable de l'expérience) |
| **Solution C** | Indicateur |
| **Solution D** | PEG |
| **Solution E** | Nanoparticules |
| **Solution F** | Crosslinker |

**Gestion des colonnes (tubes).** Un clic sur l'en-tête d'une colonne la sélectionne visuellement. Un clic droit sur l'en-tête ouvre un menu contextuel proposant de **dupliquer** ou **supprimer** ce tube. Des boutons dédiés dans la barre d'outils permettent également d'ajouter un tube vierge en fin de tableau.

#### Génération automatique des volumes (distribution gaussienne)

Le bouton **Générer volumes (gaussien)** ouvre un dialogue de configuration qui calcule automatiquement les volumes de Solutions A et B pour une titration centrée sur l'équivalence chimique.

**Paramètres principaux :**

| Paramètre | Description |
|---|---|
| Nombre de tubes | Nombre de tubes à générer |
| C₀ (nM) | Concentration initiale de l'analyte (Cu) dans l'échantillon |
| V échantillon (µL) | Volume total de l'échantillon |
| C titrant (µM) | Concentration de la solution titrante (Solution B) |
| V total cuvette (µL) | Volume final dans chaque tube après pipetage |

**Solutions fixes (C, D, E, F).** Pour chaque solution fixe, on renseigne indépendamment le volume ajouté (µL) et la concentration stock (avec unité : M, mM, µM, nM). Ces deux valeurs sont entièrement libres et sans couplage automatique entre elles.

**Mode compte-goutte.** Cocher cette option applique un preset adapté au pipetage par compte-goutte : marge à 0 µL, pas pipette à 40 µL, volumes des Solutions C et F remis à 40 µL. Après activation du preset, tous les champs restent modifiables librement.

**Options de génération :**

- *Inclure un tube à 0 µL de titrant* : ajoute un tube sans titrant en début de série.
- *Dupliquer le dernier volume pour contrôle* : répète le tube de fin pour évaluer la reproductibilité.

**Bouton "Quantités de matière…".** Ouvre un dialogue récapitulatif des quantités de matière (en nmol) dans chaque tube, calculées à partir des concentrations stock et des volumes. Les valeurs issues de concentrations molaires sont éditables directement dans ce dialogue pour un ajustement fin.

**Bouton "Valeurs par défaut".** Remet tous les paramètres du dialogue aux valeurs d'usine (utile pour repartir d'une base propre sans fermer le dialogue).

#### Correspondance spectres ↔ tubes

Cette section associe chaque fichier spectre `.txt` à un tube du tableau. Cette correspondance est indispensable pour que les concentrations calculées soient correctement rattachées à chaque spectre lors de l'assemblage.

Le label `Contrôle` est automatiquement mappé au dernier tube (réplicat du tube final).

#### Feuille de protocole de paillasse

Le bouton **Feuille de protocole…** ouvre un dialogue interactif destiné au suivi du pipetage sur la paillasse. Il affiche le tableau des volumes sous forme d'une grille où chaque ligne est un réactif et chaque colonne un tube.

Pour chaque tube, trois sous-colonnes sont proposées :

- **Volume (µL)** : lecture seule, rappelle la quantité à pipeter.
- **Coordinateur** : case à cocher validée par le coordinateur après préparation.
- **Opérateur** : case à cocher validée par l'opérateur après vérification indépendante.

Une colonne **Vérification solution** permet de confirmer l'identité du flacon avant de commencer le pipetage d'un réactif.

**Mode guidé (étape par étape).** En mode guidé, seule la cellule de l'étape en cours est active ; toutes les autres sont grisées. La séquence automatique progresse ainsi : vérification du flacon, puis tube 1 (confirmation volume, coordinateur, opérateur), tube 2, etc., avant de passer au réactif suivant. Un indicateur vert affiche en permanence l'étape courante.

Les boutons **Tout cocher — Coord.** et **Tout cocher — Opér.** permettent de valider toutes les cases d'un rôle en un seul clic. **Tout décocher** remet tout à zéro.

**Export et rechargement.** Le bouton **Exporter Excel…** génère un fichier `.xlsx` contenant quatre feuilles :

- `Protocole` : feuille visuelle mise en forme (en-têtes colorés, texte pivoté, symboles ✓/☐).
- `Volumes` : données brutes du tableau des volumes.
- `Correspondance` : données brutes de la table spectres ↔ tubes.
- `EtatProtocole` : état de chaque case (type, indices, coché ou non).

Ce fichier peut être rechargé via **Charger des métadonnées…** : toutes les cases reprennent exactement leur état sauvegardé. L'état est également préservé en mémoire entre les ouvertures du dialogue.

#### Concentrations en cuvette

Le bouton **Voir concentrations** affiche le tableau des concentrations finales dans chaque tube, calculées à partir des volumes et des concentrations stock renseignées. Ce tableau est en lecture seule.

#### Assemblage et sauvegarde

Une fois le tableau des volumes et la correspondance remplis :

1. Cliquez **Assembler** : les fichiers `.txt` sont lus, la baseline est corrigée (modpoly, ordre 5) et les métadonnées sont fusionnées sur la clé `Spectrum name`.
2. Vérifiez le message de succès (nombre de lignes et de colonnes du fichier combiné).
3. Cliquez **Enregistrer** pour sauvegarder le fichier combiné en CSV ou Excel au choix. Le dernier chemin utilisé est mémorisé.

---

### Onglet Fichiers Raman

Cet onglet permet de constituer la liste des spectres à traiter.

- Double-cliquez sur un dossier pour le parcourir.
- Sélectionnez un ou plusieurs fichiers `.txt` (résultats bruts du spectromètre Raman).
- Cliquez **Ajouter la sélection** pour les placer dans la liste de droite.
- Utilisez **Retirer** ou **Vider la liste** si nécessaire.

Seuls les fichiers `.txt` sont affichés dans le panneau de sélection ; les dossiers restent navigables.

---

### Onglet Spectres

Cet onglet affiche les spectres corrigés issus du fichier combiné produit dans l'onglet Métadonnées.

- La légende utilise `Sample description` si cette colonne est présente dans les métadonnées, sinon le nom de fichier.
- L'axe X affiche le Raman Shift (cm⁻¹), l'axe Y l'intensité corrigée (unités arbitraires).
- Cliquez **Tracer avec baseline** pour générer la figure Plotly interactive (zoom, pan, légende cliquable).

**Export du graphique.** Le bouton **Exporter le graphique…** ouvre un dialogue proposant :

- **Format de fichier** : PNG (haute résolution ×2), SVG (vectoriel), PDF (via le moteur Qt).
- **Taille de sortie** (PNG et SVG) : presets A4 paysage, A4 portrait, A3 paysage, A3 portrait, écran large (1920×1080), carré HD (1400×1400), ou dimensions personnalisées en pixels.

L'export PNG et SVG utilise l'API JavaScript de Plotly directement depuis le moteur de rendu de l'application : le résultat est pixel-perfect et identique à ce qui est affiché à l'écran. L'export PDF utilise le moteur d'impression Qt, en format A4 paysage.

---

### Onglet Analyse

Cet onglet calcule les intensités aux pics d'intérêt et les rapports d'intensité (ratios), puis trace leur évolution en fonction de la quantité de titrant.

#### Flux de travail

1. Cliquez **Recharger le fichier combiné depuis Métadonnées** si les données viennent d'être mises à jour.
2. Choisissez un **jeu de pics** dans la liste déroulante :
   - *532 nm — paires pré-enregistrées* : paires de pics pertinentes pour le laser 532 nm.
   - *785 nm* : pics à 412, 444, 471, 547 et 1561 cm⁻¹.
   - *Personnalisé* : saisissez manuellement les pics à étudier (Raman Shift en cm⁻¹, séparés par des virgules).
3. Réglez la **tolérance** (fenêtre de recherche autour de chaque pic, 5 cm⁻¹ par défaut).
4. Cliquez **Analyser les pics** : l'application recherche l'intensité maximale dans chaque fenêtre [pic ± tolérance] pour chaque spectre, calcule tous les ratios entre paires de pics, et trace un graphique interactif des ratios en fonction de la quantité de titrant (n(titrant) en mol).

#### Tableau des intensités

Le bouton **Afficher le tableau des intensités** ouvre un tableau récapitulatif des intensités extraites par pic et par spectre. Il est triable par colonne.

#### Fit sigmoïde et équivalence

Pour chaque ratio, il est possible d'ajuster une fonction sigmoïde afin de déterminer le point d'équivalence :

1. Sélectionnez le ratio à ajuster dans la liste **Ratio à ajuster**.
2. Cliquez **Ajuster par sigmoïde** : le fit est calculé par moindres carrés (scipy.optimize.curve_fit) et la courbe ajustée est superposée au graphique.
3. Cliquez **Afficher l'équivalence** pour afficher sur le graphique le point d'équivalence extrait du fit, accompagné de sa valeur numérique.

Les paramètres du fit (A, B, k, x_eq) sont affichés dans la barre de statut de l'application.

#### Export des résultats

Le bouton **Exporter résultats (Excel)…** génère un fichier `.xlsx` avec deux feuilles :

- `intensites` : intensités maximales par pic et par spectre.
- `ratios` : ratios au format long (spectre, ratio, valeur, quantité de titrant).

**Export du graphique.** Identique à l'onglet Spectres : PNG haute résolution, SVG vectoriel ou PDF, avec choix du format de sortie (A4, A3, personnalisé).

---

### Onglet Exploration (PCA)

L'onglet Exploration regroupe la visualisation des spectres normalisés, la sélection interactive de pics et deux modes d'analyse multivariée par PCA.

#### Paramètres communs

- **Pics à étudier** : liste de Raman Shifts (cm⁻¹) séparés par des virgules.
- **Tolérance (cm⁻¹)** : fenêtre de recherche autour de chaque pic.
- **Composantes PCA corrélée** : nombre de composantes pour la PCA sur les pics sélectionnés.
- **Composantes PCA libre** : nombre de composantes pour la PCA sur les spectres complets (entre 2 et 15).
- **Colorier par** : variable de métadonnées utilisée pour colorer les points dans les graphiques de scores et de spectres.

#### Sous-onglets disponibles

| Sous-onglet | Ce qu'il montre |
|---|---|
| Spectres normalisés | Superposition des spectres après normalisation par la norme vectorielle |
| Spectres superposés | Superposition brute des spectres sans normalisation |
| PCA corrélée | PCA calculée uniquement sur les intensités aux pics sélectionnés |
| PCA — Scores | Position de chaque spectre dans l'espace des composantes principales (PCA libre) |
| PCA — Loadings | Poids de chaque nombre d'onde pour chaque composante principale |
| PCA — Reconstruction | Reconstruction d'un spectre à partir de N composantes principales |
| Scatter matrix | Matrice de corrélations visuelles entre tous les pics sélectionnés |

#### Export des paires de pics (CSV)

Le bouton **Exporter tableau CSV** enregistre toutes les paires de pics et leurs ratios dans un fichier `.csv` directement importable dans Excel ou R.

---

## Comprendre la PCA en spectroscopie Raman

### Pourquoi la PCA ?

Chaque spectre Raman est un vecteur de l'ordre de 1 000 points (un par nombre d'onde). Avec 15 spectres, on travaille dans un espace à 1 000 dimensions — impossible à visualiser directement. La PCA projette cet espace sur 2 ou 3 axes qui résument l'essentiel de la variabilité entre spectres.

### Calcul pas à pas

**Étape 1 — La matrice de données**

Les spectres sont empilés en une matrice **X** de taille (n_spectres × n_points) :

```
              cm⁻¹_1  cm⁻¹_2  ...  cm⁻¹_1000
Spectre 1  [  120      45    ...    8      ]
Spectre 2  [  118      47    ...    9      ]
...
Spectre 15 [  130      41    ...    7      ]
```

**Étape 2 — Centrage**

Pour chaque colonne (chaque nombre d'onde), on soustrait la moyenne sur tous les spectres :

```
X_centré = X − moyenne(X, par colonne)
```

Chaque colonne a maintenant une moyenne de 0. On travaille sur les écarts par rapport au spectre moyen.

**Étape 3 — Matrice de covariance**

On calcule **C = Xᵀ · X / (n−1)**, une matrice 1 000 × 1 000. Chaque cellule C[i, j] mesure si l'intensité à cm⁻¹_i et celle à cm⁻¹_j covarient entre spectres.

**Étape 4 — Décomposition en valeurs propres**

On résout l'équation **C · v = λ · v**, ce qui donne 1 000 paires (λ, v) :

- **v** (vecteur de 1 000 coefficients) est le **loading** d'une composante.
- **λ** (un scalaire positif) est la variance expliquée par cette composante.

On trie par λ décroissant : le plus grand λ donne PC1, le suivant PC2, etc.

**Étape 5 — Les scores**

Le score d'un spectre sur PC1 est le produit scalaire de ce spectre centré avec le loading de PC1 :

```
score_PC1[spectre_i] = X_centré[spectre_i, :] · loading_PC1
```

Le résultat est un seul nombre par spectre, sa "coordonnée" sur l'axe PC1.

### Interprétation des composantes

**PC1** est la direction dans l'espace des spectres qui explique le plus de variance. Ce n'est pas nécessairement la plus chimiquement significative : elle peut capturer un artefact (variation de ligne de base, fluctuation de puissance laser) si c'est la source dominante de variabilité.

**PC2** est perpendiculaire à PC1 et capture la deuxième source de variabilité. Dans une titration SERS, l'effet chimique apparaît souvent sur PC2 si PC1 est dominée par un artefact instrumental.

**PC3 et suivantes** capturent des effets de plus en plus mineurs. Au-delà de PC3 à PC5, on tombe généralement dans le bruit de mesure.

### Interpréter un loading

Le loading de PC1 est un "spectre fantôme" de même longueur que les spectres expérimentaux. Un coefficient grand positif à 1 200 cm⁻¹ indique que PC1 est très sensible à la bande Raman à 1 200 cm⁻¹. Un coefficient négatif signifie que les spectres avec un score PC1 négatif présentent une intensité élevée à ce nombre d'onde. En pratique, on trace le loading comme un spectre ordinaire et on identifie ses pics par comparaison avec les spectres de référence.

### Interpréter un score

Un score très positif sur PC1 indique que le spectre ressemble fortement au pattern décrit par le loading PC1. Un score proche de 0 signifie que le spectre est neutre vis-à-vis de ce pattern. Un score négatif indique une opposition au pattern.

### Graphe Scores PC1 vs PC2

Ce graphe, coloré par une variable de métadonnées (par exemple la concentration en titrant), répond à la question : est-ce que ma variable chimique explique la séparation entre spectres ?

- Les points se séparent le long de PC1 selon la concentration : PC1 capture l'effet du titrant.
- Les points se séparent selon PC2 : le signal chimique est la deuxième source de variance (PC1 est un artefact).
- Les points sont regroupés sans structure : la variabilité capturée n'est pas liée à la variable d'intérêt.

### Reconstruction et variance expliquée

L'onglet **PCA — Reconstruction** permet de reconstruire un spectre à partir de N composantes seulement :

```
spectre_reconstruit = spectre_moyen
                    + score_PC1 × loading_PC1
                    + score_PC2 × loading_PC2
                    + ...
                    + score_PCn × loading_PCn
```

Avec N = 1, on visualise uniquement ce que PC1 a capturé. En ajoutant progressivement des composantes, on comprend intuitivement la contribution de chacune.

La variance expliquée (par exemple 72 % pour PC1) signifie que PC1 résume à elle seule 72 % de toute la variabilité entre spectres.

### Résumé pratique

| Ce qu'on regarde | Ce que ça dit |
|---|---|
| Loading PC_n tracé comme un spectre | Quelles bandes Raman portent cette composante |
| Score PC_n d'un spectre | À quel point ce spectre ressemble au pattern de la composante |
| Graphe scores PC1 vs PC2 coloré par concentration | Est-ce que la PCA sépare les échantillons selon la chimie |
| Variance expliquée (ex. 72 %) | PC1 résume 72 % de toute la variabilité entre spectres |
| Reconstruction avec N composantes | Ce que les N premières composantes ont réellement capturé |

---

## Formats de fichiers attendus

### Fichiers spectres `.txt`

Les fichiers doivent contenir un en-tête avec une ligne commençant par `Pixel;` (séparateur `;`, décimales `,`). Les colonnes nécessaires après lecture sont :

- `Raman Shift`
- `Dark Subtracted #1`

Le pipeline convertit les données en numérique, trie sur `Raman Shift` et calcule :

- `Intensity_corrected` = `Dark Subtracted #1` − baseline (modpoly, ordre 5)
- `file` = nom du fichier `.txt`
- `Spectrum name` = `file` sans l'extension `.txt`

### Fichier métadonnées (Excel ou CSV)

- Excel : `.xlsx` ou `.xls` (lecture via openpyxl).
- CSV : `.csv` (séparateur auto-détecté : `,`, `;` ou tabulation).

La fusion entre spectres et métadonnées se fait sur la colonne `Spectrum name` (clé de jointure). Colonnes utiles : `Spectrum name`, `Sample description`, et toutes les colonnes de concentrations générées par le tableau des volumes.

---

## Détails techniques (modules)

### `data_processing.py`

- `load_spectrum_file(path, poly_order=5)` : lit un fichier `.txt`, corrige la baseline par modpoly et retourne un DataFrame avec `Raman Shift`, `Dark Subtracted #1`, `Intensity_corrected`, `file`.
- `build_combined_dataframe_from_df(txt_files, metadata_df, ...)` : concatène les spectres valides et fusionne avec un DataFrame de métadonnées fourni en paramètre.
- `build_combined_dataframe_from_ui(txt_files, metadata_creator, ...)` : construit les métadonnées depuis l'onglet Métadonnées puis appelle `build_combined_dataframe_from_df`.

### `file_picker.py`

Navigation arborescente filtrée sur les `.txt`, sélection multiple, liste des fichiers choisis. Méthode principale : `get_selected_files()` renvoie la liste des chemins absolus sélectionnés.

### `metadata_creator.py`

Onglet Métadonnées complet. Gère le tableau des volumes (ajout, suppression et duplication de colonnes via menu contextuel), le dialogue de génération gaussienne des volumes (avec presets normal et compte-goutte, paramètres de solutions fixes libres), la correspondance spectres ↔ tubes, la feuille de protocole de paillasse (mode libre et mode guidé), les concentrations en cuvette et l'export/rechargement Excel.

### `metadata_model.py`

Couche données pure, sans aucun import Qt. Testable indépendamment. Contient les fonctions de calcul des concentrations, la normalisation des labels de tubes, la génération de volumes par distribution gaussienne (`sers_gaussian_volumes`), la fusion des métadonnées et les utilitaires de conversion d'unités.

### `spectra_plot.py`

Onglet Spectres. Trace les spectres corrigés depuis le fichier combiné. Propose un export haute résolution via l'API JavaScript de Plotly (PNG ×2, SVG vectoriel) et via le moteur Qt (PDF A4 paysage), avec préréglages de taille (A4, A3, écran large, personnalisé).

### `analysis_tab.py`

Onglet Analyse. Recharge le fichier combiné, propose des jeux de pics prédéfinis (532 nm, 785 nm) et un mode personnalisé. Calcule les intensités maximales par fenêtre, tous les ratios entre paires de pics, et trace le graphique interactif (ratios vs quantité de titrant). Propose un ajustement sigmoïde par ratio avec affichage du point d'équivalence. Export Excel (feuilles `intensites` et `ratios`) et export graphique haute résolution.

### `peak_selector.py`

Onglet Exploration. Visualisation des spectres normalisés et superposés, sélection interactive de pics, scatter matrix des corrélations inter-pics. PCA corrélée sur les intensités aux pics sélectionnés. PCA libre sur les spectres complets avec sous-onglets Scores, Loadings, Reconstruction et barre de variance expliquée. Export CSV des paires de pics.

### `plotly_downloads.py`

Utilitaires partagés pour l'export graphique. L'export PNG et SVG appelle `Plotly.toImage` via `QWebEnginePage.runJavaScript` (moteur identique à l'affichage, pas de Kaleido). L'export PDF utilise `QWebEnginePage.printToPdf`. Un mécanisme de polling asynchrone (`QTimer`) attend la résolution de la promesse JavaScript avant d'écrire le fichier.

---

## Export et sauvegardes

| Type de fichier | Format(s) disponible(s) | Déclencheur |
|---|---|---|
| Fichier combiné spectres + métadonnées | CSV ou Excel (.xlsx) au choix | Bouton "Enregistrer" dans l'onglet Métadonnées |
| Métadonnées et protocole | Excel (.xlsx), 4 feuilles | Bouton "Exporter Excel…" dans le protocole de paillasse |
| Résultats d'analyse (intensités + ratios) | Excel (.xlsx), 2 feuilles | Bouton "Exporter résultats (Excel)…" dans l'onglet Analyse |
| Graphique des spectres | PNG (×2), SVG (vectoriel), PDF | Bouton "Exporter le graphique…" dans l'onglet Spectres |
| Graphique des ratios | PNG (×2), SVG (vectoriel), PDF | Bouton "Enregistrer le graphique" dans l'onglet Analyse |
| Paires de pics | CSV | Bouton "Exporter tableau CSV" dans l'onglet Exploration |

---

## Bonnes pratiques

- Vérifiez que la colonne `Spectrum name` dans les métadonnées correspond exactement aux noms de fichiers `.txt` (sans extension, casse identique).
- Utilisez des noms descriptifs dans `Sample description` pour obtenir des légendes lisibles dans les graphiques.
- Travaillez en encodage UTF-8 pour les fichiers CSV.
- Contrôlez visuellement la correction de baseline dans l'onglet Spectres avant toute analyse quantitative ou PCA. Une baseline mal corrigée peut dominer PC1 et masquer complètement l'effet chimique d'intérêt.
- En PCA, examinez toujours plusieurs composantes (PC1, PC2, PC3) : l'effet chimique n'est pas systématiquement sur PC1.
- Utilisez la Reconstruction (PCA libre) avec N = 1, 2, 3 composantes pour comprendre intuitivement ce que chaque axe capture réellement.
- Après toute modification des fichiers ou des métadonnées, rechargez le fichier combiné dans les onglets Analyse et Exploration.

---

## Dépannage

**L'assemblage échoue.**
Vérifiez que la colonne `Spectrum name` est présente dans les métadonnées et correspond exactement aux noms de fichiers `.txt` (sans extension, casse exacte). Vérifiez que le séparateur CSV est standard (`,`, `;` ou tabulation) et que l'encodage est UTF-8.

**Mon fichier combiné n'a pas la colonne `Intensity_corrected`.**
Les fichiers `.txt` doivent contenir `Raman Shift` et `Dark Subtracted #1` après la ligne `Pixel;`. Vérifiez que le séparateur est `;` et que les décimales sont des virgules.

**Le tracé n'affiche rien dans l'onglet Spectres.**
Vérifiez que vous avez assemblé et enregistré le fichier combiné, que les fichiers sélectionnés dans l'onglet Fichiers Raman sont bien présents dans la colonne `file` du combiné, et que les colonnes `Raman Shift` et `Intensity_corrected` existent.

**L'analyse indique des colonnes manquantes.**
Le fichier combiné doit contenir au minimum `Raman Shift`, `Intensity_corrected` et `file`. Pour le graphique des ratios, la colonne de concentration du titrant (issue du tableau des volumes) doit être présente.

**La PCA n'affiche rien ou génère une erreur.**
Il faut au moins 2 spectres. Le nombre de composantes demandé ne peut pas dépasser le nombre de spectres (ni le nombre de pics pour la PCA corrélée). Vérifiez que l'assemblage a réussi et que les spectres sont valides.

**Les spectres Contrôle n'apparaissent pas dans les analyses.**
Le label `Contrôle` dans la correspondance spectres ↔ tubes doit être orthographié avec une majuscule initiale. Il est automatiquement mappé au dernier tube du tableau des volumes.

**Des doublons de Raman Shift sont signalés.**
Ce n'est pas bloquant : l'application moyenne automatiquement les points dupliqués. Vérifiez néanmoins la qualité du fichier `.txt` source si ce message est fréquent.

**La PCA libre et la PCA corrélée donnent des résultats très différents.**
C'est attendu. La PCA corrélée n'utilise que les intensités aux pics choisis (quelques variables) ; la PCA libre utilise tout le spectre (environ 1 000 variables). Si un effet chimique important se manifeste en dehors des pics sélectionnés, seule la PCA libre le détectera.

**L'export graphique PDF ne correspond pas à l'affichage.**
Pour un rendu pixel-perfect, préférez l'export PNG haute résolution ou SVG vectoriel, qui utilisent directement le moteur de rendu de l'application. Le PDF passe par le moteur d'impression Qt et peut présenter de légères différences de mise en page.

---

## Crédits

- Développement : **Alexandre Souchaud** — CitizenSers
- Correction de baseline : [`pybaselines`](https://github.com/derb12/pybaselines) (Erb & Pelletier)
- Interface graphique : **PySide6 / Qt**
- Graphiques interactifs : **Plotly**
- PCA : **scikit-learn** (`sklearn.decomposition.PCA`)
- Ajustement sigmoïde : **SciPy** (`scipy.optimize.curve_fit`)

---

## Licence

Projet académique et interne — CitizenSers. Tous droits réservés.
