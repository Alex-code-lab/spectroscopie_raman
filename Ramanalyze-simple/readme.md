# Ramanalyze-simple

Version allégée de Ramanalyze : visualiser des spectres Raman et faire
une titration. Deux onglets.

## Lancer

Depuis ce dossier, avec l'environnement virtuel du projet :

```bash
../.spectroscopie_raman/bin/python main.py
```

## Onglet « Visualiseur »

On va chercher des `.txt`, ils s'affichent **en grand**, et le nom de
chaque courbe est le **nom du fichier**.

1. **📂 Ouvrir des fichiers…** → sélectionner un ou plusieurs `.txt`.
2. Les spectres s'affichent à droite, légende = nom de fichier.
3. Cocher / décocher un fichier pour l'afficher ou le masquer.
   **Retirer** / **Tout vider** pour nettoyer.

## Onglet « Titration »

Les fichiers chargés (depuis l'un ou l'autre onglet) apparaissent dans un
tableau **Fichier / Concentration**.

1. Saisir la **concentration** de chaque échantillon dans le tableau
   (double-clic dans la cellule). On peut aussi **Importer un CSV** à deux
   colonnes (`nom de fichier ; concentration`) ou **Exporter** le tableau.
2. Choisir la **source lumineuse** (532 nm / 785 nm) : la liste des
   **combinaisons de pics** recommandées s'affiche. Double-cliquer une paire
   la charge directement comme ratio (Pic 1 / Pic 2). Combinaisons reprises
   de Ramanalyze :
   - **532 nm** : (1364/1538), (1364/1409), (1364/1500), (1361/1631),
     (1364/1580), (1234/1256).
   - **785 nm** : toutes les paires de {412, 444, 471, 547, 1561}.
3. Indiquer le **pic à mesurer** (cm⁻¹) et la **tolérance**. Option : faire
   le **ratio** avec un second pic, et **corriger la ligne de base**
   (baseline modpoly, comme dans Ramanalyze).
4. **Tracer la titration** → courbe intensité (ou ratio) vs concentration.
   Option : **ajuster une sigmoïde** pour estimer le point d'équivalence.
5. **Titre** du graphique éditable (laisser vide = titre automatique).
6. **Exporter le graphique** en PNG, PDF ou HTML interactif.

## Format des fichiers

`.txt` avec une ligne d'en-tête commençant par `Pixel;`, séparateur `;`,
décimale `,`, contenant les colonnes `Raman Shift` et `Dark Subtracted #1`
(le brut, sans correction). Voir `spectrum_loader.py`.

## Dépendances

`PySide6`, `PySide6-Addons` (QtWebEngine), `plotly`, `pandas`, `numpy`
— déjà présentes dans l'environnement `.spectroscopie_raman` du projet.
