import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabWidget,
    QWidget,
    QVBoxLayout,
    QStatusBar,
    QLabel,
    QScrollArea,
    QPushButton,
    QDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap

# S'assurer que les imports de modules frères fonctionnent quel que soit l'endroit
# d'où on lance le script
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from file_picker import FilePickerWidget
from spectra_plot import SpectraTab
from analysis_tab import AnalysisTab
from metadata_creator import MetadataCreatorWidget


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ramanalyze - Analyse de spectres Raman")
        self.resize(1200, 800)
        # Fixer la largeur de la fenêtre pour éviter qu'elle ne s'agrandisse en fonction du contenu,
        # tout en permettant à l'utilisateur de modifier la hauteur.
        self.setMinimumSize(1200, 600)

        # --- Status bar ---
        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Prêt")

        # --- Conteneur principal ---
        central_widget = QWidget(self)
        central_layout = QVBoxLayout(central_widget)
        self.setCentralWidget(central_widget)

        # --- Onglets ---
        tabs = QTabWidget(self)
        central_layout.addWidget(tabs)

        # Onglet Présentation (instructions d'utilisation)
        self.presentation_tab = QWidget(self)
        pres_layout = QVBoxLayout(self.presentation_tab)

        # --- Logo page d'accueil ---
        logo_path = os.path.join(APP_DIR, "Logo Ramanalyze", "logo ramanalyze.svg")
        logo_pixmap = QPixmap(logo_path)
        if not logo_pixmap.isNull():
            logo_label = QLabel()
            logo_label.setPixmap(logo_pixmap.scaledToHeight(90, Qt.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignCenter)
            pres_layout.addWidget(logo_label)

        scroll = QScrollArea(self.presentation_tab)
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)

        doc_html = """
<style>
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    color: #d4d4d4;
    max-width: 860px;
  }
  h2 { color: #46ba48; margin-top: 6px; margin-bottom: 8px; }
  h3 { color: #46ba48; margin-top: 22px; margin-bottom: 6px; border-bottom: 1px solid #444; padding-bottom: 3px; }
  h4 { color: #6fa8d6; margin-top: 16px; margin-bottom: 4px; }
  p  { line-height: 1.6; margin-bottom: 8px; }
  ul, ol { margin-left: 20px; }
  li { margin-bottom: 5px; line-height: 1.5; }
  .key    { font-weight: 700; color: #e06c75; }
  .accent { font-weight: 600; color: #6fa8d6; }
  .warn   { font-weight: 600; color: #e06c75; }
  .ok     { font-weight: 600; color: #46ba48; }
  .mono   { font-family: Menlo, Consolas, "Courier New", monospace; font-size: 0.92em;
            background: #000000 !important; color: #ffffff !important; padding: 1px 4px; border-radius: 3px; }
  .note   { border-left: 4px solid #6fa8d6; background: #1e2a36;
            padding: 8px 12px; margin: 12px 0; border-radius: 0 4px 4px 0; }
  .warn-box { border-left: 4px solid #e07b00; background: #2a1f10;
              padding: 8px 12px; margin: 12px 0; border-radius: 0 4px 4px 0; }
  table { border-collapse: collapse; margin: 8px 0; }
  td, th { border: 1px solid #444; padding: 5px 10px; font-size: 0.93em; }
  th { background: #2a2a2a; color: #d4d4d4; font-weight: 600; }
</style>

<h2>Ramanalyze — Guide d’utilisation</h2>

<p>
  Bienvenue dans <span class="key">Ramanalyze</span>, application d’analyse
  de spectres Raman et SERS.<br>
  Elle relie vos fichiers spectres <span class="mono">.txt</span> à des
  <span class="accent">métadonnées expérimentales</span> (volumes, concentrations,
  correspondance tubes), et offre des outils d’analyse quantitative et d’exploration
  multivariée (PCA).
</p>

<h3>Parcours recommandé — 5 étapes</h3>
<ol>
  <li>
    <span class="key">Métadonnées</span> — Remplissez le tableau des volumes,
    la correspondance spectres ↔ tubes et utilisez la feuille de protocole.
    Enregistrez en <span class="mono">.xlsx</span>.
  </li>
  <li>
    <span class="key">Fichiers Raman</span> — Sélectionnez vos fichiers
    <span class="mono">.txt</span> et cliquez <b>Ajouter la sélection</b>.
  </li>
  <li>
    <span class="key">Spectres</span> — Vérifiez la qualité visuelle des spectres
    (correction de baseline, légendes).
  </li>
  <li>
    <span class="key">Analyse</span> — Choisissez vos pics, lancez l’analyse
    des intensités et ratios, exportez en Excel.
  </li>
  <li>
    <span class="key">Exploration</span> — Explorez la PCA pour comprendre
    quelles bandes Raman discriminent le mieux vos échantillons.
  </li>
</ol>

<div class="note">
  <b>Guidage visuel :</b>
  les boutons <span class="warn">rouges</span> signalent une action requise ;
  les <span class="ok">verts</span> indiquent une étape validée.<br>
  Après toute modification des fichiers ou des métadonnées, rechargez le
  <b>fichier combiné</b> dans les onglets Analyse et Exploration.
</div>

<h3>1 — Onglet Métadonnées</h3>

<h4>Tableau des volumes</h4>
<p>
  Chaque ligne est un réactif, chaque colonne est un tube (Tube 1, Tube 2…).
  Les valeurs sont en µL. Le bouton <b>Générer volumes (gaussien)</b> remplit
  automatiquement les colonnes A et B avec des volumes centrés autour de
  l’équivalence chimique.
</p>

<table>
  <tr><th>Solution</th><th>Rôle</th></tr>
  <tr><td><b>Solution A</b></td><td>Tampon (volume variable, complémentaire à B)</td></tr>
  <tr><td><b>Solution B</b></td><td>Titrant (<span class="accent">paramètre variable</span>)</td></tr>
  <tr><td><b>Solution C</b></td><td>Indicateur</td></tr>
  <tr><td><b>Solution D</b></td><td>PEG</td></tr>
  <tr><td><b>Solution E</b></td><td>Nanoparticules</td></tr>
  <tr><td><b>Solution F</b></td><td>Crosslinker</td></tr>
</table>

<h4>Correspondance spectres ↔ tubes</h4>
<p>
  Associez chaque fichier spectre à un tube. Cette correspondance est utilisée
  pour relier les concentrations calculées à chaque spectre lors de l’assemblage.
  Le label <span class="mono">Contrôle</span> est automatiquement mappé au
  dernier tube (réplicat).
</p>

<h4>Feuille de protocole de paillasse</h4>
<p>
  Le bouton <span class="key">Feuille de protocole…</span> ouvre un dialogue
  interactif de <span class="accent">suivi du pipetage</span> sur la paillasse.
  Il affiche le tableau <span class="mono">df_comp</span> sous forme d'une grille
  où chaque réactif est une ligne et chaque tube une colonne.
</p>
<p>Pour chaque tube, trois sous-colonnes sont proposées :</p>
<ul>
  <li><b>Volume (µL)</b> — lecture seule, rappelle la quantité à pipeter.</li>
  <li><b>Coordinateur ☐</b> — case à cocher validée par le coordinateur.</li>
  <li><b>Opérateur ☐</b> — case à cocher validée par l'opérateur.</li>
</ul>
<p>
  Une colonne <b>Vérification solution</b> supplémentaire permet de confirmer
  l'identité du flacon avant de commencer le pipetage d'un réactif.
</p>

<h4>Mode guidé (étape par étape)</h4>
<p>
  En mode guidé, <b>seule la cellule de l'étape courante est active</b> ;
  toutes les autres sont grisées et désactivées.
  La séquence automatique est la suivante :
</p>
<ol>
  <li>Vérification de la solution (colonne « Vérif ») pour le 1<sup>er</sup> réactif.</li>
  <li>Tube 1 — Volume confirmé + case Coordinateur + case Opérateur.</li>
  <li>Tube 2 — idem, puis Tube 3… jusqu'au dernier tube.</li>
  <li>Passage au réactif suivant : Vérif → Tube 1 → Tube 2 → …</li>
</ol>
<p>
  Un libellé de statut vert indique l'étape en cours :
  <span class="mono">→ Vérifier la solution : Solution A</span> ou
  <span class="mono">→ Solution A · Tube 1</span>.
</p>

<div class="note">
  <b>Workflow coordinateur / opérateur :</b>
  le coordinateur coche sa case en premier pour indiquer qu'il a préparé
  le tube, puis l'opérateur valide après vérification indépendante.
  Les boutons <b>Tout cocher — Coord.</b> et <b>Tout cocher — Opér.</b>
  permettent de valider toutes les cases d'un rôle en un clic.
  <b>Tout décocher</b> remet tout à zéro.
</div>

<h4>Export et rechargement du fichier unifié</h4>
<p>
  Le bouton <b>Exporter Excel…</b> génère un fichier <span class="mono">.xlsx</span>
  contenant <b>4 feuilles</b> :
</p>
<ul>
  <li><span class="mono">Protocole</span> — feuille visuelle mise en forme
      (en-têtes colorés, texte pivoté, symboles ✓/☐).</li>
  <li><span class="mono">Volumes</span> — données brutes de <span class="mono">df_comp</span>.</li>
  <li><span class="mono">Correspondance</span> — données brutes de <span class="mono">df_map</span>.</li>
  <li><span class="mono">EtatProtocole</span> — état de chaque case (type, r_idx, t_idx, checked).</li>
</ul>
<p>
  Ce fichier unifié peut être rechargé via <b>Charger des métadonnées…</b> :
  toutes les cases reprennent exactement l'état sauvegardé.
  L'état est également préservé en mémoire lorsque vous fermez le dialogue,
  et inclus automatiquement dans <b>Enregistrer les métadonnées…</b>.
</p>

<h4>Concentrations en cuvette</h4>
<p>
  Le bouton <b>Voir concentrations</b> affiche le tableau des concentrations
  finales dans chaque tube, calculées à partir des volumes et des concentrations
  stock (lecture seule).
</p>

<h4>Assemblage et sauvegarde</h4>
<ul>
  <li>Cliquez <b>Assembler</b> pour charger les spectres, corriger la baseline
      et fusionner avec les métadonnées.</li>
  <li>Vérifiez le message de succès (nombre de lignes / colonnes).</li>
  <li>Cliquez <b>Enregistrer</b> pour sauvegarder le fichier combiné
      (CSV ou Excel).</li>
</ul>

<h3>2 — Onglet Fichiers Raman</h3>
<ul>
  <li>Double-cliquez sur un dossier pour le parcourir.</li>
  <li>Sélectionnez un ou plusieurs fichiers <span class="mono">.txt</span>
      (résultats Raman bruts).</li>
  <li>Cliquez <b>Ajouter la sélection</b> — les fichiers apparaissent dans
      la liste de droite.</li>
  <li>Utilisez <b>Retirer</b> ou <b>Vider la liste</b> si nécessaire.</li>
</ul>

<h3>3 — Onglet Spectres</h3>
<ul>
  <li>Affiche les spectres corrigés du fichier combiné.</li>
  <li>La légende utilise <span class="mono">Sample description</span>
      si disponible.</li>
  <li>Cliquez <b>Tracer avec baseline</b> pour générer la figure Plotly
      interactive (zoom, pan, export PNG).</li>
</ul>

<h3>4 — Onglet Analyse</h3>
<ul>
  <li>Choisissez un jeu de pics prédéfini :
    <ul>
      <li><b>532 nm</b> : 1231, 1327, 1342, 1358, 1450 cm⁻¹</li>
      <li><b>785 nm</b> : 412, 444, 471, 547, 1561 cm⁻¹</li>
    </ul>
  </li>
  <li>Réglez la <b>tolérance</b> (fenêtre de recherche autour de chaque pic,
      défaut 2 cm⁻¹).</li>
  <li>Cliquez <b>Analyser les pics</b> pour calculer :
    <ul>
      <li>l’intensité maximale dans chaque fenêtre, par spectre ;</li>
      <li>tous les ratios entre paires de pics ;</li>
      <li>un graphique interactif des ratios en fonction des métadonnées.</li>
    </ul>
  </li>
  <li>Exportez en Excel (feuilles <span class="mono">intensites</span>
      et <span class="mono">ratios</span>).</li>
</ul>

<h3>5 — Onglet Exploration (PCA)</h3>

<h4>Sous-onglets disponibles</h4>
<table>
  <tr><th>Sous-onglet</th><th>Ce qu’il montre</th></tr>
  <tr><td>Spectres normalisés</td><td>Spectres après normalisation par la norme vectorielle</td></tr>
  <tr><td>Spectres superposés</td><td>Superposition brute des spectres</td></tr>
  <tr><td>PCA corrélée</td><td>PCA sur les intensités aux pics sélectionnés</td></tr>
  <tr><td>PCA — Scores</td><td>Position de chaque spectre dans l’espace PCA libre (colorié par variable)</td></tr>
  <tr><td>PCA — Loadings</td><td>Poids de chaque nombre d’onde pour chaque composante</td></tr>
  <tr><td>PCA — Reconstruction</td><td>Reconstruction d’un spectre à partir de N composantes</td></tr>
  <tr><td>Scatter matrix</td><td>Corrélations visuelles entre tous les pics</td></tr>
</table>

<h4>Comprendre la PCA en bref</h4>
<p>
  Chaque spectre est un vecteur de ~1 000 intensités (une par nombre d’onde).
  La PCA cherche les <b>directions</b> dans cet espace à 1 000 dimensions
  qui capturent le plus de variabilité entre spectres.
</p>
<ul>
  <li>
    <b>PC1</b> = direction qui explique <i>le plus</i> de variance.
    Pas forcément la plus chimiquement intéressante : elle peut capturer
    une variation de ligne de base ou de puissance laser.
  </li>
  <li>
    <b>PC2</b> = perpendiculaire à PC1, 2<sup>e</sup> source de variabilité.
    C’est souvent là que l’effet chimique (concentration) apparaît.
  </li>
  <li>
    <b>Loading</b> = le "spectre fantôme" d’une composante. Tracé en fonction
    du Raman Shift, il indique quelles bandes Raman portent cette composante.
    Interprétez-le comme un spectre normal.
  </li>
  <li>
    <b>Score</b> d’un spectre = sa coordonnée sur un axe PCA.
    Score positif élevé = le spectre ressemble fortement au pattern du loading.
  </li>
  <li>
    <b>Reconstruction</b> avec N=1 : visualise ce que PC1 a capturé seule.
    Ajoutez des composantes progressivement pour voir ce que chacune apporte.
  </li>
</ul>

<div class="warn-box">
  <b>Conseil :</b> Si les points du graphe Scores PC1 vs PC2 colorés par
  concentration ne se séparent pas selon PC1, regardez PC2 et PC3 — l’effet
  chimique n’est pas toujours dans la première composante.
</div>

<h3>Si un résultat semble incohérent</h3>
<ul>
  <li>Vérifiez la correspondance spectres ↔ tubes (les bons spectres sont-ils
      assignés aux bons tubes ?).</li>
  <li>Vérifiez les unités des concentrations stock dans le tableau des volumes.</li>
  <li>Rechargez le fichier combiné dans l’onglet concerné après toute
      modification des métadonnées.</li>
  <li>En PCA, vérifiez que la correction de baseline est satisfaisante :
      une mauvaise baseline peut dominer PC1 et masquer l’effet chimique.</li>
</ul>
"""

        text_label = QLabel(doc_html, container)
        text_label.setWordWrap(True)
        text_label.setTextFormat(Qt.RichText)
        container_layout.addWidget(text_label)
        container_layout.addStretch(1)

        scroll.setWidget(container)
        pres_layout.addWidget(scroll)

        tabs.insertTab(0, self.presentation_tab, "Présentation")
        tabs.setCurrentIndex(0)

        # Onglet Métadonnées (création directe des tableaux dans l'application)
        self.metadata_tab = QWidget(self)
        metadata_layout = QVBoxLayout(self.metadata_tab)
        self.metadata_creator = MetadataCreatorWidget(self.metadata_tab)
        metadata_layout.addWidget(self.metadata_creator)
        tabs.addTab(self.metadata_tab, "Métadonnées")

        # Onglet Fichiers Raman
        self.file_tab = QWidget(self)
        file_layout = QVBoxLayout(self.file_tab)
        self.file_picker = FilePickerWidget(self.file_tab)
        file_layout.addWidget(self.file_picker)
        tabs.addTab(self.file_tab, "Fichiers Raman")

        # Onglet Spectres
        self.spectra_tab = SpectraTab(self.file_picker, self.metadata_creator, self)
        tabs.addTab(self.spectra_tab, "Spectres")

        # Onglet Analyse
        self.analysis_tab = AnalysisTab(self.file_picker, self.metadata_creator, self)
        tabs.addTab(self.analysis_tab, "Analyse")

        # Onglet Exploration (PCA + sélection de pics fusionnés)
        from peak_selector import PeakSelectorTab
        self.peak_selector_tab = PeakSelectorTab(self.file_picker, self.metadata_creator, self)
        tabs.addTab(self.peak_selector_tab, "Exploration")

        # --- Label des sources en bas ---
        self.sources_label = QLabel("L'ensemble des sources sont à retrouver <a href='#'>ici</a>.")
        self.sources_label.setTextFormat(Qt.RichText)
        self.sources_label.setOpenExternalLinks(False)
        self.sources_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.LinksAccessibleByMouse)
        self.sources_label.setAlignment(Qt.AlignCenter)
        self.sources_label.linkActivated.connect(self.show_sources_popup)
        central_layout.addWidget(self.sources_label)

    def show_sources_popup(self, link_str):
        """Affiche une fenêtre contextuelle contenant les sources et références de l'application."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Sources")
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)
        sources_text = """
        <p><b>Sources et Références :</b></p>
        <ul>
            <li><b><a href="https://msc.u-paris.fr/"> Laboratoire Matière et Systèmes Complexes</a></b><br>Laboratoire de l'université de Paris.</li>
            <li><b><a href="https://www.eau-et-rivieres.org/home/"> Eau & Rivières de Bretagne</a></b><br>Association de protection de l'environnement.</li>
            <li><b><a href="https://www.campus-transition.org/fr/"> Campus de la Transition</a></b><br>Association loi 1901 de formation de l'Enseignement Superieur en vue d'une transition écologique et solidaire.</li>         
        </ul>
        """
        # <p><b>Articles Scientifiques :</b></p>
        # <ul>
        #     <li><em>Using life cycle assessments to guide reduction in the carbon footprint of single-use lab consumables</em> — Isabella Ragazzi, <a href="https://doi.org/10.1371/journal.pstr.0000080">PLOS, 2023</a></li>
        #     <li><em>The environmental impact of personal protective equipment in the UK healthcare system</em> — Reed et al., <a href="https://journals.sagepub.com/doi/epub/10.1177/01410768211001583">JRSM, 2021</a></li>
        # </ul>

        # <p><b>Crédits :</b></p>
        # <ul>
        #     <li><a href="https://www.ilfotografico.net/">Dario Danile</a> : Graphiste de l'icône de l'application.</li>
        #     <li><b>Alexandre Souchaud</b> : Codeur de l'application.</li>
        # </ul>
  
        label = QLabel()
        label.setTextFormat(Qt.RichText)
        label.setOpenExternalLinks(True)
        label.setText(sources_text)
        layout.addWidget(label)

        close_button = QPushButton("Fermer")
        close_button.clicked.connect(dialog.close)
        layout.addWidget(close_button)

        dialog.setLayout(layout)
        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    if sys.platform == "win32":
        icon_path = os.path.join(APP_DIR, "Logo Ramanalyze", "logo-ramanalyze.ico")
    else:
        icon_path = os.path.join(APP_DIR, "Logo Ramanalyze", "logo ramanalyze.svg")
    app.setWindowIcon(QIcon(icon_path))

    win = MainWindow()
    win.show()
    sys.exit(app.exec())
