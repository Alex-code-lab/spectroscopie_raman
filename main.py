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
        self.setMaximumWidth(1200)

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

        scroll = QScrollArea(self.presentation_tab)
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)

        doc_html = (
            """
<style>
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    color: #2b2b2b;
  }

  h2 {
    color: #46ba48;
    margin-bottom: 6px;
  }

  h3 {
    color: #46ba48 ;
    margin-top: 18px;
    margin-bottom: 6px;
  }

  h4 {
    color: #3a7ca5;
    margin-top: 14px;
    margin-bottom: 4px;
  }

  p {
    line-height: 1.55;
    margin-bottom: 10px;
  }

  ul, ol {
    margin-left: 18px;
  }

  li {
    margin-bottom: 6px;
  }

  .key {
    font-weight: 600;
    color: #c4293d;
  }

  .accent {
    font-weight: 600;
    color: #2f6fad;
  }

  .warn {
    font-weight: 600;
    color: #b00020;
  }

  .ok {
    font-weight: 600;
    color: #1e7f43;
  }

  .note {
    border-left: 3px solid #2f6fad;
    padding-left: 10px;
    margin: 10px 0;
    color: #ffffff;
  }

  .mono {
    font-family: Menlo, Consolas, monospace;
    font-size: 0.95em;
  }
</style>

<h2>Guide d’utilisation pas à pas</h2>

<p>
Bienvenue dans <span class="key">Ramanalyze</span>.<br>
Cette application permet d’analyser des spectres Raman (<span class="mono">.txt</span>) en les liant à des
<span class="accent">métadonnées expérimentales</span>.  
Le guide ci‑dessous décrit l’ordre recommandé des étapes.
</p>

<h3>Parcours rapide — 5 étapes</h3>
<ol>
  <li><span class="key">Fichiers</span> : sélectionnez les fichiers de résultats Raman (<span class="mono">.txt</span>) puis cliquez sur <b>Ajouter la sélection</b>.</li>
  <li><span class="key">Métadonnées</span> : complétez les informations générales, le tableau des volumes et la correspondance spectres ↔ tubes.</li>
  <li><span class="key">Spectres</span> : vérifiez la qualité des spectres (baseline, axes, export PNG si besoin).</li>
  <li><span class="key">PCA</span> : explorez regroupements et variances pour identifier les pics et paramètres importants.</li>
  <li><span class="key">Analyse</span> : choisissez les pics, lancez l’analyse, puis ajustez la sigmoïde pour obtenir le point d’équivalence.</li>
</ol>

<h3>Guidage visuel</h3>
<ul>
  <li><span class="warn">Rouge</span> : action requise</li>
  <li><span class="ok">Vert</span> : étape prête</li>
  <li>Après toute modification des fichiers ou métadonnées, rechargez le <b>fichier combiné</b> dans <b>Analyse</b> ou <b>PCA</b>.</li>
</ul>

<h3>Détails par onglet</h3>

<h4>1) Onglet Fichiers</h4>
<ul>
  <li>Navigation par double‑clic dans les dossiers.</li>
  <li>Sélection de fichiers <span class="mono">.txt</span> de résultats Raman.</li>
  <li><b>Ajouter la sélection</b> pour constituer le jeu d’analyse.</li>
</ul>

<h4>2) Onglet Métadonnées</h4>
<ul>
  <li>Informations générales de l’expérience (<b>nom requis</b>).</li>
  <li>Tableau des volumes (solutions, unités, µL), entièrement personnalisable.</li>
  <li>Chargement / sauvegarde de modèles de métadonnées.</li>
  <li>Concentrations en cuvette calculées automatiquement (lecture seule).</li>
  <li>Vérification de la correspondance spectres ↔ tubes.</li>
</ul>

<h4>Glossaire des solutions</h4>
<ul>
  <li><b>Solution A</b> : Tampon</li>
  <li><b>Solution B</b> : Titrant (<span class="accent">paramètre variable</span>)</li>
  <li><b>Solution C</b> : Indicateur</li>
  <li><b>Solution D</b> : PEG</li>
  <li><b>Solution E</b> : Nanoparticules</li>
  <li><b>Solution F</b> : Crosslinker</li>
</ul>

<h4>3) Onglet Spectres</h4>
<ul>
  <li>Tracé avec correction de baseline.</li>
  <li>Option d’inclusion / exclusion du contrôle BRB.</li>
  <li>Axes configurables et export PNG.</li>
</ul>

<h4>4) Onglet Analyse (pics & équivalence)</h4>
<ul>
  <li>Sélection de presets (532 / 785 nm) ou mode personnalisé.</li>
  <li>Calcul automatique des intensités et ratios (export Excel).</li>
  <li>Ajustement <span class="accent">sigmoïdal</span> pour déterminer l’équivalence.</li>
  <li>Affichage du point d’inflexion et de son incertitude.</li>
</ul>

<h4>5) Onglet PCA</h4>
<ul>
  <li>Rechargement du fichier combiné.</li>
  <li>Choix du domaine Raman, de la normalisation et du nombre de composantes.</li>
  <li>Visualisation des scores, loadings et reconstructions.</li>
</ul>

<h3>Si un résultat semble incohérent</h3>
<div class="note">
  <ul>
    <li>Vérifiez la correspondance spectres ↔ tubes et les unités.</li>
    <li>Relancez l’analyse après toute modification.</li>
    <li>Assurez‑vous que les ratios ne sont pas quasi constants avant un fit sigmoïde.</li>
  </ul>
</div>
            """
        )

        text_label = QLabel(doc_html, container)
        text_label.setWordWrap(True)
        text_label.setTextFormat(Qt.RichText)
        container_layout.addWidget(text_label)
        container_layout.addStretch(1)

        scroll.setWidget(container)
        pres_layout.addWidget(scroll)

        tabs.insertTab(0, self.presentation_tab, "Présentation")
        tabs.setCurrentIndex(0)

        # Onglet Fichiers
        self.file_tab = QWidget(self)
        file_layout = QVBoxLayout(self.file_tab)
        self.file_picker = FilePickerWidget(self.file_tab)
        file_layout.addWidget(self.file_picker)
        tabs.addTab(self.file_tab, "Fichiers")

        # Onglet Métadonnées (création directe des tableaux dans l'application)
        self.metadata_tab = QWidget(self)
        metadata_layout = QVBoxLayout(self.metadata_tab)
        self.metadata_creator = MetadataCreatorWidget(self.metadata_tab)
        metadata_layout.addWidget(self.metadata_creator)
        tabs.addTab(self.metadata_tab, "Métadonnées")

        # Onglet Spectres
        self.spectra_tab = SpectraTab(self.file_picker, self)
        tabs.addTab(self.spectra_tab, "Spectres")

        # Onglet PCA
        from pca import PCATab
        self.pca_tab = PCATab(self)
        tabs.addTab(self.pca_tab, "PCA")

        # Onglet Analyse
        self.analysis_tab = AnalysisTab(self)
        tabs.addTab(self.analysis_tab, "Analyse")

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
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
