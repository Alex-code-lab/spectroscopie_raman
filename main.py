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
from metadata_picker import MetadataPickerWidget


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
            <h2>Présentation &amp; mode d'emploi</h2>
            <p>Bienvenue dans l'application <b>Ramanalyze</b>. 
            Cette interface vous guide du chargement des fichiers jusqu'à l'analyse des pics.</p>

            <h3>1. Onglet <i>Fichiers</i></h3>
            <ul>
              <li>Naviguez dans l'arborescence pour sélectionner un ou plusieurs fichiers <code>.txt</code> de spectres.</li>
              <li>Cliquez sur <b>Ajouter la sélection</b> pour les ajouter à la liste de droite.</li>
              <li>Utilisez <b>Retirer</b> pour enlever les éléments sélectionnés et <b>Vider la liste</b> pour repartir de zéro.</li>
            </ul>

            <h3>2. Onglet <i>Métadonnées</i></h3>
            <ul>
              <li>Sélectionnez d'abord le <b>fichier de composition des tubes</b> (contenant les concentrations, par ex. <code>C(EGTA)</code>, les volumes, etc.), puis cliquez sur <b>Valider le fichier de composition des tubes</b>.</li>
              <li>Sélectionnez ensuite le <b>fichier de correspondance des noms de spectres</b> (qui associe chaque nom de spectre <code>GCXXX_…</code> au tube correspondant), puis cliquez sur <b>Valider le fichier de correspondance des noms de spectres</b>.</li>
              <li>Vérifiez le contenu dans la table de visualisation (tri, défilement) pour contrôler que les données sont correctes.</li>
              <li>Cliquez sur <b>Assembler (données au format txt + métadonnées)</b> pour fusionner les fichiers <code>.txt</code> de spectres avec les deux fichiers de métadonnées.</li>
              <li>Si besoin, cliquez sur <b>Enregistrer le fichier</b> pour sauvegarder le fichier combiné (au choix <code>.csv</code> ou <code>.xlsx</code>). Ce fichier combiné est ensuite utilisé par les onglets <i>Spectres</i> et <i>Analyse</i>.</li>
            </ul>

            <h3>3. Onglet <i>Spectres</i></h3>
            <ul>
              <li>Le tracé utilise directement le <b>fichier combiné</b> (quand il existe) pour afficher les courbes.</li>
              <li>Les légendes utilisent prioritairement la colonne <b>Sample description</b> issue de l'Excel.</li>
              <li>Si le fichier combiné n'est pas encore prêt, un mode de secours relit les <code>.txt</code> (moins recommandé).</li>
            </ul>

            <h3>4. Onglet <i>Analyse</i></h3>
            <ul>
              <li>Cliquez sur <b>Recharger le fichier combiné</b> si nécessaire (après l'avoir validé dans Métadonnées).</li>
              <li>Choisissez un <b>jeu de pics</b> (ex. 532 nm ou 785 nm) et une <b>tolérance</b> en cm⁻¹.</li>
              <li>Cliquez sur <b>Analyser les pics</b> pour calculer les intensités autour des pics et tous les ratios.</li>
              <li>Le tableau affiche toutes les lignes ; le <b>graphique Plotly</b> montre les rapports d'intensité selon [EGTA].</li>
              <li>Utilisez <b>Exporter résultats (Excel)…</b> pour sauvegarder les tables <i>intensites</i> et <i>ratios</i>.</li>
            </ul>

            <h3>Conseils &amp; remarques</h3>
            <ul>
              <li>Le nom des courbes provient de <b>Sample description</b> dès que le fichier combiné est disponible.</li>
              <li>La correction de baseline (<i>modpoly</i>) est réalisée lors de l'assemblage.</li>
              <li>Vous pouvez trier les colonnes des tableaux en cliquant sur leurs en-têtes.</li>
              <li>Pour des performances optimales, évitez d'ouvrir des centaines de fichiers simultanément.</li>
            </ul>
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

        # Onglet Métadonnées
        self.metadata_tab = QWidget(self)
        metadata_layout = QVBoxLayout(self.metadata_tab)
        self.metadata_picker = MetadataPickerWidget(self.metadata_tab)
        metadata_layout.addWidget(self.metadata_picker)
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