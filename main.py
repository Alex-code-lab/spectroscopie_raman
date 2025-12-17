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
            <h2>Présentation et mode d’emploi</h2>

            <p>
            Bienvenue dans <b>Ramanalyze</b>.<br>
            Cette application permet d’analyser des spectres Raman à partir de fichiers <code>.txt</code>,
            en les reliant à des informations expérimentales (métadonnées) saisies directement dans le programme.
            </p>

            <p>
            L’application fonctionne en <b>4 grandes étapes</b>, réparties dans les onglets.
            </p>

            <h3>1. Onglet <i>Fichiers</i></h3>
            <ul>
              <li>Sélectionnez un ou plusieurs fichiers de spectres Raman au format <code>.txt</code>.</li>
              <li>Cliquez sur <b>Ajouter la sélection</b> pour les ajouter à la liste.</li>
              <li>Les fichiers sélectionnés seront utilisés dans les onglets <i>Spectres</i>, <i>Analyse</i> et <i>PCA</i>.</li>
              <li>Tant qu’aucun fichier n’est sélectionné, les autres onglets ne peuvent pas fonctionner.</li>
            </ul>

            <h3>2. Onglet <i>Métadonnées</i></h3>

            <h4>a) Informations générales</h4>
            <ul>
              <li>Nom de la manipulation</li>
              <li>Date</li>
              <li>Lieu</li>
              <li>Coordinateur</li>
              <li>Opérateur</li>
            </ul>
            <p>
            Ces informations décrivent l’expérience et sont automatiquement recopiées dans les tableaux.
            </p>

            <h4>b) Tableau des volumes (composition des tubes)</h4>
            <ul>
              <li>Chaque ligne correspond à un réactif (Solution A, Solution B, etc.).</li>
              <li>Chaque colonne <b>Tube 1, Tube 2, …</b> correspond à un tube expérimental.</li>
              <li>Vous indiquez la concentration de la solution mère, son unité et les volumes ajoutés (en µL).</li>
              <li>Un modèle de tableau pré-rempli est proposé et peut être modifié.</li>
            </ul>

            <h4>c) Concentrations dans les cuvettes (information)</h4>
            <ul>
              <li>Ce tableau est calculé automatiquement à partir du tableau des volumes.</li>
              <li>Il indique les concentrations finales dans chaque tube.</li>
              <li>Il est en lecture seule et sert uniquement à l’analyse.</li>
            </ul>

            <h4>d) Correspondance spectres ↔ tubes</h4>
            <ul>
              <li>Associe chaque nom de spectre (ex. <code>AN335_00</code>) à un tube expérimental (ex. <code>A1</code>).</li>
              <li>Les noms de spectres sont générés automatiquement à partir du nom de la manipulation.</li>
              <li>Un bouton <b>Reset</b> permet de revenir aux valeurs par défaut.</li>
            </ul>

            <h4>Glossaire des solutions</h4>
            <ul>
              <li><b>Solution A</b> : Tampon (milieu de base)</li>
              <li><b>Solution B</b> : Titrant (paramètre expérimental variable)</li>
              <li><b>Solution C</b> : Indicateur</li>
              <li><b>Solution D</b> : PEG</li>
              <li><b>Solution E</b> : Nanoparticules</li>
              <li><b>Solution F</b> : Crosslinker</li>
            </ul>

            <h3>3. Onglet <i>Spectres</i></h3>
            <ul>
              <li>Affiche les spectres Raman à partir du fichier combiné interne.</li>
              <li>Une correction de baseline est appliquée automatiquement.</li>
              <li>Les spectres sont regroupés selon la description des échantillons.</li>
              <li>Il est possible d’ajuster les axes et d’exporter le graphique.</li>
            </ul>

            <h3>4. Onglet <i>Analyse</i></h3>
            <ul>
              <li>Permet d’analyser des pics Raman spécifiques.</li>
              <li>Les intensités et rapports d’intensité sont calculés automatiquement.</li>
              <li>Les graphiques sont tracés en fonction de la quantité de titrant.</li>
              <li>Les résultats peuvent être exportés vers Excel.</li>
            </ul>

            <h3>5. Onglet <i>PCA</i></h3>
            <ul>
              <li>Réalise une analyse en composantes principales (PCA).</li>
              <li>Permet de visualiser les scores, les loadings et la reconstruction de spectres.</li>
              <li>Utilise le même fichier combiné que les onglets Spectres et Analyse.</li>
            </ul>

            <h3>Remarques importantes</h3>
            <ul>
              <li>Les métadonnées sont indispensables pour l’analyse.</li>
              <li>Les concentrations sont toujours calculées automatiquement.</li>
              <li>Un seul fichier combiné est utilisé dans toute l’application, garantissant la cohérence des résultats.</li>
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