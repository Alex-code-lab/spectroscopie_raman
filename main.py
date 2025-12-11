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
            <h2>Présentation &amp; mode d'emploi</h2>
            <p>Bienvenue dans l'application <b>Ramanalyze</b>. 
            Cette interface vous guide du chargement des fichiers jusqu'à l'analyse des pics et la PCA.</p>

            <h3>1. Onglet <i>Fichiers</i></h3>
            <ul>
              <li>Naviguez dans l'arborescence pour sélectionner un ou plusieurs fichiers <code>.txt</code> de spectres.</li>
              <li>Cliquez sur <b>Ajouter la sélection</b> pour les ajouter à la liste de droite.</li>
              <li>Utilisez <b>Retirer</b> pour enlever les éléments sélectionnés et <b>Vider la liste</b> pour repartir de zéro.</li>
              <li>Les fichiers <code>.txt</code> sélectionnés seront utilisés par les onglets <i>Spectres</i>, <i>PCA</i> et <i>Analyse</i>.</li>
            </ul>

            <h3>2. Onglet <i>Métadonnées</i></h3>
            <p>Dans cette version, il n’y a plus de fichiers Excel de métadonnées à charger : 
            tout est créé directement dans le programme sous forme de tableaux éditables (type Excel), 
            mais stockés en interne comme des DataFrames pandas.</p>

            <ul>
              <li>Commencez par renseigner les champs d’en-tête :
                <ul>
                  <li><b>Nom de la manip</b> (ex. <code>GC525_20251203</code>)</li>
                  <li><b>Date</b> (sélection via un calendrier)</li>
                  <li><b>Lieu</b> (laboratoire, terrain, etc.)</li>
                  <li><b>Coordinateur</b></li>
                  <li><b>Opérateur</b></li>
                </ul>
                Ces informations sont recopiées automatiquement dans le tableau de correspondance spectres ↔ tubes.
              </li>

              <li>Cliquez sur <b>Créer / éditer le tableau des volumes</b> pour définir la composition des tubes :
                <ul>
                  <li>Chaque ligne correspond à un réactif (échantillon, eau de contrôle, Solution A, Solution B, etc.).</li>
                  <li>Vous indiquez pour chaque réactif sa <b>concentration</b> et son <b>unité</b> (par ex. <code>0,5 µM</code>, <code>4 mM</code>, <code>% en masse</code>, …).</li>
                  <li>Les colonnes <b>Tube 1</b> à <b>Tube 11</b> contiennent les volumes ajoutés dans chaque tube (en µL).</li>
                  <li>Un tableau par défaut est proposé (pré-rempli) que vous pouvez adapter à votre protocole.</li>
                </ul>
              </li>

              <li>Cliquez sur <b>Calculer / afficher le tableau des concentrations</b> :
                <ul>
                  <li>Le programme calcule automatiquement, pour chaque tube, le <b>volume total</b> (en µL).</li>
                  <li>Pour chaque réactif, la <b>concentration finale</b> dans chaque tube est calculée en tenant compte :
                    <ul>
                      <li>de la concentration de stock (convertie en M : M, mM, µM, …)</li>
                      <li>et des volumes (convertis en L : µL → L).</li>
                    </ul>
                  </li>
                  <li>Un réactif identifié comme <b>titrant</b> (par exemple <code>Solution B</code>) donne lieu à des colonnes dédiées 
                  <code>[titrant] (M)</code> et <code>[titrant] (mM)</code> utilisées dans l’onglet Analyse.</li>
                </ul>
              </li>

              <li>Cliquez sur <b>Créer / éditer la correspondance spectres ↔ tubes</b> :
                <ul>
                  <li>Un tableau structuré est généré avec :
                    <ul>
                      <li>quelques lignes d’en-tête (Nom de la manip, Date, Lieu, Coordinateur, Opérateur),</li>
                      <li>puis une section <b>Nom du spectre / Tube</b>.</li>
                    </ul>
                  </li>
                  <li>Les noms de spectres sont automatiquement pré-remplis à partir du nom de la manip, 
                  sous la forme <code>NomManip_00</code>, <code>NomManip_01</code>, …</li>
                  <li>Les tubes <b>Tube BRB</b>, <b>Tube MQW</b> et les tubes numérotés (Tube 1, Tube 2, …) sont également pré-remplis.</li>
                  <li>Vous pouvez ajuster ces correspondances si nécessaire (par exemple si certains spectres sont manquants ou en double).</li>
                </ul>
              </li>

              <li>Les tableaux de volumes, de concentrations et de correspondance sont ensuite utilisés par les onglets 
              <i>Spectres</i>, <i>PCA</i> et <i>Analyse</i> pour reconstruire un fichier combiné interne 
              (<b>aucun fichier Excel de métadonnées n’est plus nécessaire</b>).</li>
            </ul>

            <h3>3. Onglet <i>Spectres</i></h3>
            <ul>
              <li>Le tracé utilise directement le <b>fichier combiné interne</b> (spectres txt + métadonnées) 
              construit à partir de l’onglet <i>Métadonnées</i>.</li>
              <li>Dès que les métadonnées sont définies, les légendes utilisent prioritairement la colonne 
              <b>Sample description</b> (issue du tableau de correspondance).</li>
              <li>Si les métadonnées ne sont pas encore complètes, un mode de secours peut afficher simplement les spectres 
              à partir des fichiers <code>.txt</code> (avec des noms de fichiers bruts).</li>
            </ul>

            <h3>4. Onglet <i>Analyse</i></h3>
            <ul>
              <li>Cliquez sur <b>Recharger le fichier combiné depuis Métadonnées</b> pour (re)construire le DataFrame fusionné 
              à partir des fichiers <code>.txt</code> et des tableaux créés dans l’onglet <i>Métadonnées</i>.</li>
              <li>Choisissez un <b>jeu de pics</b> (par exemple :
                <ul>
                  <li><b>532 nm</b> avec les pics 1231, 1327, 1342, 1358, 1450 cm⁻¹</li>
                  <li><b>785 nm</b> avec les pics 412, 444, 471, 547, 1561 cm⁻¹</li>
                  <li>ou un mode <b>Personnalisé</b> où vous saisissez directement les décalages Raman.</li>
                </ul>
              </li>
              <li>Réglez la <b>tolérance</b> (fenêtre en cm⁻¹ autour de chaque pic) puis cliquez sur 
              <b>Analyser les pics</b> :
                <ul>
                  <li>pour chaque spectre, l’intensité maximale est extraite autour de chaque pic choisi ;</li>
                  <li>tous les <b>rapports d’intensité</b> entre paires de pics sont calculés (ratio I_a / I_b).</li>
                </ul>
              </li>
              <li>Le <b>tableau des intensités et des ratios</b> peut être affiché ou masqué pour gagner de la place.</li>
              <li>Le <b>graphique Plotly</b> affiche les rapports d’intensité en fonction de la quantité de <b>titrant</b> :
                <ul>
                  <li>l’axe X correspond à <code>n(titrant) (mol)</code> quand c’est possible,</li>
                  <li>cette quantité est calculée à partir des concentrations finales (<code>[titrant] (M)</code>) 
                  et du volume de cuvette (<code>V cuvette (µL)</code>).</li>
                </ul>
              </li>
              <li>Utilisez <b>Exporter résultats (Excel)…</b> pour sauvegarder les tables d’intensités et de ratios 
              et poursuivre l’analyse dans un autre logiciel.</li>
            </ul>

            <h3>5. Onglet <i>PCA</i></h3>
            <ul>
              <li>Cliquez sur <b>Recharger le fichier combiné depuis Métadonnées</b> pour récupérer le même fichier combiné 
              utilisé par l’onglet <i>Analyse</i>.</li>
              <li>Choisissez une plage de <b>Raman Shift</b> (min/max) ainsi qu’un mode de <b>normalisation</b> 
              (aucune, norme L2, standardisation).</li>
              <li>Lancez la <b>PCA</b> pour :
                <ul>
                  <li>afficher un nuage de points des <b>scores PCA</b> (PC1 vs PC2, colorés par une métadonnée au choix),</li>
                  <li>visualiser les <b>loadings</b> (poids des composantes en fonction du shift Raman),</li>
                  <li>comparer un spectre original et sa <b>reconstruction</b> à partir des composantes principales.</li>
                </ul>
              </li>
            </ul>

            <h3>Conseils &amp; remarques</h3>
            <ul>
              <li>Le nom des courbes dans les tracés provient de <b>Sample description</b> dès que les métadonnées sont complètes.</li>
              <li>La correction de baseline (<i>modpoly</i>) est appliquée lors de la construction du fichier combiné.</li>
              <li>Vous pouvez trier les colonnes des tableaux en cliquant sur leurs en-têtes et utiliser les barres de défilement pour explorer toutes les lignes.</li>
              <li>Pour des performances optimales, évitez d’ouvrir des centaines de fichiers simultanément.</li>
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