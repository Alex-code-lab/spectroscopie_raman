import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabBar,
    QTabWidget,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QStatusBar,
    QLabel,
    QScrollArea,
    QPushButton,
    QDialog,
    QMessageBox,
    QStyle,
    QStyleOptionTab,
    QStylePainter,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QIcon, QPalette, QPixmap

# S'assurer que les imports de modules frères fonctionnent quel que soit l'endroit
# d'où on lance le script.
# En mode PyInstaller onefile, sys._MEIPASS pointe vers le dossier d'extraction
# temporaire où tous les fichiers data sont copiés.
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    APP_DIR = sys._MEIPASS
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from file_picker import FilePickerWidget
from spectra_plot import SpectraTab
from analysis_tab import AnalysisTab
from metadata_creator import MetadataCreatorWidget


class WorkflowStatusTabBar(QTabBar):
    """QTabBar avec fond rouge/vert pour certains onglets de workflow."""

    _ACTIVE_EXTRA_WIDTH = 24
    _ACTIVE_EXTRA_HEIGHT = 8
    _ACTIVE_BORDER_COLOR = "#f4f4f4"

    def __init__(self, parent=None):
        super().__init__(parent)
        self._status_colors: dict[int, str] = {}
        self.currentChanged.connect(self._refresh_active_tab)

    def tabSizeHint(self, index: int):
        size = super().tabSizeHint(index)
        if index == self.currentIndex():
            size.setWidth(size.width() + self._ACTIVE_EXTRA_WIDTH)
            size.setHeight(size.height() + self._ACTIVE_EXTRA_HEIGHT)
        return size

    def _refresh_active_tab(self, index: int) -> None:
        self.updateGeometry()
        self.update()

    def set_tab_status_color(self, index: int, color_hex: str | None) -> None:
        if color_hex:
            self._status_colors[index] = color_hex
        else:
            self._status_colors.pop(index, None)
        if 0 <= index < self.count():
            self.update(self.tabRect(index))

    def paintEvent(self, event) -> None:
        painter = QStylePainter(self)

        for index in range(self.count()):
            option = QStyleOptionTab()
            self.initStyleOption(option, index)
            color_hex = self._status_colors.get(index)
            if not color_hex:
                if option.state & QStyle.StateFlag.State_Selected:
                    self._draw_colored_tab(painter, option, QColor("#666666"))
                else:
                    painter.drawControl(QStyle.ControlElement.CE_TabBarTab, option)
                continue

            color = QColor(color_hex)
            if option.state & QStyle.StateFlag.State_Selected:
                color = color.lighter(115)

            self._draw_colored_tab(painter, option, color)

    def _draw_colored_tab(self, painter: QStylePainter, option: QStyleOptionTab, color: QColor) -> None:
        rect = option.rect.adjusted(1, 1, -1, -1)
        painter.save()
        painter.setPen(color.darker(125))
        painter.setBrush(color)
        painter.drawRoundedRect(rect, 4, 4)
        painter.restore()

        if option.state & QStyle.StateFlag.State_Selected:
            self._draw_active_frame(painter, option)

        self._draw_tab_label(
            painter,
            option,
            QColor("#ffffff"),
            bold=bool(option.state & QStyle.StateFlag.State_Selected),
        )

    def _draw_active_frame(self, painter: QStylePainter, option: QStyleOptionTab) -> None:
        rect = option.rect.adjusted(0, 0, -1, -1)
        painter.save()
        painter.setPen(QColor(self._ACTIVE_BORDER_COLOR))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(rect, 5, 5)
        painter.restore()

    def _draw_tab_label(
        self,
        painter: QStylePainter,
        option: QStyleOptionTab,
        text_color: QColor | None = None,
        *,
        bold: bool = False,
    ) -> None:
        if text_color is not None:
            option.palette.setColor(QPalette.ColorRole.WindowText, text_color)
            option.palette.setColor(QPalette.ColorRole.ButtonText, text_color)
            option.palette.setColor(QPalette.ColorRole.Text, text_color)

        painter.save()
        if bold:
            font = painter.font()
            font.setBold(True)
            painter.setFont(font)
        painter.drawControl(QStyle.ControlElement.CE_TabBarTabLabel, option)
        painter.restore()


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
        tabs.setTabBar(WorkflowStatusTabBar(tabs))
        self.tabs = tabs
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
            padding: 8px 12px; margin: 12px 0; border-radius: 0 4px 4px 0; color: #e0e0e0; }
  .warn-box { border-left: 4px solid #e07b00; background: #2a1f10;
              padding: 8px 12px; margin: 12px 0; border-radius: 0 4px 4px 0; color: #e8d5b0; }
  table { border-collapse: collapse; margin: 8px 0; }
  td, th { border: 1px solid #444; padding: 5px 10px; font-size: 0.93em; }
  th { background: #2a2a2a; color: #d4d4d4; font-weight: 600; }
</style>

<h2>Dispositif de mesure de la qualité de l'eau par les citoyens</h2>

<p>
  <span class="key">Ramanalyze</span> centralise la traçabilité du prélèvement,
  les mesures terrain, les analyses complémentaires et, si elle est réalisée,
  la titration suivie par spectroscopie Raman/SERS.
</p>

<div class="note">
  <b style="color:#ffffff">Guidage visuel :</b>
  les champs <span class="warn">rouges</span> sont à renseigner, les champs
  <span class="ok">verts</span> sont remplis. Les boutons rouges signalent une
  étape à préparer ou à valider ; les boutons verts indiquent une étape prête.
</div>

<h3>Parcours recommandé</h3>
<ol>
  <li><span class="key">Métadonnées</span> - créez ou rechargez la fiche terrain, puis enregistrez le fichier <span class="mono">.xlsx</span>.</li>
  <li><span class="key">Fichiers Raman</span> - ajoutez les fichiers <span class="mono">.txt</span> si une acquisition Raman/SERS a été réalisée.</li>
  <li><span class="key">Spectres</span> - contrôlez visuellement les spectres et la correction de baseline.</li>
  <li><span class="key">Analyse</span> - choisissez la longueur d'onde du spectromètre, ajustez les ratios et lancez l'ajustement mathématique de la courbe.</li>
  <li><span class="key">Exploration</span> - utilisez la PCA et les graphes de comparaison pour comprendre les variations entre spectres.</li>
</ol>

<h3>1 - Onglet Métadonnées</h3>

<h4>Charger ou créer une fiche</h4>
<p>
  Le bouton <b>Charger des métadonnées...</b> est placé en haut pour reprendre
  une fiche existante. Sinon, remplissez les sections de haut en bas.
  Le bouton <b>Enregistrer les métadonnées...</b> reste en bas de page et propose
  automatiquement un nom de fichier basé sur le nom du prélèvement :
  <span class="mono">NomPrelevement_T_AnBa_TAm.xlsx</span>.
</p>
<ul>
  <li><span class="mono">_T</span> est ajouté si la titration du cuivre est cochée.</li>
  <li><span class="mono">_AnBa</span> est ajouté si les analyses bactériologiques sont cochées.</li>
  <li><span class="mono">_TAm</span> est ajouté si le test ammonium est coché.</li>
  <li><span class="mono">_Tur</span> est ajouté si la turbidité est mesurée.</li>
  <li><span class="mono">_Cond</span> est ajouté si la conductivité est mesurée.</li>
  <li><span class="mono">_pH</span> est ajouté si le pH de l'eau est mesuré.</li>
  <li><span class="mono">_Ox</span> est ajouté si l'oxygène dissous est mesuré.</li>
</ul>

<div class="note">
  <b style="color:#ffffff">Nom du prélèvement — identifiant de l'échantillon :</b>
  Le champ <span class="key">Nom du prélèvement</span> est l'identifiant central de toute la session.
  Il est généré automatiquement à partir des initiales du·de la préleveur·se et de la date/heure,
  mais peut être modifié librement. <b>C'est ce nom qui apparaîtra dans tous les exports et
  dans la correspondance spectres ↔ tubes — assurez-vous qu'il est unique et explicite
  avant d'enregistrer</b> (ex. : <span class="mono">DJ_2026-04-30_14h00</span>).
  Traitez-le comme le code-barres de votre échantillon.
</div>

<h4>Prélèvement de l'échantillon</h4>
<ul>
  <li>Chaque préleveur·se est saisi·e au format <b>Prénom puis nom</b>, avec son association sur la même ligne.</li>
  <li>Le bouton <b>+1 préleveur·se</b> ajoute une personne et une association supplémentaires.</li>
  <li>La date et l'heure du prélèvement sont vides au départ ; elles passent au vert lorsqu'elles sont renseignées.</li>
  <li>Le <b>nom du prélèvement</b> est généré automatiquement avec les initiales de tous les préleveur·ses, puis la date et l'heure.</li>
  <li>Latitude et longitude acceptent les degrés décimaux ou les coordonnées GPS issues d'une carte.</li>
  <li>Le bouton <b>Carte / pointer...</b> permet de placer le point sur une carte interne quand le composant web est disponible.</li>
  <li>Si Internet est disponible, département et commune sont proposés automatiquement depuis les coordonnées GPS ; ils restent modifiables à la main.</li>
</ul>

<h4>Contexte terrain</h4>
<ul>
  <li><b>Type d'eau</b> reste rouge tant qu'aucun choix n'est fait.</li>
  <li>Pour <b>eau de mer</b> ou <b>eau estuarienne</b>, le coefficient de marée et l'heure de pleine mer apparaissent.</li>
  <li>Pour <b>eau douce</b>, le bouton <b>Débit / flux...</b> ouvre le tableau largeur, hauteur d'eau, volume, vitesse et débit.</li>
  <li>Température de l'eau, météo, température de l'air, pluie sur 24 h et commentaire décrivent les conditions de prélèvement.</li>
</ul>

<h4>Mesures effectuées</h4>
<ul>
  <li><b>Test ammonium réalisé</b> ouvre une fenêtre pour le test grossier, le test précis, le pH et la personne qui réalise le test.</li>
  <li><b>Analyses bactériologiques</b> affiche le jour de dépôt, l'heure de dépôt, l'entreprise de mesure et les résultats E. coli / entérocoques si disponibles.</li>
  <li><b>Titration</b> est cochée uniquement si une titration est prévue ou réalisée. Elle reste en dessous des autres mesures.</li>
</ul>

<h4>Information sur la titration</h4>
<ul>
  <li>Coordinateur·ice et opérateur·ice sont saisi·es au format <b>Prénom puis nom</b>.</li>
  <li>Lieu, date et heure de titration génèrent le nom de la manip de titration.</li>
  <li>La correspondance spectres ↔ tubes repasse rouge uniquement si une information de titration est modifiée.</li>
</ul>

<h4>Tableau de volumes et protocole</h4>
<p>
  Le modèle actuel est la titration classique compte-gouttes. Le bouton
  <b>Générer un tableau de volume...</b> reste disponible pour les usages avancés.
  La feuille de protocole devient verte quand le tableau des volumes est prêt.
</p>
<table>
  <tr><th>Solution</th><th>Rôle</th></tr>
  <tr><td><b>Solution A1</b></td><td>Tampon NH3+ concentré, volume fixe.</td></tr>
  <tr><td><b>Solution A2</b></td><td>Tampon NH3+ dilué, complémentaire à la solution B.</td></tr>
  <tr><td><b>Solution B</b></td><td>Titrant, volume variable d'un tube à l'autre.</td></tr>
  <tr><td><b>Solution C</b></td><td>Indicateur SERS.</td></tr>
  <tr><td><b>Solution D</b></td><td>PEG ou agent de crowding.</td></tr>
  <tr><td><b>Solution E</b></td><td>Nanoparticules SERS.</td></tr>
  <tr><td><b>Solution F</b></td><td>Crosslinker, à mélanger immédiatement tube par tube.</td></tr>
</table>
<p>
  Dans le protocole guidé, les cases qui ne sont pas à l'étape courante sont
  grisées. Les étapes actives passent en couleur. Des lignes de mélange sont
  prévues après les solutions B, C, D et E, puis une attention spécifique est
  affichée avant l'ajout de la solution F.
</p>

<h4>Correspondance spectres ↔ tubes</h4>
<p>
  La correspondance associe les noms de spectres aux tubes. Elle utilise le nom
  du prélèvement pour générer les noms de spectres et conserve les associations
  déjà définies lorsque les autres métadonnées terrain sont modifiées.
</p>

<h4>Fichier Excel enregistré</h4>
<p>
  L'enregistrement est possible même sans titration. Si des cases optionnelles
  sont cochées mais incomplètes, un avertissement signale les points à vérifier
  sans bloquer l'enregistrement.
</p>
<ul>
  <li><span class="mono">Mesures</span> - fiche terrain synthétique, avec une ligne par préleveur·se et les résultats de mesures.</li>
  <li><span class="mono">Index-métadonnées</span> - feuille technique masquée utilisée pour recharger les champs de façon robuste.</li>
  <li><span class="mono">Titration-volumes</span> - tableau de volumes si la titration est utilisée.</li>
  <li><span class="mono">Titration-correspondance</span> - métadonnées et correspondance spectres ↔ tubes.</li>
  <li><span class="mono">Titration-EtatProtocole</span> - état des cases du protocole si une feuille de protocole a été utilisée.</li>
</ul>

<h3>2 - Onglet Fichiers Raman</h3>
<ul>
  <li>Sélectionnez les fichiers <span class="mono">.txt</span> issus du spectromètre.</li>
  <li>Cliquez <b>Ajouter la sélection</b> pour les placer dans la liste de travail.</li>
  <li>Utilisez <b>Retirer</b> ou <b>Vider la liste</b> si nécessaire.</li>
</ul>

<h3>3 - Onglet Spectres</h3>
<ul>
  <li>Affiche les spectres corrigés et permet de contrôler rapidement les profils.</li>
  <li>La visualisation interactive permet de zoomer, comparer et exporter une figure.</li>
  <li>Un résultat incohérent doit d'abord faire vérifier la correspondance spectres ↔ tubes et la baseline.</li>
</ul>

<h3>4 - Onglet Analyse</h3>
<ul>
  <li><b>Longueur d'onde du spectromètre</b> choisit les pics adaptés au jeu spectral.</li>
  <li><b>Ajuster le ratio de longueur d'onde</b> permet de choisir le ratio utilisé pour la quantification.</li>
  <li><b>Ajustement mathématique de la courbe</b> lance le modèle d'ajustement sur les données disponibles.</li>
  <li>L'export Excel contient les intensités, ratios et résultats utiles à la comparaison.</li>
</ul>

<h3>5 - Onglet Exploration</h3>
<table>
  <tr><th>Sous-onglet</th><th>Ce qu'il montre</th></tr>
  <tr><td>Spectres normalisés</td><td>Spectres après normalisation par la norme vectorielle.</td></tr>
  <tr><td>Spectres superposés</td><td>Comparaison brute des spectres.</td></tr>
  <tr><td>PCA corrélée</td><td>PCA sur les intensités aux pics sélectionnés.</td></tr>
  <tr><td>PCA - Scores</td><td>Position des spectres dans l'espace PCA, colorée par variable.</td></tr>
  <tr><td>PCA - Loadings</td><td>Bandes Raman qui portent chaque composante.</td></tr>
  <tr><td>PCA - Reconstruction</td><td>Reconstruction d'un spectre à partir d'un nombre choisi de composantes.</td></tr>
</table>

<div class="warn-box">
  <b style="color:#ffffff">Points à vérifier en cas d'incohérence :</b>
  nom du prélèvement, date/heure, coordonnées GPS, type d'eau, correspondance
  spectres ↔ tubes, tableau de volumes, baseline et choix de longueur d'onde.
</div>
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

        meta_top_bar = QHBoxLayout()
        meta_top_bar.addStretch(1)
        self.btn_new_metadata = QPushButton("Nouvelle fiche", self)
        self.btn_new_metadata.setStyleSheet(
            "background-color: #6c757d; color: white; font-weight: 700; padding: 4px 14px;"
        )
        self.btn_new_metadata.setToolTip(
            "Vider tous les champs et repartir d'une fiche vierge."
        )
        self.btn_new_metadata.clicked.connect(self._on_new_metadata_clicked)
        meta_top_bar.addWidget(self.btn_new_metadata)
        metadata_layout.addLayout(meta_top_bar)

        metadata_scroll = QScrollArea(self.metadata_tab)
        metadata_scroll.setWidgetResizable(True)
        metadata_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.metadata_creator = MetadataCreatorWidget(metadata_scroll)
        metadata_scroll.setWidget(self.metadata_creator)
        metadata_layout.addWidget(metadata_scroll)
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
        self._copper_titration_tabs = [
            self.file_tab,
            self.spectra_tab,
            self.analysis_tab,
            self.peak_selector_tab,
        ]
        self.metadata_creator.titration_visibility_changed.connect(
            self._refresh_copper_titration_tabs_visible
        )
        self.metadata_creator.metadata_saved_status_changed.connect(
            self._on_metadata_status_changed
        )
        self._tab_status = {
            self.metadata_tab: not self.metadata_creator.has_unsaved_metadata(),
            self.file_tab: False,
            self.spectra_tab: False,
            self.analysis_tab: False,
        }
        self.file_picker.selection_changed.connect(self._on_file_selection_changed)
        self.spectra_tab.plot_status_changed.connect(self._on_spectra_status_changed)
        self.analysis_tab.analysis_status_changed.connect(self._on_analysis_status_changed)
        self._refresh_workflow_tab_statuses()
        self._refresh_copper_titration_tabs_visible(
            self.metadata_creator.chk_titration_done.isChecked()
        )
        self._last_tab_index = tabs.currentIndex()
        tabs.currentChanged.connect(self._on_tab_changed)

        # --- Label des sources en bas ---
        self.sources_label = QLabel("L'ensemble des sources sont à retrouver <a href='#'>ici</a>.")
        self.sources_label.setTextFormat(Qt.RichText)
        self.sources_label.setOpenExternalLinks(False)
        self.sources_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.LinksAccessibleByMouse)
        self.sources_label.setAlignment(Qt.AlignCenter)
        self.sources_label.linkActivated.connect(self.show_sources_popup)
        central_layout.addWidget(self.sources_label)

    def _set_workflow_tab_status(self, widget: QWidget, ready: bool) -> None:
        idx = self.tabs.indexOf(widget)
        if idx < 0:
            return
        color = "#5cb85c" if ready else "#d9534f"
        tab_bar = self.tabs.tabBar()
        if hasattr(tab_bar, "set_tab_status_color"):
            tab_bar.set_tab_status_color(idx, color)
        else:
            tab_bar.setTabTextColor(idx, QColor("#ffffff"))
        self.tabs.setTabIcon(idx, QIcon())

    def _refresh_workflow_tab_statuses(self) -> None:
        for widget, ready in getattr(self, "_tab_status", {}).items():
            self._set_workflow_tab_status(widget, ready)

    def _on_metadata_status_changed(self, saved: bool) -> None:
        self._tab_status[self.metadata_tab] = bool(saved)
        self._set_workflow_tab_status(self.metadata_tab, bool(saved))

    def _on_file_selection_changed(self, has_files: bool) -> None:
        self._tab_status[self.file_tab] = bool(has_files)
        self._set_workflow_tab_status(self.file_tab, bool(has_files))
        if hasattr(self, "spectra_tab"):
            self.spectra_tab.mark_plot_stale()
        self._on_spectra_status_changed(False)
        if hasattr(self, "analysis_tab"):
            self.analysis_tab.mark_analysis_stale()
        self._on_analysis_status_changed(False)

    def _on_spectra_status_changed(self, plotted: bool) -> None:
        self._tab_status[self.spectra_tab] = bool(plotted)
        self._set_workflow_tab_status(self.spectra_tab, bool(plotted))
        if not plotted and hasattr(self, "analysis_tab"):
            self.analysis_tab.mark_analysis_stale()
            self._on_analysis_status_changed(False)

    def _on_analysis_status_changed(self, analyzed: bool) -> None:
        self._tab_status[self.analysis_tab] = bool(analyzed)
        self._set_workflow_tab_status(self.analysis_tab, bool(analyzed))

    def _refresh_copper_titration_tabs_visible(self, visible: bool) -> None:
        """Affiche les onglets utiles uniquement à la titration du cuivre."""
        visible = bool(visible)
        titration_tabs = getattr(self, "_copper_titration_tabs", [])

        if not visible and self.tabs.currentWidget() in titration_tabs:
            metadata_index = self.tabs.indexOf(self.metadata_tab)
            if metadata_index >= 0:
                self.tabs.setCurrentIndex(metadata_index)

        for widget in titration_tabs:
            idx = self.tabs.indexOf(widget)
            if idx < 0:
                continue
            self.tabs.setTabVisible(idx, visible)

        self._refresh_workflow_tab_statuses()
        self._last_tab_index = self.tabs.currentIndex()

    def _on_tab_changed(self, index: int) -> None:
        previous_widget = self.tabs.widget(getattr(self, "_last_tab_index", index))
        if (
            previous_widget is self.metadata_tab
            and hasattr(self, "metadata_creator")
            and self.metadata_creator.has_unsaved_metadata()
        ):
            QMessageBox.warning(
                self,
                "Métadonnées non enregistrées",
                "Les métadonnées ont été modifiées et ne sont pas encore enregistrées.\n"
                "Vous pouvez continuer, mais pensez à enregistrer avant de quitter le logiciel.",
            )
        self._last_tab_index = index

    def _on_new_metadata_clicked(self) -> None:
        mc = getattr(self, "metadata_creator", None)
        if mc is None:
            return
        has_data = (
            mc.df_comp is not None
            or mc.df_map is not None
            or any(
                bool(getattr(mc, attr, None) and getattr(mc, attr).text().strip())
                for attr in (
                    "edit_sampler", "edit_sample_location",
                    "edit_sample_lat", "edit_sample_lon",
                    "edit_coordinator", "edit_operator",
                )
            )
            or (
                hasattr(mc, "edit_sample_date")
                and mc.edit_sample_date.date() != mc.edit_sample_date.minimumDate()
            )
        )
        if has_data:
            reply = QMessageBox.question(
                self,
                "Nouvelle fiche",
                "Toutes les données saisies seront effacées.\n"
                "Voulez-vous vraiment repartir d'une fiche vierge ?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
        mc.reset_all()
        fp = getattr(self, "file_picker", None)
        if fp is not None:
            fp.clear_selected()

    def closeEvent(self, event):
        mc = getattr(self, "metadata_creator", None)
        if mc is not None and mc.has_unsaved_metadata():
            reply = QMessageBox.question(
                self,
                "Quitter sans sauvegarder ?",
                "Des métadonnées ont été modifiées et ne sont pas encore enregistrées.\n"
                "Voulez-vous vraiment quitter ?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()

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
            <li><b><a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a></b><br>Carte et données cartographiques utilisées pour le pointage et le géocodage.</li>
            <li><b><a href="https://nominatim.org/">Nominatim</a></b><br>Service de géocodage inverse utilisé pour proposer automatiquement département et commune.</li>
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
