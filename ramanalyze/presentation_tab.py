"""Onglet Présentation : le logo et une courte présentation du logiciel."""

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QScrollArea

_TEXTE = """
<p>Bienvenue dans Ramanalyze.</p>

<p>Ce logiciel sert à regarder des spectres Raman et à faire des titrations à
partir de ces spectres.</p>

<p><b>Visualiseur de spectres.</b> Ouvrir les fichiers et regarder les spectres
en grand.</p>

<p><b>Titration par suivi de pic.</b> Repérer les pics dans une zone, renseigner
le volume de titrant et la concentration, puis suivre comment l'intensité des
pics évolue. Possibilité de marquer le moment où le signal a fini de bouger.</p>

<p><b>Titration par rapport de pics.</b> Mesurer un pic ou le rapport de deux
pics, puis tracer la courbe de titration. Possibilité d'ajouter un ajustement
sigmoïde.</p>

<p>Les volumes et les séries sont partagés entre les onglets. Possibilité
d'enregistrer le tableau et de le recharger plus tard pour ne pas tout refaire.</p>
"""


class PresentationTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        content = QWidget(self)
        v = QVBoxLayout(content)
        v.setContentsMargins(40, 30, 40, 30)
        v.setSpacing(14)
        v.setAlignment(Qt.AlignTop | Qt.AlignHCenter)

        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")
        if os.path.exists(logo_path):
            logo = QLabel(content)
            logo.setPixmap(QPixmap(logo_path).scaledToHeight(150, Qt.SmoothTransformation))
            logo.setAlignment(Qt.AlignHCenter)
            v.addWidget(logo)

        title = QLabel("Ramanalyze", content)
        title.setAlignment(Qt.AlignHCenter)
        # pas de couleur fixe : on suit le thème (clair ou sombre) du système
        title.setStyleSheet("font-size: 26px; font-weight: 700;")
        v.addWidget(title)

        texte = QLabel(_TEXTE, content)
        texte.setTextFormat(Qt.RichText)
        texte.setWordWrap(True)
        texte.setMaximumWidth(640)
        texte.setStyleSheet("font-size: 15px;")
        v.addWidget(texte, 0, Qt.AlignHCenter)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setWidget(content)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)
