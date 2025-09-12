from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel


class AnalysisTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Espace réservé pour l'analyse (à implémenter)"))
        self.setLayout(layout)