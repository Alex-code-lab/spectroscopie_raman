import os
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QListView,
    QFileSystemModel,
    QPushButton,
    QLabel,
    QLineEdit,
    QAbstractItemView,
)

# Dossier par défaut à l'ouverture (modifiez-le si besoin)
DEFAULT_DIR = os.path.expanduser("~/Documents/Travail/CitizenSers/Spectroscopie/AS003_532nm")


class FilePickerWidget(QWidget):
    """Widget de sélection de fichiers .txt avec navigation par double-clic.

    - Double-clic sur un dossier : entrer dans le dossier
    - Bouton "Monter" : remonter d'un niveau
    - Bouton "Ajouter la sélection" : ajoute les fichiers .txt sélectionnés à la liste interne
    - Bouton "Vider la liste" : vide la liste interne

    La liste des fichiers choisis est disponible via l'attribut
    `self.selected_files` ou la méthode `get_selected_files()`.
    """

    def __init__(self, parent=None, start_dir: str | None = None):
        super().__init__(parent)
        self.selected_files: list[str] = []
        root = start_dir or DEFAULT_DIR

        # --- Modèle système de fichiers ---
        self.model = QFileSystemModel(self)
        self.model.setRootPath(root)
        self.model.setNameFilters(["*.txt"])      # n'afficher que les .txt
        self.model.setNameFilterDisables(False)    # masquer les autres fichiers

        # --- Vue en mode icônes ---
        self.view = QListView(self)
        self.view.setModel(self.model)
        self.view.setRootIndex(self.model.index(root))
        self.view.setViewMode(QListView.IconMode)
        self.view.setResizeMode(QListView.Adjust)
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.doubleClicked.connect(self._on_double_clicked)

        # --- Barre de chemin + bouton Monter ---
        top = QHBoxLayout()
        self.path_edit = QLineEdit(root, self)
        self.path_edit.setReadOnly(True)
        btn_up = QPushButton("⬆︎ Monter", self)
        btn_up.clicked.connect(self._go_up)
        top.addWidget(QLabel("Dossier :", self))
        top.addWidget(self.path_edit, 1)
        top.addWidget(btn_up)

        # --- Boutons d'action ---
        btns = QHBoxLayout()
        self.btn_add = QPushButton("Ajouter la sélection", self)
        self.btn_add.clicked.connect(self._add_from_view)
        self.btn_clear = QPushButton("Vider la liste", self)
        self.btn_clear.clicked.connect(self.clear_selected)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_clear)

        self.info = QLabel("0 fichier sélectionné", self)

        # --- Layout principal ---
        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self.view, 1)
        layout.addLayout(btns)
        layout.addWidget(self.info)

    # Utilitaires
    def current_dir(self) -> str:
        return self.model.filePath(self.view.rootIndex())

    def _on_double_clicked(self, index: QModelIndex):
        path = self.model.filePath(index)
        if self.model.isDir(index):
            self.view.setRootIndex(index)
            self.path_edit.setText(path)
        # si c'est un fichier, on laisse la sélection multiple gérer l'ajout

    def _go_up(self):
        cur = self.view.rootIndex()
        parent = cur.parent()
        if parent.isValid():
            self.view.setRootIndex(parent)
            self.path_edit.setText(self.model.filePath(parent))

    def _add_from_view(self):
        # Ajoute tous les fichiers .txt sélectionnés à la liste interne
        indexes = self.view.selectionModel().selectedIndexes()
        newly: list[str] = []
        for idx in indexes:
            path = self.model.filePath(idx)
            if os.path.isfile(path) and path.lower().endswith(".txt"):
                newly.append(path)
        # dédoublonner en conservant l'ordre
        for p in newly:
            if p not in self.selected_files:
                self.selected_files.append(p)
        self._notify_count()

    def _notify_count(self):
        n = len(self.selected_files)
        self.info.setText(f"{n} fichier(s) prêt(s) à tracer")

    def clear_selected(self):
        self.selected_files.clear()
        self._notify_count()

    # Accès externe optionnel
    def get_selected_files(self) -> list[str]:
        return list(self.selected_files)