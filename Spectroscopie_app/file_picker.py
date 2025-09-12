import os
from PySide6.QtCore import Qt, QModelIndex, QSortFilterProxyModel, QDir
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
    QListWidget,
    QSizePolicy,
)

# Dossier par défaut à l'ouverture (modifiez-le si besoin)
DEFAULT_DIR = os.path.expanduser("~/Documents/Travail/CitizenSers/Spectroscopie/AS003_532nm")


class TxtAndDirsFilter(QSortFilterProxyModel):
    """
    Proxy qui accepte toujours les dossiers (pour pouvoir naviguer)
    et ne laisse passer que les fichiers .txt.
    """
    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        src = self.sourceModel()
        idx = src.index(source_row, 0, source_parent)
        if not idx.isValid():
            return False
        info = src.fileInfo(idx)
        if info.isDir():
            return True
        return info.suffix().lower() == "txt"


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

        # ---------- Modèle système de fichiers ----------
        self.model = QFileSystemModel(self)
        self.model.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot)
        root_src_index = self.model.setRootPath(root)

        # ---------- Proxy : dossiers + *.txt ----------
        self.proxy = TxtAndDirsFilter(self)
        self.proxy.setSourceModel(self.model)

        # ---------- Vue (navigateur) ----------
        self.view = QListView(self)
        self.view.setModel(self.proxy)
        self.view.setRootIndex(self.proxy.mapFromSource(root_src_index))
        self.view.setViewMode(QListView.IconMode)
        self.view.setResizeMode(QListView.Adjust)
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.doubleClicked.connect(self._on_double_clicked)

        # ---------- Instructions haut de page ----------
        instructions = QLabel(
            "<b>Sélection des spectres</b><br>"
            "Utilisez cette page pour naviguer et sélectionner vos fichiers .txt de spectroscopie Raman.<br>"
            "<ul>"
            "<li>Double-cliquez sur un dossier pour y entrer</li>"
            "<li>Sélectionnez un ou plusieurs fichiers .txt</li>"
            "<li>Cliquez sur <i>Ajouter la sélection</i> pour les inclure</li>"
            "<li>Le panneau de droite affiche vos fichiers sélectionnés</li>"
            "</ul>",
            self,
        )
        instructions.setWordWrap(True)
        instructions.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

        # ---------- Barre de chemin + bouton Monter ----------
        path_bar = QHBoxLayout()
        self.path_edit = QLineEdit(root, self)
        self.path_edit.setReadOnly(True)
        btn_up = QPushButton("⬆︎ Monter", self)
        btn_up.clicked.connect(self._go_up)
        path_bar.addWidget(QLabel("Dossier :", self))
        path_bar.addWidget(self.path_edit, 1)
        path_bar.addWidget(btn_up)

        # ---------- Panneau gauche : navigateur + bouton Ajouter ----------
        nav_layout = QVBoxLayout()
        nav_layout.addWidget(self.view, 1)
        self.btn_add = QPushButton("Ajouter la sélection", self)
        self.btn_add.clicked.connect(self._add_from_view)
        nav_layout.addWidget(self.btn_add)

        # ---------- Panneau droit : liste sélectionnée + boutons ----------
        self.selected_list_widget = QListWidget(self)
        self.selected_list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)

        sel_layout = QVBoxLayout()
        sel_layout.addWidget(QLabel("Fichiers sélectionnés", self))
        sel_layout.addWidget(self.selected_list_widget, 1)
        self.btn_remove = QPushButton("Retirer", self)
        self.btn_remove.clicked.connect(self._remove_selected_from_list)
        self.btn_clear = QPushButton("Vider la liste", self)
        self.btn_clear.clicked.connect(self.clear_selected)
        sel_btns = QHBoxLayout()
        sel_btns.addWidget(self.btn_remove)
        sel_btns.addWidget(self.btn_clear)
        sel_layout.addLayout(sel_btns)

        # ---------- Bloc central : côte à côte, même hauteur ----------
        center_layout = QHBoxLayout()
        center_layout.addLayout(nav_layout, 3)  # navigateur plus large
        center_layout.addLayout(sel_layout, 2)  # sélection un peu plus étroite

        # ---------- Bar d'info bas de page ----------
        self.info = QLabel("0 fichier sélectionné", self)

        # ---------- Layout principal ----------
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(instructions)
        main_layout.addLayout(path_bar)
        main_layout.addLayout(center_layout, 1)
        main_layout.addWidget(self.info)

    # Utilitaires
    def _current_src_root(self) -> QModelIndex:
        """Renvoie l'index source (QFileSystemModel) correspondant au root affiché dans la vue."""
        # root index actuel côté proxy
        root_proxy = self.view.rootIndex()
        # le mapper en source pour manipuler parents/chemins
        return self.proxy.mapToSource(root_proxy)

    def _set_root_from_src_index(self, src_index: QModelIndex):
        """Met à jour la vue (proxy) pour afficher le src_index donné comme racine."""
        self.view.setRootIndex(self.proxy.mapFromSource(src_index))
        self.path_edit.setText(self.model.filePath(src_index))

    def _on_double_clicked(self, proxy_index: QModelIndex):
        src_index = self.proxy.mapToSource(proxy_index)
        if self.model.isDir(src_index):
            self._set_root_from_src_index(src_index)
        # si c'est un fichier, on laisse la sélection multiple gérer l'ajout

    def _go_up(self):
        src_cur = self._current_src_root()
        parent = src_cur.parent()
        if parent.isValid():
            self._set_root_from_src_index(parent)

    def _add_from_view(self):
        # Ajoute tous les fichiers .txt sélectionnés à la liste interne
        indexes = self.view.selectionModel().selectedIndexes()
        newly: list[str] = []
        for proxy_idx in indexes:
            src_idx = self.proxy.mapToSource(proxy_idx)
            path = self.model.filePath(src_idx)
            if os.path.isfile(path) and path.lower().endswith(".txt"):
                newly.append(path)
        # dédoublonner en conservant l'ordre
        for p in newly:
            if p not in self.selected_files:
                self.selected_files.append(p)
        self._notify_count()
        self._refresh_selected_list()

    def _notify_count(self):
        n = len(self.selected_files)
        self.info.setText(f"{n} fichier(s) prêt(s) à tracer")

    def clear_selected(self):
        self.selected_files.clear()
        self._notify_count()
        self.selected_list_widget.clear()

    def _refresh_selected_list(self):
        self.selected_list_widget.clear()
        for file_path in self.selected_files:
            self.selected_list_widget.addItem(file_path)

    def _remove_selected_from_list(self):
        selected_items = self.selected_list_widget.selectedItems()
        for item in selected_items:
            file_path = item.text()
            if file_path in self.selected_files:
                self.selected_files.remove(file_path)
            self.selected_list_widget.takeItem(self.selected_list_widget.row(item))
        self._notify_count()

    # Accès externe optionnel
    def get_selected_files(self) -> list[str]:
        return list(self.selected_files)