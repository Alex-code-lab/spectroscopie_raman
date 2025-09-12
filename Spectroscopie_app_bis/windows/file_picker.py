import os
from PySide6.QtCore import Qt, QSortFilterProxyModel, QModelIndex, QDir, QSize
from PySide6.QtGui import QStandardItemModel
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QListView, QListWidget, QListWidgetItem, QPushButton, QFileSystemModel,
    QAbstractItemView, QSplitter, QMessageBox, QToolBar, QStatusBar, QStyle
)

class TxtAndDirsFilter(QSortFilterProxyModel):
    """
    Filtre : autorise tous les dossiers, et uniquement les fichiers .txt
    """
    def filterAcceptsRow(self, source_row, source_parent):
        idx = self.sourceModel().index(source_row, 0, source_parent)
        if not idx.isValid():
            return False
        # Dossiers : toujours OK
        if self.sourceModel().isDir(idx):
            return True
        # Fichiers .txt uniquement
        name = self.sourceModel().fileName(idx)
        return name.lower().endswith(".txt")


class FilePickerWindow(QMainWindow):
    """
    Fenêtre principale qui affiche :
      - à gauche : navigateur fichiers (icônes), double-clic pour entrer dans un dossier
      - en haut : un bandeau d’instructions + chemin courant + bouton 'Monter'
      - à droite : liste des .txt sélectionnés (ajout via bouton ou double-clic fichier)
    """

    def __init__(self, default_dir=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sélection des fichiers spectres (.txt)")
        self.resize(1100, 700)

        # --- état interne -----------------------------------------------------
        self.selected_paths = []  # liste de chemins absolus

        # --- barre d’état -----------------------------------------------------
        self.setStatusBar(QStatusBar(self))

        # --- toolbar (navigation simple) -------------------------------------
        tb = QToolBar("Navigation", self)
        self.addToolBar(tb)

        self.path_edit = QLineEdit(self)
        self.path_edit.setReadOnly(True)
        self.path_edit.setMinimumWidth(420)

        self.btn_up = QPushButton("Monter")
        self.btn_up.clicked.connect(self.navigate_up)

        tb.addWidget(QLabel("Chemin : "))
        tb.addWidget(self.path_edit)
        tb.addSeparator()
        tb.addWidget(self.btn_up)

        # --- modèle fichiers + proxy filtre ----------------------------------
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setRootPath("")  # requis avant setRootIndex
        self.fs_model.setFilter(QDir.AllDirs | QDir.NoDotAndDotDot | QDir.Files)

        self.proxy = TxtAndDirsFilter(self)
        self.proxy.setSourceModel(self.fs_model)

        # --- vue fichiers (gauche) -------------------------------------------
        self.view = QListView(self)
        self.view.setViewMode(QListView.IconMode)          # icônes type Finder
        self.view.setResizeMode(QListView.Adjust)
        self.view.setUniformItemSizes(True)
        # Définir une taille d'icône correcte (PM_IconViewIconSize est portable)
        pm = self.style().pixelMetric(QStyle.PM_IconViewIconSize)
        if pm <= 0:
            pm = 64  # valeur par défaut raisonnable
        self.view.setIconSize(QSize(pm, pm))
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setModel(self.proxy)
        self.view.doubleClicked.connect(self.on_double_click)

        # --- panneau droit : sélection ---------------------------------------
        right_panel = QWidget(self)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.setSpacing(8)

        info = QLabel(
            "<b>Instructions :</b><br>"
            "• Naviguez dans vos dossiers (double-clic pour entrer).<br>"
            "• Seuls les fichiers <b>.txt</b> sont affichés côté fichiers.<br>"
            "• Sélectionnez des .txt dans la vue de gauche puis cliquez "
            "<i>Ajouter</i> — ou double-cliquez directement sur un fichier .txt.<br>"
            "• La liste de droite récapitule les fichiers choisis (clic pour retirer)."
        )
        info.setWordWrap(True)
        right_layout.addWidget(info)

        self.btn_add = QPushButton("Ajouter →")
        self.btn_remove = QPushButton("Retirer")
        self.btn_clear = QPushButton("Vider la sélection")
        btn_row = QHBoxLayout()
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        right_layout.addLayout(btn_row)

        self.selected_list = QListWidget(self)
        self.selected_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        right_layout.addWidget(self.selected_list)

        # --- splitter central -------------------------------------------------
        splitter = QSplitter(self)
        splitter.addWidget(self.view)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        self.setCentralWidget(splitter)

        # --- signaux boutons --------------------------------------------------
        self.btn_add.clicked.connect(self.add_from_view)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.clear_all)
        self.selected_list.itemDoubleClicked.connect(self.remove_one)

        # --- initialisation du répertoire racine ------------------------------
        start_dir = default_dir or os.path.expanduser("~")
        self.set_root_dir(start_dir)

        # message initial
        self._notify()

    # ------------------------------------------------------------------ utils
    def set_root_dir(self, path: str):
        """Positionne la vue sur un répertoire donné (absolu)."""
        path = os.path.abspath(path)
        if not os.path.isdir(path):
            path = os.path.expanduser("~")

        src_idx = self.fs_model.index(path)
        if not src_idx.isValid():
            return

        proxy_idx = self.proxy.mapFromSource(src_idx)
        self.view.setRootIndex(proxy_idx)
        self.path_edit.setText(path)

    def current_dir(self) -> str:
        """Retourne le chemin absolu du dossier affiché."""
        proxy_root = self.view.rootIndex()
        src_root = self.proxy.mapToSource(proxy_root)
        return self.fs_model.filePath(src_root)

    # -------------------------------------------------------------- interactions
    def navigate_up(self):
        """Remonte d’un niveau dans l’arborescence."""
        parent_dir = os.path.dirname(self.current_dir())
        if parent_dir and os.path.isdir(parent_dir):
            self.set_root_dir(parent_dir)

    def on_double_click(self, proxy_index: QModelIndex):
        """
        Double-clic :
          - si dossier → entrer dedans
          - si fichier .txt → l’ajouter à la sélection
        """
        src_idx = self.proxy.mapToSource(proxy_index)
        if self.fs_model.isDir(src_idx):
            self.set_root_dir(self.fs_model.filePath(src_idx))
            return

        path = self.fs_model.filePath(src_idx)
        if path.lower().endswith(".txt"):
            self._add_path(path)

    def add_from_view(self):
        """Ajoute tous les .txt actuellement sélectionnés dans la vue de gauche."""
        indexes = self.view.selectionModel().selectedIndexes()
        if not indexes:
            QMessageBox.information(self, "Info", "Sélectionnez d’abord un ou plusieurs fichiers .txt.")
            return

        added = 0
        for proxy_idx in indexes:
            src_idx = self.proxy.mapToSource(proxy_idx)
            if not self.fs_model.isDir(src_idx):
                path = self.fs_model.filePath(src_idx)
                if path.lower().endswith(".txt"):
                    added += self._add_path(path)

        if added == 0:
            QMessageBox.information(self, "Info", "Aucun nouveau fichier .txt n’a été ajouté.")
        self._notify()

    def _add_path(self, path: str) -> int:
        """Ajoute un chemin s’il n’est pas déjà dans la sélection."""
        path = os.path.abspath(path)
        if path in self.selected_paths:
            return 0
        self.selected_paths.append(path)
        item = QListWidgetItem(path)
        item.setToolTip(path)
        self.selected_list.addItem(item)
        return 1

    def remove_selected(self):
        """Retire les éléments cochés de la liste de droite."""
        for item in self.selected_list.selectedItems():
            path = item.text()
            if path in self.selected_paths:
                self.selected_paths.remove(path)
            self.selected_list.takeItem(self.selected_list.row(item))
        self._notify()

    def remove_one(self, item: QListWidgetItem):
        """Double-clic sur un item → retrait."""
        path = item.text()
        if path in self.selected_paths:
            self.selected_paths.remove(path)
        self.selected_list.takeItem(self.selected_list.row(item))
        self._notify()

    def clear_all(self):
        """Vide complètement la sélection."""
        self.selected_paths.clear()
        self.selected_list.clear()
        self._notify()

    # -------------------------------------------------------------- feedback UI
    def _notify(self):
        n = len(self.selected_paths)
        self.statusBar().showMessage(f"{n} fichier(s) .txt sélectionné(s) — chemin : {self.current_dir()}")