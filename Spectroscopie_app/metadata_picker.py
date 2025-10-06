# metadata_picker.py
import os
import pandas as pd
from data_processing import build_combined_dataframe
import tempfile
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileSystemModel, QListView, QAbstractItemView,
    QMessageBox, QTableWidget, QTableWidgetItem, QSizePolicy, QHeaderView, QFileDialog
)
from PySide6.QtCore import QDir, QModelIndex

# Dossier par défaut à l'ouverture (modifiez-le si besoin)
DEFAULT_DIR = os.path.expanduser("~/Documents/Travail/CitizenSers/Spectroscopie/")

class MetadataPickerWidget(QWidget):
    def __init__(self, parent=None, start_dir: str | None = None):
        super().__init__(parent)

        self.metadata_path = None
        self.df = None
        root = start_dir or os.path.expanduser("~")

        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("<b>Sélection du fichier de métadonnées (Excel/CSV)</b>"))

        # Barre de chemin + bouton Monter
        path_bar = QHBoxLayout()
        self.path_edit = QLineEdit(root, self)
        self.path_edit.setReadOnly(True)
        btn_up = QPushButton("⬆︎ Monter", self)
        btn_up.clicked.connect(self._go_up)
        path_bar.addWidget(QLabel("Dossier :", self))
        path_bar.addWidget(self.path_edit, 1)
        path_bar.addWidget(btn_up)
        layout.addLayout(path_bar)

        # ----- Navigateur fichiers (comme FilePicker) -----
        self.model = QFileSystemModel(self)
        self.model.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot)
        self.model.setNameFilters(["*.xlsx", "*.xls", "*.csv"])
        self.model.setNameFilterDisables(False)
        root = start_dir or DEFAULT_DIR
        root_index = self.model.setRootPath(root)

        self.view = QListView(self)
        self.view.setModel(self.model)
        self.view.setRootIndex(root_index)
        self.view.setViewMode(QListView.IconMode)
        self.view.setResizeMode(QListView.Adjust)
        self.view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.view.doubleClicked.connect(self._on_double_clicked)

        self.btn_add = QPushButton("Ajouter la sélection", self)
        self.btn_add.clicked.connect(self._add_selection)

        layout.addWidget(self.view, 1)
        layout.addWidget(self.btn_add)

        # Info fichier choisi
        self.info = QLabel("Aucun fichier sélectionné", self)
        layout.addWidget(self.info)

        # ----- Visualiseur de métadonnées -----
        self.table = QTableWidget(self)
        # Configuration pour navigation complète
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(QLabel("<b>Visualisation des données</b>"))
        layout.addWidget(self.table, 2)
        layout.addStretch(1)

        # ----- Actions d'assemblage et validation -----
        actions = QHBoxLayout()
        self.btn_assemble = QPushButton("Assembler (données au format txt + métadonnées)", self)
        self.btn_validate = QPushButton("Enregistrer le fichier", self)
        self.btn_validate.setEnabled(False)
        self.btn_assemble.clicked.connect(self._on_assemble_clicked)
        self.btn_validate.clicked.connect(self._on_validate_clicked)
        actions.addWidget(self.btn_assemble)
        actions.addWidget(self.btn_validate)
        layout.addLayout(actions)

        # stockage du DataFrame fusionné
        self.combined_df = None
        self.last_save_path = None
    def _on_assemble_clicked(self):
        """Construit le fichier global (spectres .txt + Excel) en utilisant les fichiers sélectionnés dans l'onglet Fichiers."""
        # Récupérer la fenêtre principale pour accéder au FilePickerWidget
        main = self.window()
        file_picker = getattr(main, 'file_picker', None)
        if file_picker is None:
            QMessageBox.critical(self, "Erreur", "Sélecteur de fichiers (.txt) introuvable.")
            return
        txt_files = file_picker.get_selected_files() if hasattr(file_picker, 'get_selected_files') else []
        if not txt_files:
            QMessageBox.warning(self, "Données manquantes", "Sélectionnez d'abord des fichiers .txt dans l'onglet Fichiers.")
            return
        if not self.metadata_path:
            QMessageBox.warning(self, "Données manquantes", "Sélectionnez un fichier de métadonnées (Excel ou CSV) ci-dessus.")
            return
        try:
            meta_path_for_build = self.metadata_path
            try:
                if self.metadata_path.lower().endswith(".csv"):
                    # Convertir le CSV chargé (self.df) en Excel temporaire pour l'assembleur si nécessaire
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        temp_xlsx = tmp.name
                    # Si self.df n'est pas défini, relire le CSV
                    df_meta = self.df
                    if df_meta is None:
                        df_meta = pd.read_csv(self.metadata_path, sep=None, engine="python")
                    df_meta.to_excel(temp_xlsx, index=False)
                    meta_path_for_build = temp_xlsx
                combined = build_combined_dataframe(txt_files, meta_path_for_build, exclude_brb=True)
            except Exception as e:
                QMessageBox.critical(self, "Échec", f"Impossible d'assembler les données : {e}")
                return
        except Exception as e:
            QMessageBox.critical(self, "Échec", f"Impossible d'assembler les données : {e}")
            return
        self.combined_df = combined
        # Aperçu dans le tableau
        self._update_preview(self.combined_df)
        self.info.setText(f"Assemblage réussi — {combined.shape[0]} lignes, {combined.shape[1]} colonnes")
        self.btn_validate.setEnabled(True)

    def _on_validate_clicked(self):
        """Enregistre le fichier combiné en CSV ou Excel, au choix de l'utilisateur."""
        if self.combined_df is None or self.combined_df.empty:
            QMessageBox.warning(self, "Rien à valider", "Assemblez d'abord les données.")
            return

        # Prépare un chemin suggéré
        base_dir = os.path.dirname(self.metadata_path) if self.metadata_path else os.path.expanduser('~')
        default_name = "combined_raman.csv"
        suggested = self.last_save_path or os.path.join(base_dir, default_name)

        # Boîte de dialogue : choix du format
        filters = "CSV (*.csv);;Excel (*.xlsx)"
        file_path, selected_filter = QFileDialog.getSaveFileName(self, "Enregistrer le fichier combiné", suggested, filters)
        if not file_path:
            return

        # Normalise l'extension en fonction du filtre choisi
        if selected_filter.startswith("CSV") and not file_path.lower().endswith(".csv"):
            file_path += ".csv"
        elif selected_filter.startswith("Excel") and not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        # Écrit selon l'extension
        try:
            if file_path.lower().endswith('.csv'):
                self.combined_df.to_csv(file_path, index=False)
            else:
                self.combined_df.to_excel(file_path, index=False)
            self.last_save_path = file_path
            QMessageBox.information(self, "Fichier créé", f"Fichier global enregistré :\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Échec écriture", f"Impossible d'écrire le fichier : {e}")

    def _add_selection(self):
        indexes = self.view.selectedIndexes()
        if not indexes:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un fichier Excel ou CSV.")
            return

        file_path = self.model.filePath(indexes[0])
        if not (file_path.endswith(".xlsx") or file_path.endswith(".xls") or file_path.endswith(".csv")):
            QMessageBox.warning(self, "Erreur", "Le fichier doit être un Excel (.xlsx ou .xls) ou un CSV (.csv).")
            return

        self.metadata_path = file_path
        self.path_edit.setText(file_path)
        self.info.setText("Fichier chargé : " + os.path.basename(file_path))

        # Charger et afficher un aperçu
        try:
            if file_path.endswith(".csv"):
                # Détection automatique du séparateur
                self.df = pd.read_csv(file_path, sep=None, engine="python")
            else:
                self.df = pd.read_excel(file_path, skiprows=1)
            self._update_preview(self.df)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier : {e}")
            self.df = None

    def _update_preview(self, df: pd.DataFrame, nrows: int = 20):  # nrows conservé pour compatibilité, ignoré
        # Afficher TOUT le DataFrame avec défilement et tri
        self.table.setSortingEnabled(False)
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str).tolist())
        for i in range(len(df)):
            for j, col in enumerate(df.columns):
                val = str(df.iloc[i, j])
                self.table.setItem(i, j, QTableWidgetItem(val))
        self.table.setSortingEnabled(True)

    def _set_root_path(self, path: str):
        if not path:
            return
        idx = self.model.index(path)
        if idx.isValid():
            self.view.setRootIndex(idx)
            self.path_edit.setText(path)

    def _on_double_clicked(self, index: QModelIndex):
        path = self.model.filePath(index)
        if self.model.isDir(index):
            self._set_root_path(path)
            return
        # fichier sélectionné
        if path.endswith(".xlsx") or path.endswith(".xls") or path.endswith(".csv"):
            self.metadata_path = path
            self.path_edit.setText(path)
            self.info.setText("Fichier chargé : " + os.path.basename(path))
            try:
                if path.endswith(".csv"):
                    self.df = pd.read_csv(path, sep=None, engine="python")
                else:
                    self.df = pd.read_excel(path, skiprows=1)
                self._update_preview(self.df)
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier : {e}")
                self.df = None
        else:
            QMessageBox.information(self, "Info", "Sélectionnez un fichier Excel (.xlsx ou .xls) ou un CSV (.csv).")

    def _go_up(self):
        current_idx = self.view.rootIndex()
        parent_idx = current_idx.parent()
        if parent_idx.isValid():
            self.view.setRootIndex(parent_idx)
            self.path_edit.setText(self.model.filePath(parent_idx))

    def load_metadata(self) -> pd.DataFrame | None:
        return self.df