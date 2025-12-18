import os
import pandas as pd
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QMessageBox,
    QLabel,
    QDoubleSpinBox,
    QFileDialog,
    QCheckBox,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
import plotly.graph_objects as go
from data_processing import load_spectrum_file, build_combined_dataframe_from_ui

class SpectraTab(QWidget):
    def __init__(self, file_picker, parent=None):
        super().__init__(parent)
        self.file_picker = file_picker

        # Dernière figure Plotly affichée (pour export PNG)
        self._last_fig = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Tracer les spectres sélectionnés"))

        # Checkbox pour afficher le contrôle BRB
        self.chk_show_brb = QCheckBox("Afficher le contrôle BRB", self)
        self.chk_show_brb.setChecked(False)
        layout.addWidget(self.chk_show_brb)

        # --- Contrôle manuel des axes (optionnel) ---
        # Convention : valeur = -1 => autoscale (affiché comme "auto")
        axes = QHBoxLayout()
        axes.addWidget(QLabel("X min (cm⁻¹):"))
        self.spin_xmin = QDoubleSpinBox(self)
        self.spin_xmin.setRange(-1.0, 5000.0)
        self.spin_xmin.setDecimals(1)
        self.spin_xmin.setSingleStep(10.0)
        self.spin_xmin.setSpecialValueText("auto")
        self.spin_xmin.setValue(-1.0)
        axes.addWidget(self.spin_xmin)

        axes.addWidget(QLabel("X max (cm⁻¹):"))
        self.spin_xmax = QDoubleSpinBox(self)
        self.spin_xmax.setRange(-1.0, 5000.0)
        self.spin_xmax.setDecimals(1)
        self.spin_xmax.setSingleStep(10.0)
        self.spin_xmax.setSpecialValueText("auto")
        self.spin_xmax.setValue(-1.0)
        axes.addWidget(self.spin_xmax)

        axes.addSpacing(20)

        axes.addWidget(QLabel("Y min:"))
        self.spin_ymin = QDoubleSpinBox(self)
        self.spin_ymin.setRange(-1.0, 1e12)
        self.spin_ymin.setDecimals(2)
        self.spin_ymin.setSingleStep(100.0)
        self.spin_ymin.setSpecialValueText("auto")
        self.spin_ymin.setValue(-1.0)
        axes.addWidget(self.spin_ymin)

        axes.addWidget(QLabel("Y max:"))
        self.spin_ymax = QDoubleSpinBox(self)
        self.spin_ymax.setRange(-1.0, 1e12)
        self.spin_ymax.setDecimals(2)
        self.spin_ymax.setSingleStep(100.0)
        self.spin_ymax.setSpecialValueText("auto")
        self.spin_ymax.setValue(-1.0)
        axes.addWidget(self.spin_ymax)

        axes.addStretch(1)

        self.btn_reset_axes = QPushButton("Reset axes", self)
        self.btn_reset_axes.clicked.connect(self._reset_axes)
        axes.addWidget(self.btn_reset_axes)

        layout.addLayout(axes)

        self.btn_plot = QPushButton("Tracer avec baseline")
        self.btn_plot.clicked.connect(self.plot_selected_with_baseline)
        layout.addWidget(self.btn_plot)
        self.btn_export_png = QPushButton("Exporter le graphique (PNG)")
        self.btn_export_png.clicked.connect(self._export_png)
        layout.addWidget(self.btn_export_png)



        self.plot_view = QWebEngineView(self)
        layout.addWidget(self.plot_view, 1)

    def _reset_axes(self):
        """Remet les axes en autoscale."""
        if hasattr(self, "spin_xmin"):
            self.spin_xmin.setValue(-1.0)
        if hasattr(self, "spin_xmax"):
            self.spin_xmax.setValue(-1.0)
        if hasattr(self, "spin_ymin"):
            self.spin_ymin.setValue(-1.0)
        if hasattr(self, "spin_ymax"):
            self.spin_ymax.setValue(-1.0)

    def plot_selected_with_baseline(self):
        """
        Trace les spectres sélectionnés.

        Priorité : utilise le fichier combiné construit depuis l'onglet Métadonnées.
        Fallback : lit directement les fichiers .txt sélectionnés.
        """


        try:
            # 1) On tente de construire un fichier combiné à partir des métadonnées créées dans l'onglet Métadonnées
            main = self.window()
            metadata_creator = getattr(main, 'metadata_creator', None) if main is not None else None
            combined_df = None
            if metadata_creator is not None and hasattr(metadata_creator, 'build_merged_metadata'):
                # Récupérer les fichiers .txt à partir du FilePicker associé à cet onglet
                if hasattr(self.file_picker, 'get_selected_files'):
                    txt_files = self.file_picker.get_selected_files()
                else:
                    txt_files = getattr(self.file_picker, 'selected_files', [])
                if txt_files:
                    try:
                        exclude_brb = not self.chk_show_brb.isChecked()
                        combined_df = build_combined_dataframe_from_ui(
                            txt_files,
                            metadata_creator,
                            poly_order=5,
                            exclude_brb=exclude_brb,
                            apply_baseline=True,
                        )
                    except Exception as e_inner:
                        QMessageBox.warning(
                            self,
                            "Avertissement",
                            f"Impossible de construire le fichier combiné à partir des métadonnées :\n{e_inner}"
                        )
        except Exception as e:
            QMessageBox.critical(self, "Erreur interne", f"Impossible de construire le fichier combiné :\n{e}")
            combined_df = None

        # --- CAS 1 : Fichier combiné présent ---
        if isinstance(combined_df, pd.DataFrame) and not combined_df.empty:
            try:
                cols = set(combined_df.columns)
                if not {"Raman Shift", "Intensity_corrected"}.issubset(cols):
                    QMessageBox.critical(
                        self,
                        "Colonnes manquantes",
                        f"Le fichier combiné ne contient pas les colonnes 'Raman Shift' et 'Intensity_corrected'.\n"
                        f"Colonnes présentes : {sorted(cols)}"
                    )
                    return

                df_plot = combined_df.copy()
                df_plot = df_plot.dropna(subset=["Raman Shift", "Intensity_corrected"])

                if df_plot.empty:
                    QMessageBox.warning(
                        self,
                        "Avertissement",
                        "Le fichier combiné ne contient pas de données exploitables "
                        "(valeurs manquantes dans 'Raman Shift' / 'Intensity_corrected')."
                    )
                    return

                # Choix de la légende
                if "Sample description" in df_plot.columns:
                    label_col = "Sample description"
                elif "file" in df_plot.columns:
                    label_col = "file"
                else:
                    label_col = None

                traces = []

                if label_col is not None:
                    for label, g in df_plot.groupby(label_col):
                        g = g.dropna(subset=["Raman Shift", "Intensity_corrected"]).sort_values("Raman Shift")
                        if g.empty:
                            continue
                        traces.append(
                            go.Scatter(
                                x=g["Raman Shift"].to_numpy(),
                                y=g["Intensity_corrected"].to_numpy(),
                                name=str(label),
                                mode="lines",
                            )
                        )
                else:
                    g = df_plot.sort_values("Raman Shift")
                    traces.append(
                        go.Scatter(
                            x=g["Raman Shift"].to_numpy(),
                            y=g["Intensity_corrected"].to_numpy(),
                            name="spectres",
                            mode="lines",
                        )
                    )

                if not traces:
                    QMessageBox.warning(
                        self,
                        "Avertissement",
                        "Aucun spectre à tracer après filtrage."
                    )
                    return

                # Récupérer le nom de la manip pour le titre
                manip_name = None
                main = self.window()
                if main is not None:
                    metadata_creator = getattr(main, "metadata_creator", None)
                    if metadata_creator is not None and hasattr(metadata_creator, "edit_manip"):
                        manip_name = metadata_creator.edit_manip.text().strip()

                title = manip_name if manip_name else "Spectres Raman"

                fig = go.Figure(traces)
                fig.update_layout(
                    title=title,
                    xaxis_title="Raman Shift (cm⁻¹)",
                    yaxis_title="Intensité corrigée (a.u.)",
                    width=1200,
                    height=600,
                )

                # --- Axes manuels si définis (valeur -1 => auto) ---
                xmin = float(self.spin_xmin.value()) if hasattr(self, "spin_xmin") else -1.0
                xmax = float(self.spin_xmax.value()) if hasattr(self, "spin_xmax") else -1.0
                ymin = float(self.spin_ymin.value()) if hasattr(self, "spin_ymin") else -1.0
                ymax = float(self.spin_ymax.value()) if hasattr(self, "spin_ymax") else -1.0

                if xmin >= 0 and xmax >= 0 and xmin < xmax:
                    fig.update_xaxes(range=[xmin, xmax])
                if ymin >= 0 and ymax >= 0 and ymin < ymax:
                    fig.update_yaxes(range=[ymin, ymax])

                self._last_fig = fig  # Sauvegarde pour export PNG
                self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))
                return

            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Erreur tracé (combined_df)",
                    f"Exception lors de l'utilisation du fichier combiné :\n{e}"
                )
                return

        # --- CAS 2 : fallback – lecture directe des .txt sélectionnés ---
        files = getattr(self.file_picker, "selected_files", [])
        if not files:
            QMessageBox.information(
                self,
                "Info",
                "Aucun fichier combiné valide et aucun fichier .txt sélectionné.\n"
                "Allez dans l’onglet Fichiers, ajoutez des .txt, puis réessayez."
            )
            return

        traces = []
        for path in files:
            try:
                temp = load_spectrum_file(path)
                if temp is None or temp.empty:
                    continue

                # Sécurité : s'assurer que les colonnes nécessaires sont là
                if not {"Raman Shift", "Intensity_corrected"}.issubset(temp.columns):
                    continue

                g = temp.dropna(subset=["Raman Shift", "Intensity_corrected"]).sort_values("Raman Shift")
                if g.empty:
                    continue

                traces.append(
                    go.Scatter(
                        x=g["Raman Shift"].to_numpy(),
                        y=g["Intensity_corrected"].to_numpy(),
                        name=os.path.basename(path),
                        mode="lines",
                    )
                )
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Erreur sur un fichier",
                    f"Impossible de traiter {os.path.basename(path)} :\n{e}"
                )

        if not traces:
            QMessageBox.warning(
                self,
                "Avertissement",
                "Aucun spectre valide n'a pu être chargé depuis les fichiers .txt sélectionnés."
            )
            return

        fig = go.Figure(traces)
        fig.update_layout(
            title="Spectres Raman (lecture directe des .txt via data_processing)",
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité corrigée (a.u.)",
            width=1200,
            height=600,
        )

        # --- Axes manuels si définis (valeur -1 => auto) ---
        xmin = float(self.spin_xmin.value()) if hasattr(self, "spin_xmin") else -1.0
        xmax = float(self.spin_xmax.value()) if hasattr(self, "spin_xmax") else -1.0
        ymin = float(self.spin_ymin.value()) if hasattr(self, "spin_ymin") else -1.0
        ymax = float(self.spin_ymax.value()) if hasattr(self, "spin_ymax") else -1.0

        if xmin >= 0 and xmax >= 0 and xmin < xmax:
            fig.update_xaxes(range=[xmin, xmax])
        if ymin >= 0 and ymax >= 0 and ymin < ymax:
            fig.update_yaxes(range=[ymin, ymax])

        self._last_fig = fig
        self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))

    def _export_png(self):
        """Exporte le dernier graphique Plotly affiché vers un fichier PNG."""
        if self._last_fig is None:
            QMessageBox.information(
                self,
                "Aucun graphique",
                "Aucun graphique n'est disponible à l'export.\n"
                "Tracez d'abord les spectres."
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Exporter le graphique (PNG)",
            "spectres_raman.png",
            "Image PNG (*.png)"
        )
        if not path:
            return

        try:
            # Export haute résolution (nécessite kaleido)
            self._last_fig.write_image(path, format="png", scale=2)
            QMessageBox.information(self, "Export réussi", f"Graphique exporté :\n{path}")
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erreur d'export",
                "Impossible d'exporter le graphique en PNG.\n"
                "Vérifiez que 'kaleido' est installé (pip install -U kaleido).\n\n"
                f"Détail : {e}"
            )
