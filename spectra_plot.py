import os
import numpy as np
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
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from pybaselines import Baseline
import plotly.graph_objects as go
from dataclasses import dataclass, field
from typing import Dict, Optional
from data_processing import load_spectrum_file, build_combined_dataframe_from_df

@dataclass
class Spectrum:
    x: np.ndarray
    y: np.ndarray
    baseline: np.ndarray
    ycorr: np.ndarray
    meta: dict = field(default_factory=dict)


class DataStore:
    """Cache en mémoire des spectres chargés (clé = chemin fichier)."""
    def __init__(self):
        self._data: Dict[str, Spectrum] = {}
        self._mtimes: Dict[str, float] = {}

    def get(self, path: str) -> Optional[Spectrum]:
        return self._data.get(path)

    def set(self, path: str, spec: Spectrum, mtime: float):
        self._data[path] = spec
        self._mtimes[path] = mtime

    def ensure_fresh(self, path: str) -> bool:
        """Retourne True si présent et mtime inchangé, sinon False (à recharger)."""
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return False
        return path in self._data and self._mtimes.get(path) == mtime

    def invalidate(self, path: str):
        self._data.pop(path, None)
        self._mtimes.pop(path, None)


DATASTORE = DataStore()


def load_spectrum(path: str) -> Optional[Spectrum]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        header_idx = next((i for i, line in enumerate(lines) if line.strip().startswith("Pixel;")), 0)

        df = pd.read_csv(
            path,
            skiprows=header_idx,
            sep=";",
            decimal=",",
            encoding="utf-8",
            skipinitialspace=True,
            na_values=["", " ", "   ", "\t"],
            keep_default_na=True,
        )
        if df.columns[-1].startswith("Unnamed"):
            df = df.iloc[:, :-1]
        df.columns = [c.strip().replace("\xa0", " ") for c in df.columns]

        if not {"Raman Shift", "Dark Subtracted #1"}.issubset(df.columns):
            return None

        temp = df[["Raman Shift", "Dark Subtracted #1"]].copy()
        temp["Raman Shift"] = pd.to_numeric(temp["Raman Shift"], errors="coerce")
        temp["Dark Subtracted #1"] = pd.to_numeric(temp["Dark Subtracted #1"], errors="coerce")
        temp = temp.dropna(subset=["Raman Shift", "Dark Subtracted #1"])
        if temp.empty:
            return None

        x = temp["Raman Shift"].to_numpy()
        y = temp["Dark Subtracted #1"].to_numpy()
        order = np.argsort(x)
        x = x[order]
        y = y[order]

        try:
            baseline, _ = Baseline(x).modpoly(y, poly_order=5)
            ycorr = y - baseline
        except Exception:
            baseline = np.zeros_like(y)
            ycorr = y

        return Spectrum(x=x, y=y, baseline=baseline, ycorr=ycorr, meta={"path": path})
    except Exception:
        return None


class SpectraTab(QWidget):
    def __init__(self, file_picker, parent=None):
        super().__init__(parent)
        self.file_picker = file_picker

        # Dernière figure Plotly affichée (pour export PNG)
        self._last_fig = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Tracer les spectres sélectionnés"))

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
        Trace les spectres à partir du fichier combiné s'il est disponible.
        Version très simplifiée + DEBUG : 
        - affiche un message HTML dès l'entrée dans la fonction,
        - essaye d'abord d'utiliser combined_df,
        - sinon lit directement les .txt sélectionnés,
        - et montre des QMessageBox si quelque chose cloche.
        """
        # DEBUG : vérifier que la fonction est bien appelée et que la vue HTML réagit
        self.plot_view.setHtml("<h3>DEBUG : entrée dans plot_selected_with_baseline()</h3>")

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
                        merged_meta = metadata_creator.build_merged_metadata()
                        combined_df = build_combined_dataframe_from_df(
                            txt_files,
                            merged_meta,
                            poly_order=5,
                            exclude_brb=True,
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
                        "DEBUG : combined_df est présent mais toutes les lignes ont des NaN "
                        "dans Raman Shift / Intensity_corrected."
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
                        "DEBUG : combined_df non vide, mais aucune trace construite (traces = [])."
                    )
                    return

                fig = go.Figure(traces)
                fig.update_layout(
                    title="Spectres Raman (depuis le fichier combiné)",
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
                "DEBUG : Aucun fichier combiné valide et aucun fichier .txt sélectionné.\n"
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
                "DEBUG : Lecture directe des fichiers .txt, mais aucune trace valide construite."
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