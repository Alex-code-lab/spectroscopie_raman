import os
import numpy as np
import pandas as pd
from PySide6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QMessageBox, QLabel
from PySide6.QtWebEngineWidgets import QWebEngineView
from pybaselines import Baseline
import plotly.graph_objects as go
from dataclasses import dataclass, field
from typing import Dict, Optional


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

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Tracer les spectres sélectionnés"))
        self.btn_plot = QPushButton("Tracer avec baseline")
        self.btn_plot.clicked.connect(self.plot_selected_with_baseline)
        layout.addWidget(self.btn_plot)

        self.plot_view = QWebEngineView(self)
        layout.addWidget(self.plot_view, 1)

    def plot_selected_with_baseline(self):
        files = getattr(self.file_picker, "selected_files", [])
        if not files:
            QMessageBox.information(self, "Info", "Aucun fichier sélectionné.")
            return

        # Utiliser le DataFrame combiné (propre) depuis l'onglet Métadonnées
        main = self.window()
        metadata_picker = getattr(main, 'metadata_picker', None)
        combined_df = getattr(metadata_picker, 'combined_df', None) if metadata_picker is not None else None

        if not isinstance(combined_df, pd.DataFrame) or combined_df.empty:
            QMessageBox.warning(self, "Données manquantes", "Assemble d'abord les données dans l'onglet Métadonnées (bouton ‘Assembler’).")
            return

        required_cols = {"Raman Shift", "Intensity_corrected"}
        if not required_cols.issubset(set(combined_df.columns)):
            QMessageBox.critical(self, "Colonnes manquantes", "Le fichier combiné ne contient pas les colonnes ‘Raman Shift’ et ‘Intensity_corrected’. Réassemblez les données.")
            return

        # Ne tracer que les fichiers sélectionnés (filtre sur la colonne 'file' si présente)
        basenames = {os.path.basename(p) for p in files}
        df_plot = combined_df.copy()
        if "file" in df_plot.columns:
            df_plot = df_plot[df_plot["file"].isin(basenames)]
        if df_plot.empty:
            QMessageBox.warning(self, "Avertissement", "Aucune ligne du fichier combiné ne correspond à la sélection actuelle.")
            return

        label_col = "Sample description" if "Sample description" in df_plot.columns else ("file" if "file" in df_plot.columns else None)

        traces = []
        if label_col is not None:
            for label, g in df_plot.groupby(label_col):
                g = g.dropna(subset=["Raman Shift", "Intensity_corrected"]).sort_values("Raman Shift")
                if g.empty:
                    continue
                traces.append(
                    go.Scatter(x=g["Raman Shift"].to_numpy(), y=g["Intensity_corrected"].to_numpy(), name=str(label), mode="lines")
                )
        else:
            g = df_plot.dropna(subset=["Raman Shift", "Intensity_corrected"]).sort_values("Raman Shift")
            if not g.empty:
                traces.append(
                    go.Scatter(x=g["Raman Shift"].to_numpy(), y=g["Intensity_corrected"].to_numpy(), name="spectre", mode="lines")
                )

        if not traces:
            QMessageBox.warning(self, "Avertissement", "Aucun spectre valide après filtrage du fichier combiné.")
            return

        fig = go.Figure(traces)
        fig.update_layout(
            title="Spectres Raman (depuis le fichier combiné)",
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité corrigée (a.u.)",
            width=1500,
            height=800,
        )
        self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))