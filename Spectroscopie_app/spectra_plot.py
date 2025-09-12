import os
import numpy as np
import pandas as pd
from PySide6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QMessageBox, QLabel
from PySide6.QtWebEngineWidgets import QWebEngineView
from pybaselines import Baseline
import plotly.graph_objects as go


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

        traces = []
        for path in files:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                # Trouver la ligne d'en-tête qui commence par 'Pixel;'
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
                # retirer une éventuelle dernière colonne vide
                if df.columns[-1].startswith("Unnamed"):
                    df = df.iloc[:, :-1]
                # normaliser les noms de colonnes (espaces insécables, etc.)
                df.columns = [c.strip().replace("\xa0", " ") for c in df.columns]

                if not {"Raman Shift", "Dark Subtracted #1"}.issubset(df.columns):
                    continue

                temp = df[["Raman Shift", "Dark Subtracted #1"]].copy()
                temp["Raman Shift"] = pd.to_numeric(temp["Raman Shift"], errors="coerce")
                temp["Dark Subtracted #1"] = pd.to_numeric(temp["Dark Subtracted #1"], errors="coerce")
                temp = temp.dropna(subset=["Raman Shift", "Dark Subtracted #1"])
                if temp.empty:
                    continue

                x = temp["Raman Shift"].to_numpy()
                y = temp["Dark Subtracted #1"].to_numpy()

                # s'assurer que x est croissant
                order = np.argsort(x)
                x = x[order]
                y = y[order]

                try:
                    baseline, _ = Baseline(x).modpoly(y, poly_order=5)
                    ycorr = y - baseline
                except Exception:
                    baseline = np.zeros_like(y)
                    ycorr = y

                name = os.path.basename(path)
                traces.append(
                    go.Scatter(x=x,
                               y=ycorr,
                               name=f"{name}",
                               mode="lines"))
                # traces.append(
                #     go.Scatter(
                #         x=x,
                #         y=baseline,
                #         name=f"{name} – baseline",
                #         mode="lines",
                #         line=dict(dash="dash"),
                #     )
                # )
                # traces.append(go.Scatter(x=x, y=ycorr, name=f"{name} – corrigé", mode="lines"))
            except Exception:
                # ignorer silencieusement les fichiers problématiques
                continue

        if not traces:
            QMessageBox.warning(self, "Erreur", "Aucun spectre valide n'a été trouvé dans la sélection.")
            return

        fig = go.Figure(traces)
        fig.update_layout(
            title="Spectres Raman",
            xaxis_title="Raman Shift (cm⁻¹)",
            yaxis_title="Intensité (a.u.)",
            width=1500,
            height=800,
        )
        self.plot_view.setHtml(fig.to_html(include_plotlyjs="cdn"))