"""Vue Plotly réutilisable basée sur QWebEngineView.

QWebEngineView.setHtml() est limité à ~2 Mo ; avec le Plotly inline le HTML
dépasse cette taille et la page reste blanche. On écrit donc le HTML dans un
fichier temporaire chargé via setUrl(), qui n'a pas de limite.
"""

import os
import tempfile

from PySide6.QtCore import QUrl
from PySide6.QtWebEngineWidgets import QWebEngineView
import plotly.io as pio


class PlotlyView(QWebEngineView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._temp_html_path = None

    def show_figure(self, fig) -> None:
        html = pio.to_html(
            fig, include_plotlyjs="inline", full_html=True,
            config={"responsive": True},
        )
        fd, path = tempfile.mkstemp(suffix=".html", prefix="ramanalyze_simple_")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(html)
        old = self._temp_html_path
        if old and old != path and os.path.exists(old):
            try:
                os.unlink(old)
            except OSError:
                pass
        self._temp_html_path = path
        self.setUrl(QUrl.fromLocalFile(path))
