"""Magasin partagé des spectres chargés.

Les onglets (Visualiseur, Titration, Suivi de pic) lisent et écrivent ici, et se
synchronisent via deux signaux :
- `changed`      : l'ensemble des spectres a changé (ajout / retrait / chargement) ;
- `meta_changed` : les métadonnées (volume, série, réglages titrant) ont changé,
                   sans toucher à l'ensemble des spectres.

Le tableau (fichiers + volumes + séries + réglages titrant) peut être enregistré
et rechargé pour ne pas tout ressaisir.
"""

import os
import json
from PySide6.QtCore import QObject, Signal


class SpectraStore(QObject):
    changed = Signal()       # ensemble des spectres modifié
    meta_changed = Signal()  # volumes / séries / réglages titrant modifiés

    def __init__(self):
        super().__init__()
        # path -> (x, y) en numpy
        self.spectra: dict[str, tuple] = {}
        # path -> volume de titrant saisi (texte)
        self.volumes: dict[str, str] = {}
        # path -> étiquette de série (texte)
        self.series: dict[str, str] = {}
        # réglages globaux du titrant (concentration + unités)
        self.titrant: dict = {"conc": 1.0, "conc_unit": "µM", "vol_unit": "µL"}

    def add(self, path: str, data: tuple) -> None:
        is_new = path not in self.spectra
        self.spectra[path] = data
        if is_new:
            self.volumes.setdefault(path, "")
            self.series.setdefault(path, "")
        self.changed.emit()

    def remove(self, path: str) -> None:
        self.spectra.pop(path, None)
        self.volumes.pop(path, None)
        self.series.pop(path, None)
        self.changed.emit()

    def clear(self) -> None:
        self.spectra.clear()
        self.volumes.clear()
        self.series.clear()
        self.changed.emit()

    def paths(self) -> list[str]:
        return list(self.spectra.keys())

    @staticmethod
    def name(path: str) -> str:
        return os.path.basename(path)

    # ------------------------------------------------------------------
    # Persistance du tableau (volumes / séries / titrant + chemins fichiers)
    # ------------------------------------------------------------------
    def save_session(self, path: str) -> None:
        data = {
            "titrant": self.titrant,
            "rows": [
                {
                    "path": p,
                    "name": self.name(p),
                    "volume": self.volumes.get(p, ""),
                    "serie": self.series.get(p, ""),
                }
                for p in self.paths()
            ],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_session(self, path: str) -> list[str]:
        """Recharge un tableau enregistré. Réimporte les spectres si les fichiers
        existent encore. Renvoie la liste des fichiers introuvables."""
        from spectrum_loader import load_spectrum

        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data.get("titrant"), dict):
            self.titrant.update(data["titrant"])

        missing = []
        for row in data.get("rows", []):
            p = row.get("path", "")
            if not p:
                continue
            if p not in self.spectra:
                spec = load_spectrum(p) if os.path.exists(p) else None
                if spec is None:
                    missing.append(row.get("name") or os.path.basename(p))
                    continue
                self.spectra[p] = spec
            self.volumes[p] = row.get("volume", "")
            self.series[p] = row.get("serie", "")
        self.changed.emit()
        return missing
