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
        # métadonnées d'un tableau chargé dont le fichier était introuvable, en
        # attente d'être appliquées par NOM quand l'utilisateur rechargera le .txt
        self._pending_meta: dict[str, tuple] = {}  # nom de fichier -> (volume, série)

    def add(self, path: str, data: tuple) -> None:
        is_new = path not in self.spectra
        self.spectra[path] = data
        if is_new:
            self.volumes.setdefault(path, "")
            self.series.setdefault(path, "")
            # si un tableau chargé attendait ce fichier (par son nom), on applique
            meta = self._pending_meta.pop(self.name(path), None)
            if meta is not None:
                self.volumes[path] = meta[0]
                self.series[path] = meta[1]
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
        self._pending_meta.clear()
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
        if not isinstance(data, dict) or "rows" not in data:
            raise ValueError(
                "Fichier JSON non reconnu : ce n'est pas un tableau Ramanalyze "
                "(enregistré via « Enregistrer le tableau… »)."
            )
        if isinstance(data.get("titrant"), dict):
            self.titrant.update(data["titrant"])

        # index des spectres déjà chargés, par nom de fichier (pour la correspondance)
        by_name = {}
        for q in self.spectra:
            by_name.setdefault(self.name(q), q)

        missing = []
        for row in data.get("rows", []):
            if not isinstance(row, dict):
                continue
            p = row.get("path", "")
            if not p:
                continue
            name = row.get("name") or os.path.basename(p)
            volume = row.get("volume", "")
            serie = row.get("serie", "")

            if p in self.spectra:
                target = p                                  # même chemin déjà chargé
            elif os.path.exists(p):
                spec = load_spectrum(p)
                if spec is None:
                    self._pending_meta[name] = (volume, serie)
                    missing.append(name)
                    continue
                self.spectra[p] = spec
                target = p
            elif name in by_name:
                target = by_name[name]                      # déjà chargé sous un autre chemin → match par NOM
            else:
                # fichier absent : on parque les métadonnées par nom ; elles seront
                # appliquées automatiquement quand l'utilisateur rechargera ce .txt
                self._pending_meta[name] = (volume, serie)
                missing.append(name)
                continue

            self.volumes[target] = volume
            self.series[target] = serie
        self.changed.emit()
        return missing
