"""Magasin partagé des spectres chargés.

Les deux onglets (Visualiseur, Titration) lisent et écrivent ici, et se
synchronisent via le signal `changed`.
"""

import os
from PySide6.QtCore import QObject, Signal


class SpectraStore(QObject):
    changed = Signal()  # émis quand l'ensemble des spectres change

    def __init__(self):
        super().__init__()
        # path -> (x, y) en numpy
        self.spectra: dict[str, tuple] = {}
        # path -> concentration saisie (texte libre, converti en float au besoin)
        self.concentrations: dict[str, str] = {}

    def add(self, path: str, data: tuple) -> None:
        is_new = path not in self.spectra
        self.spectra[path] = data
        if is_new:
            self.concentrations.setdefault(path, "")
        self.changed.emit()

    def remove(self, path: str) -> None:
        self.spectra.pop(path, None)
        self.concentrations.pop(path, None)
        self.changed.emit()

    def clear(self) -> None:
        self.spectra.clear()
        self.concentrations.clear()
        self.changed.emit()

    def paths(self) -> list[str]:
        return list(self.spectra.keys())

    @staticmethod
    def name(path: str) -> str:
        return os.path.basename(path)
