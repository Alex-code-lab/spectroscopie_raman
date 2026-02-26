from __future__ import annotations

import os
import re

from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import QStandardPaths
from PySide6.QtWebEngineWidgets import QWebEngineView

_CONNECTED_PROFILE_IDS: set[int] = set()
_LAST_DOWNLOAD_DIR: str | None = None
_KNOWN_EXTS = {".png", ".pdf", ".svg"}
_LAST_PLOTLY_FILENAME: str | None = None


def sanitize_filename(name: str) -> str:
    if not name:
        return ""
    text = str(name).strip()
    if not text:
        return ""
    text = text.replace(" ", "_")
    text = re.sub(r"[<>:\"/\\\\|?*\\n\\r\\t]", "", text)
    text = re.sub(r"(?i)\\.(png|pdf|svg)", "", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("._")


def _strip_known_ext(name: str) -> str:
    base = str(name).strip()
    # Retire toutes les extensions connues en fin de nom (même répétées)
    while True:
        m = re.search(r"(?:\.(png|pdf|svg))\s*$", base, flags=re.IGNORECASE)
        if not m:
            return base
        base = base[: m.start()].rstrip(" ._\t\r\n")


def _get_page_from_download(download):
    for attr in ("page", "originatingPage"):
        try:
            val = getattr(download, attr, None)
            page = val() if callable(val) else val
        except Exception:
            page = None
        if page is not None:
            return page
    return None


def _get_profile_from_download(download):
    for attr in ("profile",):
        try:
            val = getattr(download, attr, None)
            prof = val() if callable(val) else val
        except Exception:
            prof = None
        if prof is not None:
            return prof
    page = _get_page_from_download(download)
    try:
        return page.profile() if page is not None else None
    except Exception:
        return None


def _desired_ext_from_filter(selected_filter: str | None) -> str | None:
    if not selected_filter:
        return None
    if "PDF" in selected_filter:
        return ".pdf"
    if "SVG" in selected_filter:
        return ".svg"
    if "PNG" in selected_filter:
        return ".png"
    return None


def _apply_extension(path: str, desired_ext: str | None) -> str:
    if not desired_ext:
        return path
    dir_name, file_name = os.path.split(path)
    base = _strip_known_ext(file_name)
    if not base:
        base = "graphique"
    return os.path.join(dir_name, base + desired_ext)


def _default_download_dir() -> str:
    dl = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
    if dl:
        return dl
    return os.path.expanduser("~")


def install_plotly_download_handler(view: QWebEngineView) -> None:
    """Installe un gestionnaire global pour les snapshots Plotly."""
    if view is None:
        return
    try:
        profile = view.page().profile()
    except Exception:
        return
    pid = id(profile)
    if pid in _CONNECTED_PROFILE_IDS:
        return
    _CONNECTED_PROFILE_IDS.add(pid)
    profile.downloadRequested.connect(_on_download_requested)


def set_plotly_filename(view: QWebEngineView | None, filename: str | None) -> None:
    """Enregistre un nom de base préféré pour les exports (sans extension)."""
    global _LAST_PLOTLY_FILENAME
    name = filename or None
    if name:
        name = sanitize_filename(_strip_known_ext(name)) or None
    if view is not None:
        try:
            setattr(view, "_plotly_filename", name)
        except Exception:
            pass
        try:
            profile = view.page().profile()
            profile.setProperty("plotly_filename", name or "")
        except Exception:
            pass
    _LAST_PLOTLY_FILENAME = name


def _on_download_requested(download) -> None:
    global _LAST_DOWNLOAD_DIR
    if download is None:
        return

    # Parent + vue + figure Plotly associée (si disponible)
    parent = None
    view = None
    fig = None
    try:
        page = _get_page_from_download(download)
        view = page.view() if page is not None else None
        parent = view.window() if view is not None else None
        fig = getattr(view, "_plotly_fig", None) if view is not None else None
    except Exception:
        parent = None
        view = None
        fig = None

    try:
        fname = download.downloadFileName()
    except Exception:
        fname = ""
    if not fname:
        fname = "newplot.png"

    base_name = None
    if view is not None:
        base_name = getattr(view, "_plotly_filename", None)
    if not base_name:
        try:
            prof = _get_profile_from_download(download)
            if prof is not None:
                val = prof.property("plotly_filename")
                base_name = val if isinstance(val, str) and val.strip() else None
        except Exception:
            base_name = None
    if not base_name:
        base_name = _LAST_PLOTLY_FILENAME
    if base_name:
        base_name = os.path.basename(str(base_name))
        base_name = _strip_known_ext(base_name)
        base_name = sanitize_filename(base_name)
    if not base_name:
        base_name = sanitize_filename(_strip_known_ext(fname)) or "graphique"

    base_dir = _LAST_DOWNLOAD_DIR or _default_download_dir()
    default_ext = ".pdf" if fig is not None else ".png"
    suggested = os.path.join(base_dir, base_name + default_ext)
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    path, selected_filter = QFileDialog.getSaveFileName(
        parent,
        "Enregistrer l'image",
        suggested,
        "PDF (*.pdf);;PNG (*.png);;SVG (*.svg);;Tous les fichiers (*.*)",
        options=options,
    )
    if not path:
        try:
            download.cancel()
        except Exception:
            pass
        return

    desired_ext = _desired_ext_from_filter(selected_filter)
    path = _apply_extension(path, desired_ext)
    _, ext = os.path.splitext(path)
    fmt = ext.lstrip(".").lower() if ext else "png"
    if fmt not in {"pdf", "png", "svg"}:
        fmt = "png"

    # Si on a accès à la figure, on exporte via Kaleido (meilleure qualité + PDF)
    if fig is not None:
        try:
            if fmt == "png":
                fig.write_image(path, format=fmt, scale=3)
            else:
                fig.write_image(path, format=fmt)
            _LAST_DOWNLOAD_DIR = os.path.dirname(path)
            try:
                download.cancel()
            except Exception:
                pass
            QMessageBox.information(parent, "Snapshot", f"Image enregistrée :\n{path}")
            return
        except Exception as e:
            if fmt != "png":
                try:
                    download.cancel()
                except Exception:
                    pass
                QMessageBox.critical(
                    parent,
                    "Export PDF/SVG impossible",
                    "L'export en PDF/SVG nécessite 'kaleido'.\n"
                    "Installez-le avec : pip install -U kaleido\n\n"
                    f"Détail : {e}",
                )
                return
            # Pour PNG, on tentera un fallback via le téléchargement du navigateur.

    # Fallback PNG via téléchargement intégré (qualité standard)
    if fmt != "png":
        try:
            download.cancel()
        except Exception:
            pass
        QMessageBox.warning(
            parent,
            "Export indisponible",
            "PDF/SVG indisponible ici. Choisissez PNG ou installez kaleido.",
        )
        return

    try:
        download.setDownloadDirectory(os.path.dirname(path))
        download.setDownloadFileName(os.path.basename(path))
        download.accept()
        _LAST_DOWNLOAD_DIR = os.path.dirname(path)
        QMessageBox.information(parent, "Snapshot", f"Image enregistrée :\n{path}")
    except Exception:
        try:
            download.cancel()
        except Exception:
            pass
