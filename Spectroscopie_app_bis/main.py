import sys
import os
import traceback

from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt

from windows.file_picker import FilePickerWindow


# --- utilitaire ressource (même nomenclature) --------------------------------
def resource_path(relative):
    """
    Renvoie un chemin absolu valide que l’app soit lancée depuis le source ou un binaire (PyInstaller).
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative)


def main():
    print("\n===== Démarrage de l'application Spectroscopie Raman =====")
    app = QApplication(sys.argv)

    # --- Splash (optionnel, safe) --------------------------------------------
    try:
        splash_pix = QPixmap(resource_path(os.path.join("images", "Logo.png")))
        if not splash_pix.isNull():
            splash = QSplashScreen(splash_pix)
            splash.setWindowFlag(Qt.FramelessWindowHint)
            splash.showMessage("Chargement…", Qt.AlignBottom | Qt.AlignCenter, Qt.white)
            splash.show()
            app.processEvents()
        else:
            splash = None
    except Exception:
        splash = None

    # --- Style QSS ------------------------------------------------------------
    try:
        qss_path = resource_path(os.path.join("styles", "styles.qss"))
        if os.path.exists(qss_path):
            with open(qss_path, "r", encoding="utf-8") as f:
                app.setStyleSheet(f.read())
    except Exception as e:
        print("QSS non appliqué :", e)

    # --- Icône ----------------------------------------------------------------
    try:
        icon_path = resource_path(os.path.join("images", "icon.icns"))
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
    except Exception as e:
        print("Icône non appliquée :", e)

    # --- Fenêtre principale : FilePickerWindow -------------------------------
    try:
        window = FilePickerWindow(
            default_dir=os.path.expanduser("~/Documents")  # <— change si besoin
        )
        window.show()
        if splash:
            splash.finish(window)
        sys.exit(app.exec())
    except Exception as e:
        print("Erreur au lancement :", e)
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()