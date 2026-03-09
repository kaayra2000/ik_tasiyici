"""Uygulama giris noktasi."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import MutableSequence


def _is_direct_script_execution(package_name: str | None) -> bool:
    """Dosyanin paket disindan dogrudan calistirilip calistirilmadigini doner."""
    return package_name in (None, "")


def _get_project_root(entry_file: str | Path) -> Path:
    """Giris dosyasindan proje kokunu hesaplar."""
    return Path(entry_file).resolve().parent.parent


def _prepend_import_path(
    import_path: MutableSequence[str], project_root: Path
) -> None:
    """Proje kokunu import arama yolunun basina ekler."""
    project_root_str = str(project_root)
    if project_root_str not in import_path:
        import_path.insert(0, project_root_str)


def bootstrap_local_package_resolution(
    package_name: str | None,
    entry_file: str | Path,
    import_path: MutableSequence[str],
) -> None:
    """`python src/main.py` kullaniminda yerel proje paketini one alir."""
    if not _is_direct_script_execution(package_name):
        return

    project_root = _get_project_root(entry_file)
    _prepend_import_path(import_path, project_root)


def _create_application(argv: list[str]):
    """QApplication ornegini olusturur ve temel stilleri uygular."""
    from src.config.constants import APP_NAME, APP_ORGANIZATION_NAME
    from PyQt6.QtWidgets import QApplication

    app = QApplication(argv)
    app.setOrganizationName(APP_ORGANIZATION_NAME)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    return app


def _get_stylesheet_path() -> Path:
    """Calisma ortamina gore QSS dosyasi yolunu doner."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "src" / "gui" / "style.qss"
    return Path(__file__).parent / "gui" / "style.qss"


def _apply_stylesheet(app) -> None:
    """QSS dosyasi varsa uygulamaya uygular."""
    qss_path = _get_stylesheet_path()
    if not qss_path.exists():
        return

    with open(qss_path, "r", encoding="utf-8") as stylesheet_file:
        app.setStyleSheet(stylesheet_file.read())


def _create_main_menu_window():
    """Ana pencereyi olusturur."""
    from src.gui.main_menu_window import MainMenuWindow

    return MainMenuWindow()


bootstrap_local_package_resolution(__package__, __file__, sys.path)


def main() -> None:
    """Uygulamayi baslatir."""
    app = _create_application(sys.argv)
    _apply_stylesheet(app)

    window = _create_main_menu_window()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
