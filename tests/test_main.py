"""src.main birim testleri."""

from __future__ import annotations

from pathlib import Path

from src import main


class TestMainBootstrap:
    """Giris noktasi bootstrap davranisi testleri."""

    def test_bootstrap_local_package_resolution_adds_project_root_for_script(self):
        """Dogrudan script calismasinda proje koku import yoluna eklenmeli."""
        import_path = ["/tmp/project/src", "/usr/lib/python3.12"]
        entry_file = "/tmp/project/src/main.py"

        main.bootstrap_local_package_resolution(None, entry_file, import_path)

        assert import_path[0] == "/tmp/project"

    def test_bootstrap_local_package_resolution_skips_when_running_as_package(self):
        """Paket icinden calismada import yolu degistirilmemeli."""
        original_path = ["/tmp/project", "/usr/lib/python3.12"]
        import_path = list(original_path)
        entry_file = "/tmp/project/src/main.py"

        main.bootstrap_local_package_resolution("src", entry_file, import_path)

        assert import_path == original_path

    def test_prepend_import_path_is_idempotent(self):
        """Ayni proje koku import yoluna iki kez eklenmemeli."""
        import_path = ["/tmp/project", "/usr/lib/python3.12"]

        main._prepend_import_path(import_path, Path("/tmp/project"))

        assert import_path == ["/tmp/project", "/usr/lib/python3.12"]


class TestMainStylesheetPath:
    """QSS yol secimi testleri."""

    def test_get_stylesheet_path_uses_source_tree_without_meipass(self):
        """Normal calismada stil dosyasi kaynak agacindan alinmali."""
        stylesheet_path = main._get_stylesheet_path()

        assert stylesheet_path == Path(main.__file__).parent / "gui" / "style.qss"
