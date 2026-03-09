"""EducationImportWindow birim testleri."""

import os
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from src.core.education_importer import EducationImportResult
from src.gui.education_import_window import EducationImportWindow
from src.gui.settings_manager import SettingsManager


@pytest.fixture(scope="session")
def qapp():
    """Test oturumu için tek bir QApplication örneği sağlar."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture()
def settings(tmp_path: Path):
    """EducationImportWindow için yalın ayar yöneticisi taklidi döndürür."""
    manager = MagicMock()
    manager.get_existing_file.side_effect = lambda key: {
        SettingsManager.KEY_EDUCATION_TARGET_PATH: str(tmp_path / "hedef.xlsx"),
        SettingsManager.KEY_EDUCATION_SOURCE_PATH: "",
    }.get(key, "")
    manager.get_parent_dir.side_effect = lambda key: {
        SettingsManager.KEY_EDUCATION_SOURCE_PATH: str(tmp_path),
    }.get(key, "")
    return manager


@pytest.fixture()
def service(tmp_path: Path):
    """EducationImportWindow için servis taklidi döndürür."""
    mock_service = MagicMock()
    mock_service.import_education.return_value = EducationImportResult(
        backup_path=tmp_path / "hedef_eski_20260309_1200.xlsx",
        matched_sheet_count=1,
        updated_sheet_count=1,
        appended_record_count=2,
        skipped_record_count=0,
        unmatched_tckns=["10000000078"],
    )
    mock_service.son_import_uyarilari.return_value = [
        "Kaynak satır 4 atlandı: Geçersiz TCKN: 35519215090. "
        "TCKN='35519215090', AD='AYŞE', ÜNİVERSİTE='ÖRNEK', PROGRAM='BİLGİSAYAR'"
    ]
    return mock_service


@pytest.fixture()
def window(qapp, settings, service):
    """Test için pencere örneği üretir."""
    widget = EducationImportWindow(settings=settings, service=service)
    yield widget
    widget.close()


class TestEducationImportWindow:
    """EducationImportWindow davranış testleri."""

    def test_load_settings_restores_file_and_last_directory(
        self,
        window: EducationImportWindow,
        tmp_path: Path,
    ):
        """Var olan dosya ve son klasör seçicilere yüklenmeli."""
        assert window._target_selector.get_path() == str(tmp_path / "hedef.xlsx")
        assert window._source_selector.get_path() == ""
        assert window._source_selector._dialog_path == str(tmp_path)

    @patch("src.gui.education_import_window.QMessageBox.information")
    def test_start_import_calls_service_and_logs_result(
        self,
        mock_information,
        window: EducationImportWindow,
        service,
        tmp_path: Path,
    ):
        """Başarılı işlemde servis çağrılmalı ve kullanıcı bilgilendirilmeli."""
        target_path = tmp_path / "hedef.xlsx"
        source_path = tmp_path / "mezuniyet.xlsx"
        target_path.touch()
        source_path.touch()

        window._target_selector.set_path(str(target_path))
        window._source_selector.set_path(str(source_path))

        window._start_import()

        service.import_education.assert_called_once_with(
            source_path=str(source_path),
            target_path=str(target_path),
        )
        log_lines = window._log_widget._text_edit.toPlainText().splitlines()
        detail_index = next(
            i for i, line in enumerate(log_lines)
            if "Geçersiz TCKN: 35519215090" in line
        )
        summary_index = log_lines.index("Özet:")
        assert detail_index < summary_index
        assert "Eklenen eğitim kaydı: 2" in "\n".join(log_lines)
        assert log_lines[-1] == "Hedefte bulunamayan TCKN'ler: 10000000078"
        mock_information.assert_called_once()

    @patch("src.gui.education_import_window.QMessageBox.critical")
    def test_start_import_logs_warnings_on_failure(
        self,
        mock_critical,
        window: EducationImportWindow,
        service,
        tmp_path: Path,
    ):
        """Hata durumunda da servis uyarıları log'a yazılmalı."""
        target_path = tmp_path / "hedef.xlsx"
        source_path = tmp_path / "mezuniyet.xlsx"
        target_path.touch()
        source_path.touch()

        service.import_education.side_effect = ValueError("Kaynak dosyada işlenecek geçerli mezuniyet kaydı bulunamadı.")

        window._target_selector.set_path(str(target_path))
        window._source_selector.set_path(str(source_path))

        window._start_import()

        log_lines = window._log_widget._text_edit.toPlainText().splitlines()
        detail_index = next(
            i for i, line in enumerate(log_lines)
            if "Geçersiz TCKN: 35519215090" in line
        )
        summary_index = log_lines.index("Özet:")
        assert detail_index < summary_index
        assert "işlenecek geçerli mezuniyet kaydı bulunamadı" in "\n".join(log_lines)
        assert log_lines[-1] == "Hata: Kaynak dosyada işlenecek geçerli mezuniyet kaydı bulunamadı."
        mock_critical.assert_called_once()

    @patch("src.gui.education_import_window.QMessageBox.warning")
    def test_start_import_requires_target_file(
        self,
        mock_warning,
        window: EducationImportWindow,
    ):
        """Hedef dosya seçilmeden işlem başlamamalı."""
        window._target_selector.set_path("")
        window._source_selector.set_path("/tmp/kaynak.xlsx")

        window._start_import()

        mock_warning.assert_called_once()
