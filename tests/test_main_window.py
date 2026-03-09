"""MainWindow birim testleri."""

import os
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.main_window import MainWindow


@pytest.fixture(scope="session")
def qapp():
    """Test oturumu için tek bir QApplication örneği sağlar."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture()
def settings():
    """MainWindow için yalın ayar yöneticisi taklidi döndürür."""
    manager = MagicMock()
    manager.get_existing_file.return_value = ""
    manager.get.side_effect = lambda key, default="": default
    return manager


@pytest.fixture()
def service():
    """MainWindow için yalın servis taklidi döndürür."""
    return MagicMock()


@pytest.fixture()
def window(qapp, settings, service):
    """Test için ana pencere örneği üretir."""
    main_window = MainWindow(settings=settings, service=service)
    yield main_window
    main_window.close()


class TestMainWindow:
    """MainWindow davranış testleri."""

    @patch("src.gui.main_window.QDesktopServices.openUrl", return_value=True)
    def test_open_generated_output_opens_folder_and_file(
        self, mock_open_url, window, tmp_path
    ):
        """Oluşan çıktı için klasör ve dosya ayrı ayrı açılmalı."""
        result_path = (tmp_path / "DK_Tutanaklari.xlsx").resolve()

        window._open_generated_output(result_path)

        opened_paths = [
            call.args[0].toLocalFile() for call in mock_open_url.call_args_list
        ]
        assert opened_paths == [str(result_path.parent), str(result_path)]

    @patch.object(MainWindow, "_open_generated_output")
    @patch("src.gui.main_window.QMessageBox.information")
    def test_start_processing_opens_generated_output_after_success(
        self,
        mock_information,
        mock_open_generated_output,
        window,
        service,
        tmp_path,
    ):
        """Başarılı üretimden sonra çıktı otomatik açılmalı."""
        template_path = tmp_path / "taslak.xlsx"
        template_path.touch()
        output_path = tmp_path / "cikti.xlsx"
        result_path = tmp_path / "DK_Tutanaklari.xlsx"

        service.personel_oku.return_value = [MagicMock()]
        service.tutanak_olustur.return_value = result_path

        window._input_selector.set_path(str(tmp_path / "girdi.xlsx"))
        window._template_selector.set_path(str(template_path))
        window._output_selector.set_path(str(output_path))

        window._start_processing()

        service.personel_oku.assert_called_once_with(
            str(tmp_path / "girdi.xlsx")
        )
        service.tutanak_olustur.assert_called_once_with(
            personeller=service.personel_oku.return_value,
            template_path=str(template_path),
            output_path=str(output_path),
            version="v1",
        )
        mock_open_generated_output.assert_called_once_with(result_path)
        mock_information.assert_called_once()

    @patch("src.gui.main_window.QMessageBox.information")
    def test_start_processing_logs_row_rejection_reasons_when_no_valid_personnel(
        self,
        mock_information,
        window,
        service,
        tmp_path,
    ):
        """Geçersiz satır nedenleri log'a yazılmalı."""
        template_path = tmp_path / "taslak.xlsx"
        template_path.touch()
        output_path = tmp_path / "cikti.xlsx"

        service.personel_oku.return_value = []
        service.son_personel_okuma_uyarilari.return_value = [
            "Satır 2 atlandı: Geçersiz TCKN: 35519215090. "
            "TCKN='35519215090', AD SOYAD='Ayşe KOŞUK', BİRİMİ='C123'"
        ]

        window._input_selector.set_path(str(tmp_path / "girdi.xlsx"))
        window._template_selector.set_path(str(template_path))
        window._output_selector.set_path(str(output_path))

        window._start_processing()

        log_text = window._log_widget._text_edit.toPlainText()
        assert "Geçersiz TCKN: 35519215090" in log_text
        assert "İşlenecek personel bulunamadı" in log_text
        mock_information.assert_called_once()
