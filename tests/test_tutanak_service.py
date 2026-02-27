"""TutanakService birim testleri."""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.gui.tutanak_service import TutanakService


class TestTutanakService:
    """TutanakService sınıfı testleri."""

    @pytest.fixture()
    def service(self):
        return TutanakService()

    @patch("src.gui.tutanak_service.oku_personel_listesi")
    def test_personel_oku_basarili(self, mock_oku, service):
        """Personel listesi başarıyla okunmalı."""
        mock_oku.return_value = [MagicMock(), MagicMock()]
        result = service.personel_oku("/test/girdi.xlsx")
        mock_oku.assert_called_once_with("/test/girdi.xlsx")
        assert len(result) == 2

    @patch("src.gui.tutanak_service.oku_personel_listesi")
    def test_personel_oku_dosya_bulunamadi(self, mock_oku, service):
        """Dosya yoksa FileNotFoundError fırlatmalı."""
        mock_oku.side_effect = FileNotFoundError("Dosya yok")
        with pytest.raises(FileNotFoundError):
            service.personel_oku("/olmayan/dosya.xlsx")

    @patch("src.gui.tutanak_service.olustur_dk_dosyasi")
    def test_tutanak_olustur_basarili(self, mock_olustur, service):
        """Tutanak başarıyla oluşturulmalı."""
        expected_path = Path("/cikti/DK_Tutanaklari.xlsx")
        mock_olustur.return_value = expected_path
        personeller = [MagicMock(), MagicMock()]

        result = service.tutanak_olustur(
            personeller=personeller,
            template_path="/taslak/sablon.xlsx",
            output_path="/cikti/DK_Tutanaklari.xlsx",
        )

        mock_olustur.assert_called_once_with(
            personeller=personeller,
            cikti_dizini=Path("/cikti"),
            dosya_adi="DK_Tutanaklari.xlsx",
            template_path="/taslak/sablon.xlsx",
            version="v1",
        )
        assert result == expected_path

    @patch("src.gui.tutanak_service.olustur_dk_dosyasi")
    def test_tutanak_olustur_hata(self, mock_olustur, service):
        """Core katmanıdaki hatalar yayılmalı."""
        mock_olustur.side_effect = ValueError("Şablon hatası")
        with pytest.raises(ValueError, match="Şablon hatası"):
            service.tutanak_olustur(
                personeller=[MagicMock()],
                template_path="/taslak.xlsx",
                output_path="/cikti/dosya.xlsx",
            )
