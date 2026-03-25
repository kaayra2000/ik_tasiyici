"""TutanakService birim testleri."""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.core.excel_reader import PersonelOkumaRaporu, SatirReddi
from src.core.excel_writer import TutanakOlusturmaRaporu
from src.gui.tutanak_service import TutanakService


class TestTutanakService:
    """TutanakService sınıfı testleri."""

    @pytest.fixture()
    def service(self):
        return TutanakService()

    @patch("src.gui.tutanak_service.oku_personel_listesi_raporlu")
    def test_personel_oku_basarili(self, mock_oku, service):
        """Personel listesi başarıyla okunmalı."""
        mock_oku.return_value = PersonelOkumaRaporu(
            personeller=[MagicMock(), MagicMock()],
            reddedilen_satirlar=[],
        )
        result = service.personel_oku("/test/girdi.xlsx")
        mock_oku.assert_called_once_with("/test/girdi.xlsx")
        assert len(result) == 2

    @patch("src.gui.tutanak_service.oku_personel_listesi_raporlu")
    def test_personel_oku_dosya_bulunamadi(self, mock_oku, service):
        """Dosya yoksa FileNotFoundError fırlatmalı."""
        mock_oku.side_effect = FileNotFoundError("Dosya yok")
        with pytest.raises(FileNotFoundError):
            service.personel_oku("/olmayan/dosya.xlsx")

    @patch("src.gui.tutanak_service.oku_personel_listesi_raporlu")
    def test_son_personel_okuma_uyarilari_doner(self, mock_oku, service):
        """Atlanan satırlar GUI için log mesajına çevrilmeli."""
        mock_oku.return_value = PersonelOkumaRaporu(
            personeller=[],
            reddedilen_satirlar=[
                SatirReddi(
                    excel_satir_no=2,
                    sebep="Geçersiz TCKN: 35519215090",
                    tckn="35519215090",
                    ad_soyad="Ayşe KOŞUK",
                    birim="C123",
                )
            ],
        )

        service.personel_oku("/test/girdi.xlsx")

        warnings = service.son_personel_okuma_uyarilari()
        assert len(warnings) == 1
        assert "Geçersiz TCKN: 35519215090" in warnings[0]

    @patch("src.gui.tutanak_service.olustur_dk_klasoru_raporlu")
    def test_tutanak_olustur_basarili(self, mock_olustur, service):
        """Tutanak başarıyla oluşturulmalı."""
        expected_path = Path("/cikti")
        mock_olustur.return_value = TutanakOlusturmaRaporu(
            output_path=expected_path,
            added_file_count=2,
            skipped_existing_file_count=0,
            warning_messages=[],
        )
        personeller = [MagicMock(), MagicMock()]

        result = service.tutanak_olustur(
            personeller=personeller,
            template_path="/taslak/sablon.xlsx",
            output_dir="/cikti",
        )

        mock_olustur.assert_called_once_with(
            personeller=personeller,
            cikti_klasoru=Path("/cikti"),
            template_path="/taslak/sablon.xlsx",
            version="v1",
        )
        assert result == expected_path

    @patch("src.gui.tutanak_service.olustur_dk_klasoru_raporlu")
    def test_tutanak_olustur_hata(self, mock_olustur, service):
        """Core katmanıdaki hatalar yayılmalı."""
        mock_olustur.side_effect = ValueError("Şablon hatası")
        with pytest.raises(ValueError, match="Şablon hatası"):
            service.tutanak_olustur(
                personeller=[MagicMock()],
                template_path="/taslak.xlsx",
                output_dir="/cikti",
            )

    @patch("src.gui.tutanak_service.olustur_dk_klasoru_raporlu")
    def test_son_tutanak_olusturma_uyarilari_doner(self, mock_olustur, service):
        """Var olan kayıt nedeniyle atlanan sayfalar GUI'ye iletilmeli."""
        mock_olustur.return_value = TutanakOlusturmaRaporu(
            output_path=Path("/cikti"),
            added_file_count=1,
            skipped_existing_file_count=2,
            warning_messages=[
                "Kayıt atlandı: hedef dosyada zaten mevcut. "
                "SAYFA='Fatma KARACA - 10000000146', "
                "TCKN='10000000146', AD SOYAD='Fatma KARACA', BİRİMİ='Marmara Enstitüsü'"
            ],
        )

        service.tutanak_olustur(
            personeller=[MagicMock()],
            template_path="/taslak.xlsx",
            output_dir="/cikti",
        )

        warnings = service.son_tutanak_olusturma_uyarilari()
        assert len(warnings) == 1
        assert "hedef dosyada zaten mevcut" in warnings[0]
