"""
excel_reader modülü için testler.

Testler, gerçek bir xlsx dosyasına ihtiyaç duymadan openpyxl ile
hafızada test dosyaları oluşturur.
"""

from __future__ import annotations

import io
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from src.core.excel_reader import Personel, oku_personel_listesi


# ---------------------------------------------------------------------------
# Yardımcı: hafızada test xlsx oluşturucu
# ---------------------------------------------------------------------------


def _xlsx_yaz(satirlar: list[dict], dosya_yolu: Path) -> None:
    """Verilen satırları bir xlsx dosyasına yazar."""
    df = pd.DataFrame(satirlar)
    df.to_excel(dosya_yolu, index=False)


# ---------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------


@pytest.fixture()
def gecerli_xlsx(tmp_path: Path) -> Path:
    """Üç geçerli personel içeren test xlsx dosyası oluşturur."""
    dosya = tmp_path / "test.xlsx"
    _xlsx_yaz(
        [
            {
                "TCKN": "10000000146",
                "AD SOYAD": "Fatma KARACA",
                "BİRİMİ": "Marmara Enstitüsü",
            },
            {
                "TCKN": "10000000078",
                "AD SOYAD": "Ali YILMAZ",
                "BİRİMİ": "Gebze Enstitüsü",
            },
            {
                "TCKN": "10000050028",
                "AD SOYAD": "Ayşe DEMİR",
                "BİRİMİ": "Kocaeli Enstitüsü",
            },
        ],
        dosya,
    )
    return dosya


# ---------------------------------------------------------------------------
# Testler: oku_personel_listesi
# ---------------------------------------------------------------------------


class TestOkuPersonelListesi:
    """oku_personel_listesi fonksiyonu için testler."""

    def test_gecerli_kayitlari_okur(self, gecerli_xlsx: Path):
        """Geçerli üç kayıt doğru okunmalı."""
        personeller = oku_personel_listesi(gecerli_xlsx)
        assert len(personeller) == 3

    def test_ilk_personel_bilgileri(self, gecerli_xlsx: Path):
        """İlk personelin alanları doğru dolu olmalı."""
        personeller = oku_personel_listesi(gecerli_xlsx)
        p = personeller[0]
        assert p.tckn == "10000000146"
        assert p.ad_soyad == "Fatma KARACA"
        assert p.birim == "Marmara Enstitüsü"

    def test_personel_frozen_dataclass(self, gecerli_xlsx: Path):
        """Personel nesnesi değiştirilemez olmalı (frozen dataclass)."""
        personel = oku_personel_listesi(gecerli_xlsx)[0]
        with pytest.raises((AttributeError, TypeError)):
            personel.tckn = "00000000000"  # type: ignore[misc]

    def test_gecersiz_tckn_filtelenir(self, tmp_path: Path):
        """Geçersiz TCKN içeren satır atlanmalı."""
        dosya = tmp_path / "test_gecersiz.xlsx"
        _xlsx_yaz(
            [
                {"TCKN": "00000000000", "AD SOYAD": "Geçersiz Kişi", "BİRİMİ": "Birim A"},
                {"TCKN": "10000000146", "AD SOYAD": "Geçerli Kişi", "BİRİMİ": "Birim B"},
            ],
            dosya,
        )
        personeller = oku_personel_listesi(dosya)
        assert len(personeller) == 1
        assert personeller[0].ad_soyad == "Geçerli Kişi"

    def test_bos_satir_filtelenir(self, tmp_path: Path):
        """TCKN veya Ad Soyad boş olan satırlar atlanmalı."""
        dosya = tmp_path / "test_bos.xlsx"
        _xlsx_yaz(
            [
                {"TCKN": None, "AD SOYAD": "Birisi", "BİRİMİ": "Birim"},
                {"TCKN": "10000000146", "AD SOYAD": None, "BİRİMİ": "Birim"},
                {"TCKN": "10000000078", "AD SOYAD": "Geçerli", "BİRİMİ": "Birim"},
            ],
            dosya,
        )
        personeller = oku_personel_listesi(dosya)
        assert len(personeller) == 1

    def test_dosya_bulunamazsa_hata(self, tmp_path: Path):
        """Var olmayan dosya için FileNotFoundError fırlatılmalı."""
        with pytest.raises(FileNotFoundError):
            oku_personel_listesi(tmp_path / "yok.xlsx")

    def test_eksik_sutun_hata(self, tmp_path: Path):
        """Zorunlu sütun eksikse ValueError fırlatılmalı."""
        dosya = tmp_path / "test_eksik.xlsx"
        df = pd.DataFrame(
            [{"TCKN": "10000000146", "AD SOYAD": "Birisi"}]  # BİRİMİ eksik
        )
        df.to_excel(dosya, index=False)
        with pytest.raises(ValueError, match="BİRİMİ"):
            oku_personel_listesi(dosya)

    def test_tckn_sayisal_normalize(self, tmp_path: Path):
        """Excel'den sayısal olarak okunan TCKN normalize edilmeli."""
        # openpyxl ile sayısal değer yaz, pandas float/int olarak okur
        dosya = tmp_path / "test_numeric_tckn.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["TCKN", "AD SOYAD", "BİRİMİ"])
        ws.append([10000000146, "Test Kişi", "Birim"])
        wb.save(dosya)
        personeller = oku_personel_listesi(dosya)
        assert len(personeller) == 1
        assert personeller[0].tckn == "10000000146"
