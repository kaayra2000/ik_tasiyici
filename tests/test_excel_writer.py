"""
excel_writer modülü için entegrasyon testleri.

xlsx bellekte oluşturulur (BytesIO ile), openpyxl ile açılıp
sayfa sayısı, sayfa adları ve hücre değerleri doğrulanır.
"""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from src.core.excel_reader import Personel
from src.core.excel_writer import (
    _sayfa_adi_olustur,
    olustur_dk_bytes,
)
from src.core.excel_write_strategy_v1 import ExcelWriteStrategyV1
from src.config.constants import TECRUBE_BASLANGIC_SATIR

_HUCRE_UNVAN = ExcelWriteStrategyV1.HUCRE_UNVAN


# ---------------------------------------------------------------------------
# Fixture'lar
# ---------------------------------------------------------------------------


@pytest.fixture()
def tek_personel() -> list[Personel]:
    return [Personel(tckn="10000000146", ad_soyad="Fatma KARACA", birim="Marmara Enstitüsü")]


@pytest.fixture()
def uc_personel() -> list[Personel]:
    return [
        Personel(tckn="10000000146", ad_soyad="Fatma KARACA", birim="Marmara Enstitüsü"),
        Personel(tckn="10000000078", ad_soyad="Ali YILMAZ", birim="Gebze Enstitüsü"),
        Personel(tckn="10000050028", ad_soyad="Ayşe DEMİR", birim="Kocaeli Enstitüsü"),
    ]


def _workbook_from_bytes(data: bytes) -> openpyxl.Workbook:
    """Bayt verisinden openpyxl Workbook oluşturur."""
    return openpyxl.load_workbook(BytesIO(data), data_only=False)


# ---------------------------------------------------------------------------
# _sayfa_adi_olustur testleri
# ---------------------------------------------------------------------------


class TestSayfaAdiOlustur:
    """_sayfa_adi_olustur fonksiyonu için testler."""

    def test_normal_ad(self):
        p = Personel(tckn="10000000146", ad_soyad="Ali VELİ", birim="Birim")
        sonuc = _sayfa_adi_olustur(p)
        assert "Ali VELİ" in sonuc
        assert "10000000146" in sonuc

    def test_max_31_karakter(self):
        """Sayfa adı 31 karakteri geçmemeli."""
        p = Personel(
            tckn="10000000146",
            ad_soyad="Çok Uzun Bir İsim Soyisim Örneği",
            birim="Birim",
        )
        sonuc = _sayfa_adi_olustur(p)
        assert len(sonuc) <= 31

    def test_gecersiz_karakterler_temizlenir(self):
        """Geçersiz Excel sayfa karakterleri temizlenmeli."""
        p = Personel(tckn="10000000146", ad_soyad="Ali/Veli", birim="Birim")
        sonuc = _sayfa_adi_olustur(p)
        assert "/" not in sonuc


# ---------------------------------------------------------------------------
# olustur_dk_bytes testleri
# ---------------------------------------------------------------------------


class TestOlusturDkBytes:
    """olustur_dk_bytes fonksiyonu için entegrasyon testleri."""

    def test_tek_personel_tek_sayfa(self, tek_personel):
        """Bir personel için bir sayfa oluşturulmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        gorunen = [s for s in wb.sheetnames if not wb[s].sheet_state == "hidden"]
        assert len(gorunen) == 1

    def test_uc_personel_uc_sayfa(self, uc_personel):
        """Üç personel için üç sayfa oluşturulmalı."""
        data = olustur_dk_bytes(uc_personel)
        wb = _workbook_from_bytes(data)
        gorunen = [s for s in wb.sheetnames if wb[s].sheet_state != "hidden"]
        assert len(gorunen) == 3

    def test_gecerli_xlsx_uretilir(self, tek_personel):
        """Üretilen bytes geçerli xlsx formatında olmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        assert wb is not None

    def test_ad_soyad_hucrede(self, tek_personel):
        """Ad Soyad değeri B3 hücresine yazılmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        assert ws["B3"].value == "Fatma KARACA"

    def test_tckn_hucrede(self, tek_personel):
        """TCKN değeri C3 hücresine yazılmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        assert ws["C3"].value == "10000000146"

    def test_birim_hucrede(self, tek_personel):
        """Birim değeri D3 hücresine yazılmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        assert ws["D3"].value == "Marmara Enstitüsü"

    def test_tecrube_formulleri_var(self, tek_personel):
        """Tecrübe satırlarında K sütununda formül olmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        k10 = ws.cell(row=TECRUBE_BASLANGIC_SATIR, column=11).value
        assert k10 is not None
        assert str(k10).startswith("=")

    def test_alanda_formulleri_var(self, tek_personel):
        """Tecrübe satırlarında L sütununda formül olmalı."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        l10 = ws.cell(row=TECRUBE_BASLANGIC_SATIR, column=12).value
        assert l10 is not None
        assert str(l10).startswith("=")

    def test_sayfa_adi_tckn_icerir(self, tek_personel):
        """Sayfa adı TCKN numarasını içermeli."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        assert any("10000000146" in name for name in wb.sheetnames)

    def test_bos_liste_sayfa_kleri_bos(self):
        """Boş liste verildiğinde personel sayfası oluşturulmamalı, sadece _bos yer tutucu olmalı."""
        data = olustur_dk_bytes([])
        wb = _workbook_from_bytes(data)
        # Sayfa adlarında hiçbir geçerli TCKN olmamalı
        personel_sayfalari = [
            s for s in wb.sheetnames
            if s != "_bos"
        ]
        assert len(personel_sayfalari) == 0

    def test_unvan_formul_yazildi(self, tek_personel):
        """Ünvan hücresi formül içermeli."""
        data = olustur_dk_bytes(tek_personel)
        wb = _workbook_from_bytes(data)
        ws = wb.worksheets[0]
        unvan_hucre = ws[_HUCRE_UNVAN].value
        assert unvan_hucre is not None
        assert str(unvan_hucre).startswith("=")
