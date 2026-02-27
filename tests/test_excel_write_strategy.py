"""
ExcelWriteStrategy, ExcelWriteStrategyV1 ve ExcelWriterFactory testleri.

Strategy Pattern + Factory Method'un doğru çalıştığını doğrular.
"""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from src.core.excel_reader import Personel
from src.core.excel_write_strategy import ExcelWriteStrategy
from src.core.excel_write_strategy_v1 import ExcelWriteStrategyV1
from src.core.excel_writer_factory import ExcelWriterFactory


# ---------------------------------------------------------------------------
# Factory testleri
# ---------------------------------------------------------------------------


class TestExcelWriterFactory:
    """ExcelWriterFactory sınıfı testleri."""

    def test_create_v1(self):
        """v1 versiyonu ExcelWriteStrategyV1 döndürmeli."""
        strategy = ExcelWriterFactory.create("v1")
        assert isinstance(strategy, ExcelWriteStrategyV1)

    def test_create_v1_is_strategy(self):
        """v1 versiyonu ExcelWriteStrategy arayüzünü implemente etmeli."""
        strategy = ExcelWriterFactory.create("v1")
        assert isinstance(strategy, ExcelWriteStrategy)

    def test_create_invalid_version(self):
        """Geçersiz versiyon ValueError fırlatmalı."""
        with pytest.raises(ValueError, match="Desteklenmeyen"):
            ExcelWriterFactory.create("v99")

    def test_create_empty_version(self):
        """Boş versiyon ValueError fırlatmalı."""
        with pytest.raises(ValueError):
            ExcelWriterFactory.create("")


# ---------------------------------------------------------------------------
# V1 Strategy testleri
# ---------------------------------------------------------------------------


class TestExcelWriteStrategyV1:
    """ExcelWriteStrategyV1 sınıfı testleri."""

    @pytest.fixture()
    def strategy(self) -> ExcelWriteStrategyV1:
        return ExcelWriteStrategyV1()

    @pytest.fixture()
    def personel(self) -> Personel:
        return Personel(
            tckn="10000000146",
            ad_soyad="Fatma KARACA",
            birim="Marmara Enstitüsü",
        )

    @pytest.fixture()
    def template_ws(self):
        """Test için basit bir şablon çalışma sayfası oluşturur."""
        wb = openpyxl.Workbook()
        ws = wb.active
        # Minimum şablon yapısı – strateji bunu dolduracak
        return ws

    def test_sayfa_doldur_ad_soyad(self, strategy, personel, template_ws):
        """Ad soyad B3 hücresine yazılmalı."""
        strategy.sayfa_doldur(template_ws, personel)
        assert template_ws["B3"].value == "Fatma KARACA"

    def test_sayfa_doldur_tckn(self, strategy, personel, template_ws):
        """TCKN C3 hücresine yazılmalı."""
        strategy.sayfa_doldur(template_ws, personel)
        assert template_ws["C3"].value == "10000000146"

    def test_sayfa_doldur_birim(self, strategy, personel, template_ws):
        """Birim D3 hücresine yazılmalı."""
        strategy.sayfa_doldur(template_ws, personel)
        assert template_ws["D3"].value == "Marmara Enstitüsü"

    def test_sayfa_doldur_unvan_formul(self, strategy, personel, template_ws):
        """Ünvan hücresi formül içermeli."""
        strategy.sayfa_doldur(template_ws, personel)
        unvan = template_ws["E3"].value
        assert unvan is not None
        assert str(unvan).startswith("=")

    def test_sayfa_doldur_kademe_formul(self, strategy, personel, template_ws):
        """Kademe hücresi formül içermeli."""
        strategy.sayfa_doldur(template_ws, personel)
        kademe = template_ws["F3"].value
        assert kademe is not None
        assert str(kademe).startswith("=")

    def test_sayfa_doldur_tecrube_formulleri(self, strategy, personel, template_ws):
        """Tecrübe satırlarında formül olmalı."""
        strategy.sayfa_doldur(template_ws, personel)
        from src.config.constants import TECRUBE_BASLANGIC_SATIR
        k_val = template_ws.cell(row=TECRUBE_BASLANGIC_SATIR, column=11).value
        assert k_val is not None
        assert str(k_val).startswith("=")

    def test_is_abstract_strategy_subclass(self):
        """V1 stratejisi ExcelWriteStrategy alt sınıfı olmalı."""
        assert issubclass(ExcelWriteStrategyV1, ExcelWriteStrategy)
