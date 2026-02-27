"""
V1 çıktı şablonu için Excel yazma stratejisi.

``cikti_taslagi.xlsx`` şablonuna uygun hücre adresleri, formül yazma
mantığı ve veri doğrulama kurallarını içerir.
Mevcut ``excel_writer.py`` içindeki iç fonksiyonlardan birebir taşınmıştır.
"""

from __future__ import annotations

from src.config.constants import (
    TECRUBE_BASLANGIC_SATIR,
    TECRUBE_BITIS_SATIR,
)
from src.core.excel_reader import Personel
from src.core.excel_write_strategy import ExcelWriteStrategy
from src.core.formula_builder import (
    alanda_prim_formulu,
    en_yuksek_ogrenim_formulu,
    hizmet_grubu_formulu,
    kademe_formulu,
    prim_gunu_formulu,
    tecrube_yili_formulu,
    toplam_alanda_prim_formulu,
    toplam_prim_formulu,
    unvan_formulu,
)


# ---------------------------------------------------------------------------
# V1 şablonuna ait sabitler – hücre adresleri
# ---------------------------------------------------------------------------

# Otomatik doldurulacak (o) hücreler
_HUCRE_AD_SOYAD = "B3"
_HUCRE_TCKN = "C3"
_HUCRE_BIRIM = "D3"
_HUCRE_UNVAN = "E3"
_HUCRE_KADEME = "F3"

# Toplam / hesap satırları
_SATIR_TOPLAM_PRIM = TECRUBE_BITIS_SATIR + 1        # 19
_SATIR_ALANDA_PRIM = TECRUBE_BITIS_SATIR + 1        # 19

# Z sütununu gizli hesaplama için kullanıyoruz:
# Z1 = Tecrübe Yılı, Z2 = Hizmet Grubu, Z3 = Kademe, Z4 = En Yüksek Öğrenim
_TECRUBE_YILI_HUCRE = "Z1"
_EN_YUKSEK_OGRENIM_HUCRE = "Z4"

# Şablondaki öğrenim satırları parametreleri (B=Ad, C=Okul, K=Alanında)
_OGRENIM_BAS_SATIR = 6
_OGRENIM_BIT_SATIR = 8
_OGRENIM_AD_SUTUN = "B"
_OGRENIM_OKUL_SUTUN = "C"
_OGRENIM_ALANINDA_SUTUN = "K"


class ExcelWriteStrategyV1(ExcelWriteStrategy):
    """V1 (Standart DK Taslağı) için Excel yazma stratejisi.

    ``cikti_taslagi.xlsx`` / ``cikti_taslagi_dolu.xlsx`` şablonlarına
    uygun olarak sayfayı doldurur.
    """

    # V1 sabitlerini dışarıya da açıyoruz (test vb. için)
    HUCRE_UNVAN = _HUCRE_UNVAN

    def sayfa_doldur(self, ws, personel: Personel) -> None:
        """V1 şablonuna göre çalışma sayfasını doldurur."""
        self._doldur_otomatik(ws, personel)
        self._ekle_veri_dogrulama(ws)
        self._yaz_tecrube_satirlari(ws)
        self._yaz_hesap_satirlari(ws)

    # ------------------------------------------------------------------
    # İç yardımcılar
    # ------------------------------------------------------------------

    @staticmethod
    def _doldur_otomatik(ws, personel: Personel) -> None:
        """Otomatik (o) alanları personel verisinden doldurur."""
        ws[_HUCRE_AD_SOYAD] = personel.ad_soyad
        ws[_HUCRE_TCKN] = personel.tckn
        ws[_HUCRE_BIRIM] = personel.birim

    @staticmethod
    def _yaz_tecrube_satirlari(ws) -> None:
        """Her mesleki tecrübe satırı için Excel formüllerini yazar."""
        for satir in range(TECRUBE_BASLANGIC_SATIR, TECRUBE_BITIS_SATIR + 1):
            ws.cell(row=satir, column=11).value = prim_gunu_formulu(satir)
            ws.cell(row=satir, column=12).value = alanda_prim_formulu(satir)

    @staticmethod
    def _yaz_hesap_satirlari(ws) -> None:
        """Toplam, tecrübe yılı, ünvan, hizmet grubu ve kademe satırlarını yazar."""
        toplam_alanda_hucre = f"L{_SATIR_ALANDA_PRIM}"

        # Toplam Prim Günü (K19) ve Alanda Toplam Prim Günü (L19)
        ws.cell(row=_SATIR_TOPLAM_PRIM, column=11).value = toplam_prim_formulu()
        ws.cell(row=_SATIR_ALANDA_PRIM, column=12).value = toplam_alanda_prim_formulu()

        # Z sütununa gizli formülleri yazalım:
        # Z1 = Tecrübe Yılı (alanda toplam prim / 360)
        ws["Z1"] = tecrube_yili_formulu(toplam_alanda_hucre)

        # Z4 = En yüksek alanında öğrenim
        ws[_EN_YUKSEK_OGRENIM_HUCRE] = en_yuksek_ogrenim_formulu(
            baslangic_satir=_OGRENIM_BAS_SATIR,
            bitis_satir=_OGRENIM_BIT_SATIR,
            ad_sutun=_OGRENIM_AD_SUTUN,
            okul_sutun=_OGRENIM_OKUL_SUTUN,
            alaninda_sutun=_OGRENIM_ALANINDA_SUTUN,
        )

        # Z2 = Hizmet Grubu
        ws["Z2"] = hizmet_grubu_formulu(_TECRUBE_YILI_HUCRE)

        # Z3 = Kademe
        ws["Z3"] = kademe_formulu(_TECRUBE_YILI_HUCRE, _EN_YUKSEK_OGRENIM_HUCRE)

        # E3: Ünvan
        ws[_HUCRE_UNVAN] = unvan_formulu(_TECRUBE_YILI_HUCRE)

        # F3: Derece/Kademe
        ws[_HUCRE_KADEME] = '=IF(Z3="", Z2, Z2 & "/" & Z3)'

    @staticmethod
    def _ekle_veri_dogrulama(ws) -> None:
        """B6, B7 ve B8 hücreleri için öğrenim seviyeleri açılır listesini ekler."""
        from openpyxl.worksheet.datavalidation import DataValidation
        from src.config.constants import OGRENIM_SEVIYELERI

        dv = DataValidation(
            type="list",
            formula1=f'"{",".join(OGRENIM_SEVIYELERI)}"',
            allow_blank=True,
        )
        ws.add_data_validation(dv)
        for row in range(6, 9):  # 6, 7, 8
            dv.add(f"B{row}")
