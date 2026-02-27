"""
Çıktı Excel dosyasını oluşturan modül.

Her personel için ``cikti_ornegi.xlsx`` şablonunu kopyalarak bir sayfa oluşturur
ve tüm sayfaları tek bir ``DK_Tutanaklari_2026.xlsx`` dosyasında birleştirir.
"""

from __future__ import annotations

import datetime
from io import BytesIO
from pathlib import Path
from typing import List

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from src.config.constants import (
    MAX_SHEET_NAME_LEN,
    OUTPUT_FILENAME,
    TECRUBE_BASLANGIC_SATIR,
    TECRUBE_BITIS_SATIR,
    TEMPLATE_PATH,
)
from src.core.excel_reader import Personel
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
# Sabitler – şablondaki hücre adresleri
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

# Z sütununu gizli hesaplama için kullanacağız:
# Z1 = Tecrübe Yılı, Z2 = Hizmet Grubu, Z3 = Kademe, Z4 = En Yüksek Öğrenim (alanında)
_TECRUBE_YILI_HUCRE = "Z1"
_EN_YUKSEK_OGRENIM_HUCRE = "Z4"  # kademe formülü bu hücreye bakacak

# Şablondaki öğrenim satırları parametreleri (B=Ad, C=Okul, K=Alanında)
_OGRENIM_BAS_SATIR = 6
_OGRENIM_BIT_SATIR = 8
_OGRENIM_AD_SUTUN = "B"
_OGRENIM_OKUL_SUTUN = "C"
_OGRENIM_ALANINDA_SUTUN = "K"


# ---------------------------------------------------------------------------
# Ana yazma fonksiyonları
# ---------------------------------------------------------------------------


def olustur_dk_dosyasi(
    personeller: List[Personel],
    cikti_dizini: str | Path = ".",
    dosya_adi: str = OUTPUT_FILENAME,
) -> Path:
    """
    Personel listesinden DK Tutanağı Excel dosyasını oluşturur.

    :param personeller: İşlenecek personel listesi.
    :param cikti_dizini: Çıktı dosyasının kaydedileceği dizin.
    :param dosya_adi: Çıktı dosya adı.
    :returns: Oluşturulan dosyanın tam yolu.
    """
    wb = _workbook_olustur(personeller)
    cikti_dizini = Path(cikti_dizini)
    cikti_dizini.mkdir(parents=True, exist_ok=True)
    cikti_yolu = cikti_dizini / dosya_adi
    wb.save(cikti_yolu)
    return cikti_yolu


def olustur_dk_bytes(personeller: List[Personel]) -> bytes:
    """
    Personel listesinden DK Tutanağı Excel dosyasını bellekte oluşturur.

    Test ve ön izleme amacıyla kullanışlıdır.

    :param personeller: İşlenecek personel listesi.
    :returns: xlsx içeriği bayt dizisi olarak.
    """
    wb = _workbook_olustur(personeller)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# İç yardımcılar
# ---------------------------------------------------------------------------


def _workbook_olustur(personeller: List[Personel]) -> Workbook:
    """Her personel için şablon sayfasından kopyalanmış Workbook oluşturur."""
    # template path
    project_root = Path(__file__).resolve().parent.parent.parent
    template_path = project_root / TEMPLATE_PATH
    
    wb = openpyxl.load_workbook(template_path)
    template_ws = wb.active
    template_ws_title = template_ws.title

    for personel in personeller:
        sayfa_adi = _sayfa_adi_olustur(personel)
        ws = wb.copy_worksheet(template_ws)
        ws.title = sayfa_adi
        _sayfayi_doldur(ws, personel)

    # Orijinal şablon sayfasını kaldır
    if template_ws_title in wb.sheetnames:
        del wb[template_ws_title]

    # openpyxl en az 1 sayfa olmadan kaydedemez;
    # boş liste durumunda boş bir kılavuz sayfası ekleriz (görünür olması zorunlu).
    if not wb.sheetnames:
        wb.create_sheet("_bos")

    return wb


def _sayfa_adi_olustur(personel: Personel) -> str:
    """
    Personel için Excel sayfa adını oluşturur.

    Format: ``{Ad Soyad} - {TCKN}`` (max 31 karakter).

    :param personel: Personel nesnesi.
    :returns: Excel sayfa adı.
    """
    tam_ad = f"{personel.ad_soyad} - {personel.tckn}"
    if len(tam_ad) > MAX_SHEET_NAME_LEN:
        kisaltilmis = personel.ad_soyad[: MAX_SHEET_NAME_LEN - len(personel.tckn) - 5]
        tam_ad = f"{kisaltilmis}.. - {personel.tckn}"
    for karakter in r"\/?*[]":
        tam_ad = tam_ad.replace(karakter, "")
    return tam_ad[:MAX_SHEET_NAME_LEN]


def _sayfayi_doldur(ws, personel: Personel) -> None:
    """
    Kopyalanmış şablon çalışma sayfasına formülleri ve personel verisini doldurur.
    """
    _doldur_otomatik(ws, personel)
    _yaz_tecrube_satirlari(ws)
    _yaz_hesap_satirlari(ws)


def _doldur_otomatik(ws, personel: Personel) -> None:
    """Otomatik (o) alanları personel verisinden doldurur."""
    ws[_HUCRE_AD_SOYAD] = personel.ad_soyad
    ws[_HUCRE_TCKN] = personel.tckn
    ws[_HUCRE_BIRIM] = personel.birim


def _yaz_tecrube_satirlari(ws) -> None:
    """Her mesleki tecrübe satırı için Excel formüllerini yazar."""
    for satir in range(TECRUBE_BASLANGIC_SATIR, TECRUBE_BITIS_SATIR + 1):
        ws.cell(row=satir, column=11).value = prim_gunu_formulu(satir)
        ws.cell(row=satir, column=12).value = alanda_prim_formulu(satir)


def _yaz_hesap_satirlari(ws) -> None:
    """Toplam, tecrübe yılı, ünvan, hizmet grubu ve kademe satırlarını yazar."""
    toplam_alanda_hucre = f"L{_SATIR_ALANDA_PRIM}"

    # Toplam Prim Günü (K19) ve Alanda Toplam Prim Günü (L19)
    ws.cell(row=_SATIR_TOPLAM_PRIM, column=11).value = toplam_prim_formulu()
    ws.cell(row=_SATIR_ALANDA_PRIM, column=12).value = toplam_alanda_prim_formulu()

    # Z sütununa gizli formülleri yazalım:
    # Z1 = Tecrübe Yılı (alanda toplam prim / 360)
    ws["Z1"] = tecrube_yili_formulu(toplam_alanda_hucre)

    # Z4 = En yüksek alanında öğrenim (B sütunundaki hücreden ismi okur)
    ws[_EN_YUKSEK_OGRENIM_HUCRE] = en_yuksek_ogrenim_formulu(
        baslangic_satir=_OGRENIM_BAS_SATIR,
        bitis_satir=_OGRENIM_BIT_SATIR,
        ad_sutun=_OGRENIM_AD_SUTUN,
        okul_sutun=_OGRENIM_OKUL_SUTUN,
        alaninda_sutun=_OGRENIM_ALANINDA_SUTUN,
    )

    # Z2 = Hizmet Grubu (A/AG-2 ... A/AG-6)
    ws["Z2"] = hizmet_grubu_formulu(_TECRUBE_YILI_HUCRE)

    # Z3 = Kademe (tecrübe yılı + en yüksek öğrenim kombinasyonu)
    ws["Z3"] = kademe_formulu(_TECRUBE_YILI_HUCRE, _EN_YUKSEK_OGRENIM_HUCRE)

    # E3: Ünvan
    ws[_HUCRE_UNVAN] = unvan_formulu(_TECRUBE_YILI_HUCRE)

    # F3: Derece/Kademe (Hizmet Grubu / Kademe)
    # Eğer Kademe (Z3) boş dönerse sadece Hizmet Grubu (Z2) yazılır
    ws[_HUCRE_KADEME] = '=IF(Z3="", Z2, Z2 & "/" & Z3)'
