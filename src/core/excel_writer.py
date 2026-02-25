"""
Çıktı Excel dosyasını oluşturan modül.

Her personel için ``EN YENİ DK format1 v2.xlsx`` formatında bir sayfa oluşturur
ve tüm sayfaları tek bir ``DK_Tutanaklari_2026.xlsx`` dosyasında birleştirir.

Hücre rolleri:
    - ``o`` (otomatik) → Python tarafından doldurulur
    - ``h`` (hesaplanacak) → Excel formülü yazılır
    - ``e`` (elle) → boş bırakılır

.. note::
    Bu modül gerçek şablon dosyasına bağımlı değildir; karar tutanağı düzenini
    openpyxl ile programatik olarak oluşturur. Şablon entegrasyonu
    Faz 3'te eklenebilir.
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
)
from src.core.excel_reader import Personel
from src.core.formula_builder import (
    alanda_prim_formulu,
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
_HUCRE_TCKN = "B4"
_HUCRE_BIRIM = "B5"
_HUCRE_TARIH = "B6"

# Toplam / hesap satırları
_SATIR_TOPLAM_PRIM = TECRUBE_BITIS_SATIR + 1        # 26
_SATIR_ALANDA_PRIM = TECRUBE_BITIS_SATIR + 2        # 27
_SATIR_TECRUBE_YILI = TECRUBE_BITIS_SATIR + 3       # 28
_SATIR_UNVAN = TECRUBE_BITIS_SATIR + 4              # 29
_SATIR_HIZMET_GRUBU = TECRUBE_BITIS_SATIR + 5       # 30
_SATIR_KADEME = TECRUBE_BITIS_SATIR + 6             # 31

_COL_DEGER = "B"
_TECRUBE_YILI_HUCRE = f"{_COL_DEGER}{_SATIR_TECRUBE_YILI}"


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
    """Her personel için bir sayfa eklenmiş Workbook oluşturur."""
    wb = Workbook()
    # Varsayılan boş sayfayı kaldır
    if wb.active is not None:
        del wb[wb.active.title]

    for personel in personeller:
        sayfa_adi = _sayfa_adi_olustur(personel)
        ws = wb.create_sheet(title=sayfa_adi)
        _sayfayi_doldur(ws, personel)

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
    Bir çalışma sayfasını karar tutanağı formatında doldurur.

    :param ws: Doldurulacak sayfa.
    :param personel: Sayfa için personel bilgisi.
    """
    _yaz_baslik(ws)
    _doldur_otomatik(ws, personel)
    _yaz_ogrenim_etiketleri(ws)
    _yaz_tecrube_basliklari(ws)
    _yaz_tecrube_satirlari(ws)
    _yaz_hesap_satirlari(ws)


def _yaz_baslik(ws) -> None:
    """Sayfa başlığını yazar."""
    ws["A1"] = "DERECE-KADEME (D-K) KARAR TUTANAĞI"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A1:M1")


def _doldur_otomatik(ws, personel: Personel) -> None:
    """Otomatik (o) alanları personel verisinden doldurur."""
    bugun = datetime.date.today().strftime("%d.%m.%Y")
    _etiket_deger_yaz(ws, "A3", "Ad Soyad", _HUCRE_AD_SOYAD, personel.ad_soyad)
    _etiket_deger_yaz(ws, "A4", "TC Kimlik No", _HUCRE_TCKN, personel.tckn)
    _etiket_deger_yaz(ws, "A5", "Birimi", _HUCRE_BIRIM, personel.birim)
    _etiket_deger_yaz(ws, "A6", "Tutanak Tarihi", _HUCRE_TARIH, bugun)


def _yaz_ogrenim_etiketleri(ws) -> None:
    """Eğitim bilgileri bölüm başlığını ve satır etiketlerini yazar (satır 7–8)."""
    ws["A7"] = "ÖĞRENİM BİLGİLERİ"
    ws["A7"].font = Font(bold=True)
    ws["C7"] = "Okul/Bölüm (elle girilecek)"
    ws["D7"] = "Alanında (E/H)"
    ws["A8"] = "Lisans"
    ws.cell(row=9, column=1).value = "Tezsiz Yüksek Lisans"
    # Not: satır 10 tecrübe başlık satırı (row 9) ile çakışmaz —
    # TECRUBE_BASLANGIC_SATIR=10, başlık satırı=9


def _yaz_tecrube_basliklari(ws) -> None:
    """
    Mesleki tecrübe sütun başlıklarını satır 9'a yazar.

    Satır 10-25 tecrübe veri satırları için ayrılmıştır.
    Merge kullanılmaz çünkü merge, veri hücreleriyle çakışmaya sebep olur.
    """
    baslik_satiri = TECRUBE_BASLANGIC_SATIR - 1  # 9

    sutunlar = {
        1: "No",
        2: "Kurum/Görev",
        4: "Başlangıç",
        5: "Bitiş",
        10: "Alanında",
        11: "Toplam Prim Günü",
        12: "Alanda Prim Günü",
    }
    for col, etiket in sutunlar.items():
        hucre = ws.cell(row=baslik_satiri, column=col)
        hucre.value = etiket
        hucre.font = Font(bold=True)


def _yaz_tecrube_satirlari(ws) -> None:
    """Her mesleki tecrübe satırı için sıra no ve Excel formüllerini yazar."""
    for satir in range(TECRUBE_BASLANGIC_SATIR, TECRUBE_BITIS_SATIR + 1):
        ws.cell(row=satir, column=1).value = satir - TECRUBE_BASLANGIC_SATIR + 1
        ws.cell(row=satir, column=11).value = prim_gunu_formulu(satir)
        ws.cell(row=satir, column=12).value = alanda_prim_formulu(satir)


def _yaz_hesap_satirlari(ws) -> None:
    """Toplam, tecrübe yılı, ünvan, hizmet grubu ve kademe satırlarını yazar."""
    alanda_hucre = f"L{_SATIR_ALANDA_PRIM}"

    ws.cell(row=_SATIR_TOPLAM_PRIM, column=1).value = "Toplam Prim Günü"
    ws.cell(row=_SATIR_TOPLAM_PRIM, column=11).value = toplam_prim_formulu()

    ws.cell(row=_SATIR_ALANDA_PRIM, column=1).value = "Alanda Toplam Prim Günü"
    ws.cell(row=_SATIR_ALANDA_PRIM, column=12).value = toplam_alanda_prim_formulu()

    ws.cell(row=_SATIR_TECRUBE_YILI, column=1).value = "Tecrübe Yılı"
    ws.cell(row=_SATIR_TECRUBE_YILI, column=2).value = tecrube_yili_formulu(alanda_hucre)

    ws.cell(row=_SATIR_UNVAN, column=1).value = "Ünvan"
    ws.cell(row=_SATIR_UNVAN, column=2).value = unvan_formulu(_TECRUBE_YILI_HUCRE)

    ws.cell(row=_SATIR_HIZMET_GRUBU, column=1).value = "Hizmet Grubu"
    ws.cell(row=_SATIR_HIZMET_GRUBU, column=2).value = hizmet_grubu_formulu(_TECRUBE_YILI_HUCRE)

    # C8 = Lisans hücresi; kullanıcı değiştirebilir
    ws.cell(row=_SATIR_KADEME, column=1).value = "Kademe"
    ws.cell(row=_SATIR_KADEME, column=2).value = kademe_formulu(_TECRUBE_YILI_HUCRE, "C8")


def _etiket_deger_yaz(ws, etiket_hucre: str, etiket: str, deger_hucre: str, deger: str) -> None:
    """Etiket ve değeri iki ayrı hücreye yazar."""
    ws[etiket_hucre] = etiket
    ws[etiket_hucre].font = Font(bold=True)
    ws[deger_hucre] = deger
