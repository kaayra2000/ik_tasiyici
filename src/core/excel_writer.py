"""
Çıktı Excel dosyasını oluşturan modül.

Her personel için ``cikti_ornegi.xlsx`` şablonunu kopyalarak bir sayfa oluşturur
ve tüm sayfaları tek bir ``DK_Tutanaklari_2026.xlsx`` dosyasında birleştirir.

Sayfa doldurma mantığı **Strategy Pattern** ile soyutlanmıştır.
Farklı şablon versiyonları ``ExcelWriteStrategy`` arayüzünü implemente eden
sınıflar olarak tanımlanır ve ``ExcelWriterFactory`` aracılığıyla yaratılır.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import List
from copy import copy

import openpyxl
from openpyxl import Workbook

from src.config.constants import (
    DEFAULT_VERSION,
    MAX_SHEET_NAME_LEN,
    OUTPUT_FILENAME,
    TEMPLATE_PATH,
)
from src.core.excel_reader import Personel
from src.core.excel_write_strategy import ExcelWriteStrategy
from src.core.excel_writer_factory import ExcelWriterFactory


# ---------------------------------------------------------------------------
# Ana yazma fonksiyonları
# ---------------------------------------------------------------------------


def olustur_dk_dosyasi(
    personeller: List[Personel],
    cikti_dizini: str | Path = ".",
    dosya_adi: str = OUTPUT_FILENAME,
    template_path: str | Path | None = None,
    version: str = DEFAULT_VERSION,
) -> Path:
    """
    Personel listesinden DK Tutanağı Excel dosyasını oluşturur.

    :param personeller: İşlenecek personel listesi.
    :param cikti_dizini: Çıktı dosyasının kaydedileceği dizin.
    :param dosya_adi: Çıktı dosya adı.
    :param template_path: Özel çıktı şablonu yolu (opsiyonel).
    :param version: Çıktı versiyonu (ör. ``"v1"``).
    :returns: Oluşturulan dosyanın tam yolu.
    """
    strategy = ExcelWriterFactory.create(version)
    wb = _workbook_olustur(personeller, strategy, template_path)
    cikti_dizini = Path(cikti_dizini)
    cikti_dizini.mkdir(parents=True, exist_ok=True)
    cikti_yolu = cikti_dizini / dosya_adi
    wb.save(cikti_yolu)
    return cikti_yolu


def olustur_dk_bytes(
    personeller: List[Personel],
    template_path: str | Path | None = None,
    version: str = DEFAULT_VERSION,
) -> bytes:
    """
    Personel listesinden DK Tutanağı Excel dosyasını bellekte oluşturur.

    Test ve ön izleme amacıyla kullanışlıdır.

    :param personeller: İşlenecek personel listesi.
    :param template_path: Özel çıktı şablonu yolu (opsiyonel).
    :param version: Çıktı versiyonu (ör. ``"v1"``).
    :returns: xlsx içeriği bayt dizisi olarak.
    """
    strategy = ExcelWriterFactory.create(version)
    wb = _workbook_olustur(personeller, strategy, template_path)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# İç yardımcılar
# ---------------------------------------------------------------------------


def _workbook_olustur(
    personeller: List[Personel],
    strategy: ExcelWriteStrategy,
    template_path: str | Path | None = None,
) -> Workbook:
    """Her personel için şablon sayfasından kopyalanmış Workbook oluşturur."""
    # template path
    if template_path is None:
        project_root = Path(__file__).resolve().parent.parent.parent
        t_path = project_root / TEMPLATE_PATH
    else:
        t_path = Path(template_path)

    wb = openpyxl.load_workbook(t_path)
    template_ws = wb.active
    template_ws_title = template_ws.title

    for personel in personeller:
        sayfa_adi = _sayfa_adi_olustur(personel)
        ws = wb.copy_worksheet(template_ws)
        ws.title = sayfa_adi

        # Şablondaki eksik kalan önemli biçimleri/kuralları manuel aktar
        ws.conditional_formatting = copy(template_ws.conditional_formatting)
        if hasattr(template_ws, 'data_validations'):
            ws.data_validations = copy(template_ws.data_validations)

        ws.print_area = template_ws.print_area
        ws.print_options = copy(template_ws.print_options)
        ws.page_setup = copy(template_ws.page_setup)
        ws.freeze_panes = template_ws.freeze_panes

        # Garanti olması için sütun/satır genişliklerini/yüksekliklerini aktar
        for row_idx, row_dim in template_ws.row_dimensions.items():
            ws.row_dimensions[row_idx] = copy(row_dim)
            ws.row_dimensions[row_idx].worksheet = ws
        for col_idx, col_dim in template_ws.column_dimensions.items():
            ws.column_dimensions[col_idx] = copy(col_dim)
            ws.column_dimensions[col_idx].worksheet = ws

        # Strateji aracılığıyla sayfayı doldur
        strategy.sayfa_doldur(ws, personel)

    # Orijinal şablon sayfasını kaldır
    if template_ws_title in wb.sheetnames:
        del wb[template_ws_title]

    # openpyxl en az 1 sayfa olmadan kaydedemez;
    # boş liste durumunda boş bir kılavuz sayfası ekleriz.
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
