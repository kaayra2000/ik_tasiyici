"""
Çıktı Excel dosyasını oluşturan modül.

Her personel için ``cikti_ornegi.xlsx`` şablonunu kopyalarak bir sayfa oluşturur
ve tüm sayfaları tek bir ``DK_Tutanaklari_2026.xlsx`` dosyasında birleştirir.

Sayfa doldurma mantığı **Strategy Pattern** ile soyutlanmıştır.
Farklı şablon versiyonları ``ExcelWriteStrategy`` arayüzünü implemente eden
sınıflar olarak tanımlanır ve ``ExcelWriterFactory`` aracılığıyla yaratılır.
"""

from __future__ import annotations

from dataclasses import dataclass, field
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


@dataclass(frozen=True)
class TutanakOlusturmaRaporu:
    """Tutanak oluşturma işleminin özetini taşır."""

    output_path: Path
    added_sheet_count: int = 0
    skipped_existing_count: int = 0
    warning_messages: list[str] = field(default_factory=list)


def olustur_dk_dosyasi(
    personeller: List[Personel],
    cikti_dizini: str | Path = ".",
    dosya_adi: str = OUTPUT_FILENAME,
    template_path: str | Path | None = None,
    version: str = DEFAULT_VERSION,
) -> Path:
    """
    Personel listesinden DK Tutanağı Excel dosyasını oluşturur.

    Hedef dosya zaten varsa mevcut sayfalar korunur; yalnızca workbook'ta
    bulunmayan personeller yeni sayfalar olarak sona eklenir.

    :param personeller: İşlenecek personel listesi.
    :param cikti_dizini: Çıktı dosyasının kaydedileceği dizin.
    :param dosya_adi: Çıktı dosya adı.
    :param template_path: Özel çıktı şablonu yolu (opsiyonel).
    :param version: Çıktı versiyonu (ör. ``"v1"``).
    :returns: Oluşturulan dosyanın tam yolu.
    """
    return olustur_dk_dosyasi_raporlu(
        personeller=personeller,
        cikti_dizini=cikti_dizini,
        dosya_adi=dosya_adi,
        template_path=template_path,
        version=version,
    ).output_path


def olustur_dk_dosyasi_raporlu(
    personeller: List[Personel],
    cikti_dizini: str | Path = ".",
    dosya_adi: str = OUTPUT_FILENAME,
    template_path: str | Path | None = None,
    version: str = DEFAULT_VERSION,
) -> TutanakOlusturmaRaporu:
    """
    Personel listesinden DK Tutanağı üretip ayrıntılı işlem raporu döner.

    :param personeller: İşlenecek personel listesi.
    :param cikti_dizini: Çıktı dosyasının kaydedileceği dizin.
    :param dosya_adi: Çıktı dosya adı.
    :param template_path: Özel çıktı şablonu yolu (opsiyonel).
    :param version: Çıktı versiyonu (ör. ``"v1"``).
    :returns: Oluşturma özeti.
    """
    strategy = ExcelWriterFactory.create(version)
    cikti_dizini = Path(cikti_dizini)
    cikti_dizini.mkdir(parents=True, exist_ok=True)
    cikti_yolu = cikti_dizini / dosya_adi

    if cikti_yolu.exists():
        wb = openpyxl.load_workbook(cikti_yolu)
        added_sheet_count, skipped_existing_count, warning_messages = (
            _personelleri_workbooka_ekle(wb, personeller, strategy, template_path)
        )
    else:
        wb, added_sheet_count, skipped_existing_count, warning_messages = (
            _workbook_olustur(personeller, strategy, template_path)
        )

    wb.save(cikti_yolu)
    return TutanakOlusturmaRaporu(
        output_path=cikti_yolu,
        added_sheet_count=added_sheet_count,
        skipped_existing_count=skipped_existing_count,
        warning_messages=warning_messages,
    )


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
    wb, _, _, _ = _workbook_olustur(personeller, strategy, template_path)
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
) -> tuple[Workbook, int, int, list[str]]:
    """Her personel için şablon sayfasından kopyalanmış Workbook oluşturur."""
    wb = Workbook()
    added_sheet_count, skipped_existing_count, warning_messages = (
        _personelleri_workbooka_ekle(wb, personeller, strategy, template_path)
    )

    # openpyxl en az 1 sayfa olmadan kaydedemez;
    # Eğer hiç personel sayfası oluşturulmadıysa varsayılan aktif sayfayı
    # silmek yerine `_bos` olarak yeniden adlandırıyoruz. Bu koruma
    # openpyxl'in workbook seviyesindeki varsayılan stillerinin korunmasını
    # sağlar (uyarıyı ortadan kaldırır) ve önceki davranışı taklit eder.
    if not wb.sheetnames:
        wb.create_sheet("_bos")
    elif added_sheet_count == 0 and len(wb.sheetnames) == 1 and wb.sheetnames[0] == "Sheet":
        wb.active.title = "_bos"

    return wb, added_sheet_count, skipped_existing_count, warning_messages


def _personelleri_workbooka_ekle(
    wb: Workbook,
    personeller: List[Personel],
    strategy: ExcelWriteStrategy,
    template_path: str | Path | None = None,
) -> tuple[int, int, list[str]]:
    """Workbook'a yalnızca eksik personel sayfalarını sona ekler."""
    template_wb = None
    template_ws = None
    added_sheet_count = 0
    skipped_existing_count = 0
    warning_messages: list[str] = []

    try:
        for personel in personeller:
            sayfa_adi = _sayfa_adi_olustur(personel)
            if sayfa_adi in wb.sheetnames:
                skipped_existing_count += 1
                warning_messages.append(_build_skip_message(personel, sayfa_adi))
                continue

            if template_ws is None:
                template_wb = openpyxl.load_workbook(_template_yolunu_coz(template_path))
                if not template_wb.sheetnames:
                    raise ValueError("Şablon workbook içinde hiç sayfa yok.")
                template_ws = template_wb.active

            # Reuse the default active sheet for the first generated page to
            # preserve workbook-level defaults (openpyxl's default style).
            if (
                len(wb.sheetnames) == 1
                and wb.active is not None
                and not getattr(wb.active, "_cells", {})
            ):
                ws = wb.active
                ws.title = sayfa_adi
            else:
                ws = wb.create_sheet(title=sayfa_adi)
            _sayfa_icerigini_kopyala(template_ws, ws)
            strategy.sayfa_doldur(ws, personel)
            added_sheet_count += 1
    finally:
        if template_wb is not None:
            template_wb.close()

    return added_sheet_count, skipped_existing_count, warning_messages


def _template_yolunu_coz(template_path: str | Path | None = None) -> Path:
    """Şablon dosya yolunu çözümler ve dosyanın varlığını doğrular."""
    if template_path is None:
        project_root = Path(__file__).resolve().parent.parent.parent
        t_path = project_root / TEMPLATE_PATH
    else:
        t_path = Path(template_path)

    if not t_path.exists():
        raise FileNotFoundError(f"Excel şablonu bulunamadı: {t_path}")

    return t_path


def _sayfa_icerigini_kopyala(kaynak_ws, hedef_ws) -> None:
    """Harici workbook'taki şablon sayfasını hedef workbook'a klonlar."""
    for (satir, sutun), kaynak_hucre in kaynak_ws._cells.items():
        hedef_hucre = hedef_ws.cell(row=satir, column=sutun)
        hedef_hucre._value = kaynak_hucre._value
        hedef_hucre.data_type = kaynak_hucre.data_type

        if kaynak_hucre.has_style:
            hedef_hucre.font = copy(kaynak_hucre.font)
            hedef_hucre.fill = copy(kaynak_hucre.fill)
            hedef_hucre.border = copy(kaynak_hucre.border)
            hedef_hucre.alignment = copy(kaynak_hucre.alignment)
            hedef_hucre.number_format = kaynak_hucre.number_format
            hedef_hucre.protection = copy(kaynak_hucre.protection)

        if kaynak_hucre.hyperlink:
            hedef_hucre._hyperlink = copy(kaynak_hucre.hyperlink)

        if kaynak_hucre.comment:
            hedef_hucre.comment = copy(kaynak_hucre.comment)

    for anahtar, boyut in kaynak_ws.row_dimensions.items():
        hedef_boyut = hedef_ws.row_dimensions[anahtar]
        hedef_boyut.height = boyut.height
        hedef_boyut.hidden = boyut.hidden
        hedef_boyut.outlineLevel = boyut.outlineLevel
        hedef_boyut.outline_level = boyut.outline_level
        hedef_boyut.collapsed = boyut.collapsed

    for anahtar, boyut in kaynak_ws.column_dimensions.items():
        hedef_boyut = hedef_ws.column_dimensions[anahtar]
        hedef_boyut.width = boyut.width
        hedef_boyut.hidden = boyut.hidden
        hedef_boyut.bestFit = boyut.bestFit
        hedef_boyut.outlineLevel = boyut.outlineLevel
        hedef_boyut.outline_level = boyut.outline_level
        hedef_boyut.collapsed = boyut.collapsed

    hedef_ws.sheet_format = copy(kaynak_ws.sheet_format)
    hedef_ws.sheet_properties = copy(kaynak_ws.sheet_properties)
    hedef_ws.views = copy(kaynak_ws.views)
    hedef_ws.merged_cells = copy(kaynak_ws.merged_cells)
    hedef_ws.page_margins = copy(kaynak_ws.page_margins)
    hedef_ws.page_setup = copy(kaynak_ws.page_setup)
    hedef_ws.print_options = copy(kaynak_ws.print_options)
    hedef_ws.protection = copy(kaynak_ws.protection)
    hedef_ws.conditional_formatting = copy(kaynak_ws.conditional_formatting)

    if hasattr(kaynak_ws, "data_validations"):
        hedef_ws.data_validations = copy(kaynak_ws.data_validations)

    hedef_ws.freeze_panes = kaynak_ws.freeze_panes
    hedef_ws._print_area = copy(kaynak_ws._print_area)


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


def _build_skip_message(personel: Personel, sayfa_adi: str) -> str:
    """Workbook'ta zaten bulunan kayıt için log mesajı üretir."""
    return (
        "Kayıt atlandı: hedef dosyada zaten mevcut. "
        f"SAYFA='{sayfa_adi}', "
        f"TCKN='{personel.tckn}', "
        f"AD SOYAD='{personel.ad_soyad}', "
        f"BİRİMİ='{personel.birim}'"
    )
