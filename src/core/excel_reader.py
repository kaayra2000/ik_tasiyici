"""
Kaynak Excel dosyasından personel verilerini okuyan modül.

Kaynak dosya örneği: ``coklu_girdi.xlsx``

Kullanılan sütunlar: TCKN, AD SOYAD, BİRİMİ
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

import pandas as pd

from src.config.constants import COL_TCKN, COL_AD_SOYAD, COL_BIRIM
from src.core.validators import normalize_tckn, validate_tckn, validate_ad_soyad

# ---------------------------------------------------------------------------
# Veri sınıfı
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Personel:
    """Bir personele ait temel bilgileri tutan değişmez veri sınıfı."""

    tckn: str
    """11 haneli TC Kimlik Numarası."""

    ad_soyad: str
    """Personelin adı ve soyadı."""

    birim: str
    """Çalıştığı enstitü / birim."""


@dataclass(frozen=True)
class SatirReddi:
    """Geçersiz olduğu için atlanan bir Excel satırını açıklar."""

    excel_satir_no: int
    sebep: str
    tckn: str = ""
    ad_soyad: str = ""
    birim: str = ""

    @property
    def log_mesaji(self) -> str:
        """GUI log'u için kullanıcıya dönük açıklama üretir."""
        alanlar = [
            f"TCKN='{self.tckn or '-'}'",
            f"AD SOYAD='{self.ad_soyad or '-'}'",
            f"BİRİMİ='{self.birim or '-'}'",
        ]
        return f"Satır {self.excel_satir_no} atlandı: {self.sebep}. " + ", ".join(
            alanlar
        )


@dataclass(frozen=True)
class PersonelOkumaRaporu:
    """Excel okuma sonucundaki geçerli kayıtları ve red nedenlerini taşır."""

    personeller: List[Personel]
    reddedilen_satirlar: List[SatirReddi]


# ---------------------------------------------------------------------------
# Okuma fonksiyonları
# ---------------------------------------------------------------------------


def oku_personel_listesi(dosya_yolu: str | Path) -> List[Personel]:
    """
    Kaynak Excel dosyasını okuyarak geçerli personel listesini döner.

    Geçersiz veya eksik satırlar sessizce atlanır.

    :param dosya_yolu: Kaynak xlsx dosyasının yolu.
    :returns: Geçerli :class:`Personel` nesnelerinin listesi.
    :raises FileNotFoundError: Dosya bulunamazsa.
    :raises ValueError: Zorunlu sütunlar eksikse.
    """
    return oku_personel_listesi_raporlu(dosya_yolu).personeller


def oku_personel_listesi_raporlu(dosya_yolu: str | Path) -> PersonelOkumaRaporu:
    """
    Kaynak Excel dosyasını okuyarak geçerli kayıtları ve red nedenlerini döner.

    :param dosya_yolu: Kaynak xlsx dosyasının yolu.
    :returns: Geçerli kayıtlar ve reddedilen satırlar.
    :raises FileNotFoundError: Dosya bulunamazsa.
    :raises ValueError: Zorunlu sütunlar eksikse.
    """
    dosya_yolu = Path(dosya_yolu)
    if not dosya_yolu.exists():
        raise FileNotFoundError(f"Kaynak dosya bulunamadı: {dosya_yolu}")

    df = pd.read_excel(dosya_yolu, dtype=str)

    _zorunlu_sutunları_dogrula(df)

    personeller: List[Personel] = []
    reddedilen_satirlar: List[SatirReddi] = []
    for index, satir in df.iterrows():
        excel_satir_no = int(index) + 2
        personel, red_nedeni = _satiri_isle(satir, excel_satir_no)
        if personel is not None:
            personeller.append(personel)
        if red_nedeni is not None:
            reddedilen_satirlar.append(red_nedeni)

    return PersonelOkumaRaporu(
        personeller=personeller,
        reddedilen_satirlar=reddedilen_satirlar,
    )


# ---------------------------------------------------------------------------
# Yardımcı (dahili) fonksiyonlar
# ---------------------------------------------------------------------------


def _zorunlu_sutunları_dogrula(df: pd.DataFrame) -> None:
    """
    DataFrame'in zorunlu sütunları içerip içermediğini denetler.

    :param df: Okunan veri çerçevesi.
    :raises ValueError: Eksik sütun varsa.
    """
    zorunlu = {COL_TCKN, COL_AD_SOYAD, COL_BIRIM}
    eksik = zorunlu - set(df.columns)
    if eksik:
        raise ValueError(f"Kaynak dosyada zorunlu sütunlar eksik: {eksik}")


def _satiri_isle(
    satir: pd.Series,
    excel_satir_no: int,
) -> tuple[Personel | None, SatirReddi | None]:
    """
    Tek bir DataFrame satırını işler ve geçerliyse :class:`Personel` döner.

    Geçersiz ya da eksik veri içeren satırlar için red nedeni üretir.

    :param satir: İşlenecek satır.
    :param excel_satir_no: Excel içindeki gerçek satır numarası.
    :returns: ``(Personel | None, SatirReddi | None)``.
    """
    ham_tckn = satir.get(COL_TCKN)
    ham_ad_soyad = satir.get(COL_AD_SOYAD)
    ham_birim = satir.get(COL_BIRIM)

    tckn = normalize_tckn(str(ham_tckn).strip()) if not pd.isna(ham_tckn) else ""
    ad_soyad = "" if pd.isna(ham_ad_soyad) else str(ham_ad_soyad).strip()
    birim = "" if pd.isna(ham_birim) else str(ham_birim).strip()

    hata_nedenleri: list[str] = []
    if not tckn:
        hata_nedenleri.append("TCKN boş")
    elif not validate_tckn(tckn):
        hata_nedenleri.append(f"Geçersiz TCKN: {tckn}")

    if not validate_ad_soyad(ad_soyad):
        hata_nedenleri.append("AD SOYAD boş")

    if hata_nedenleri:
        return None, SatirReddi(
            excel_satir_no=excel_satir_no,
            sebep="; ".join(hata_nedenleri),
            tckn=tckn,
            ad_soyad=ad_soyad,
            birim=birim,
        )

    return Personel(tckn=tckn, ad_soyad=ad_soyad, birim=birim), None
