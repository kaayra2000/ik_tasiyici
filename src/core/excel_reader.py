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
from src.core.validators import (
    normalize_tckn,
    validate_tckn,
    validate_ad_soyad,
    validate_birim,
)


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
    dosya_yolu = Path(dosya_yolu)
    if not dosya_yolu.exists():
        raise FileNotFoundError(f"Kaynak dosya bulunamadı: {dosya_yolu}")

    df = pd.read_excel(dosya_yolu, dtype=str)

    _zorunlu_sutunları_dogrula(df)

    personeller: List[Personel] = []
    for _, satir in df.iterrows():
        personel = _satiri_isle(satir)
        if personel is not None:
            personeller.append(personel)

    return personeller


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


def _satiri_isle(satir: pd.Series) -> Personel | None:
    """
    Tek bir DataFrame satırını işler ve geçerliyse :class:`Personel` döner.

    Geçersiz ya da eksik veri içeren satırlar için ``None`` döner.

    :param satir: İşlenecek satır.
    :returns: :class:`Personel` veya ``None``.
    """
    ham_tckn = satir.get(COL_TCKN)
    ham_ad_soyad = satir.get(COL_AD_SOYAD)
    ham_birim = satir.get(COL_BIRIM)

    # Pandas NaN veya None kontrolü
    if pd.isna(ham_tckn) or pd.isna(ham_ad_soyad) or pd.isna(ham_birim):
        return None

    tckn = normalize_tckn(str(ham_tckn).strip())
    ad_soyad = str(ham_ad_soyad).strip()
    birim = str(ham_birim).strip()

    if not validate_tckn(tckn):
        return None
    if not validate_ad_soyad(ad_soyad):
        return None
    if not validate_birim(birim):
        return None

    return Personel(tckn=tckn, ad_soyad=ad_soyad, birim=birim)
