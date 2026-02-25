"""
D-K (Derece-Kademe) Tablosu kuralları ve hesaplama mantığı.

Bu modül İlk Atama D-K Tablosu PDF'sinden çıkarılan
ünvan, derece ve kademe belirleme kurallarını içerir.

Tablo mantığı:
- Tecrübe yılı = Çalışma alanında toplam prim günü / 360
- Öğrenim durumu: Lisans, Tezsiz YL, Tezli YL, Doktora
- Ünvan/Hizmet Grubu/Kademe bu iki parametreye göre belirlenir
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from src.config.constants import (
    OGRENIM_LISANS,
    OGRENIM_TEZSIZ_YL,
    OGRENIM_TEZLI_YL,
    OGRENIM_DOKTORA,
    GUN_PER_YIL,
)


# ---------------------------------------------------------------------------
# Sonuç veri sınıfı
# ---------------------------------------------------------------------------

@dataclass
class DKSonuc:
    """D-K hesaplama sonucunu tutan veri sınıfı."""

    unvan: str
    """Belirlenen ünvan (ör. 'Uzman Yardımcısı')."""

    hizmet_grubu: str
    """Belirlenen hizmet grubu (ör. 'A/AG-6')."""

    kademe: str
    """Belirlenen kademe veya kademe aralığı (ör. '5' veya '3-4')."""


# ---------------------------------------------------------------------------
# D-K Tablosu
# ---------------------------------------------------------------------------

# Yapı: {hizmet_grubu: [(tecrube_min, tecrube_max, {ogrenim: kademe})]}
# tecrube_min dahil, tecrube_max hariç (son grupta hariç değil)
_DK_TABLOSU: dict[str, list[tuple[float, float, dict[str, str]]]] = {
    "A/AG-6": [
        # 0 - 2 yıl (2 dahil)
        (
            0.0,
            2.0,
            {
                OGRENIM_LISANS: "5-6",
                OGRENIM_TEZSIZ_YL: "5-6",
                OGRENIM_TEZLI_YL: "3",
                OGRENIM_DOKTORA: "-",
            },
        ),
        # 2 - 3 yıl (2 hariç, 3 dahil)
        (
            2.0,
            3.0,
            {
                OGRENIM_LISANS: "3-4",
                OGRENIM_TEZSIZ_YL: "3-4",
                OGRENIM_TEZLI_YL: "2",
                OGRENIM_DOKTORA: "-",
            },
        ),
    ],
    "A/AG-5": [
        # 3 - 5 yıl
        (
            3.0,
            5.0,
            {
                OGRENIM_LISANS: "5",
                OGRENIM_TEZSIZ_YL: "5",
                OGRENIM_TEZLI_YL: "4",
                OGRENIM_DOKTORA: "2",
            },
        ),
        # 6 - 8 yıl
        (
            5.0,
            8.0,
            {
                OGRENIM_LISANS: "3",
                OGRENIM_TEZSIZ_YL: "3",
                OGRENIM_TEZLI_YL: "2",
                OGRENIM_DOKTORA: "2",
            },
        ),
    ],
    "A/AG-4": [
        # 8 - 9 yıl
        (
            8.0,
            9.0,
            {
                OGRENIM_LISANS: "5",
                OGRENIM_TEZSIZ_YL: "5",
                OGRENIM_TEZLI_YL: "4",
                OGRENIM_DOKTORA: "3",
            },
        ),
        # 10 - 12 yıl
        (
            9.0,
            12.0,
            {
                OGRENIM_LISANS: "3",
                OGRENIM_TEZSIZ_YL: "3",
                OGRENIM_TEZLI_YL: "3",
                OGRENIM_DOKTORA: "3",
            },
        ),
    ],
    "A/AG-3": [
        # 12 - 14 yıl
        (
            12.0,
            14.0,
            {
                OGRENIM_LISANS: "5",
                OGRENIM_TEZSIZ_YL: "5",
                OGRENIM_TEZLI_YL: "4",
                OGRENIM_DOKTORA: "2",
            },
        ),
        # 15 - 16 yıl
        (
            14.0,
            16.0,
            {
                OGRENIM_LISANS: "3",
                OGRENIM_TEZSIZ_YL: "3",
                OGRENIM_TEZLI_YL: "2",
                OGRENIM_DOKTORA: "2",
            },
        ),
    ],
    "A/AG-2": [
        # 16+ yıl
        (
            16.0,
            float("inf"),
            {
                OGRENIM_LISANS: "4",
                OGRENIM_TEZSIZ_YL: "3-4",
                OGRENIM_TEZLI_YL: "3",
                OGRENIM_DOKTORA: "3",
            },
        ),
    ],
}

# Ünvan → Hizmet Grubu eşleşmesi
_UNVAN_HIZMET_GRUBU: dict[str, str] = {
    "Uzman Yardımcısı": "A/AG-6",
    "Uzman": "A/AG-5",
    "Kıdemli Uzman": "A/AG-4",
    "Başuzman": "A/AG-3",
    "Kıdemli Başuzman": "A/AG-2",
}

# Hizmet Grubu → tecrübe yılı alt sınırı
_HIZMET_GRUBU_SINIRI: list[tuple[float, str]] = [
    (16.0, "A/AG-2"),
    (12.0, "A/AG-3"),
    (8.0, "A/AG-4"),
    (3.0, "A/AG-5"),
    (0.0, "A/AG-6"),
]


# ---------------------------------------------------------------------------
# Yardımcı fonksiyonlar
# ---------------------------------------------------------------------------

def hesapla_tecrube_yili(alanda_prim_gunu: float) -> float:
    """
    Alanda prim gününü tecrübe yılına çevirir.

    :param alanda_prim_gunu: Çalışma alanında toplam prim günü sayısı.
    :returns: Tecrübe yılı (float).
    """
    return alanda_prim_gunu / GUN_PER_YIL


def belirle_hizmet_grubu(tecrube_yili: float) -> str:
    """
    Tecrübe yılına göre hizmet grubunu belirler.

    :param tecrube_yili: Hesaplanmış tecrübe yılı.
    :returns: Hizmet grubu kodu (ör. 'A/AG-5').
    """
    for sinir, grup in _HIZMET_GRUBU_SINIRI:
        if tecrube_yili >= sinir:
            return grup
    return "A/AG-6"


def belirle_unvan(hizmet_grubu: str) -> str:
    """
    Hizmet grubuna göre ünvanı belirler.

    :param hizmet_grubu: Hizmet grubu kodu.
    :returns: Ünvan adı.
    :raises ValueError: Bilinmeyen hizmet grubu verildiğinde.
    """
    ters_eslesme = {v: k for k, v in _UNVAN_HIZMET_GRUBU.items()}
    if hizmet_grubu not in ters_eslesme:
        raise ValueError(f"Bilinmeyen hizmet grubu: {hizmet_grubu}")
    return ters_eslesme[hizmet_grubu]


def belirle_kademe(
    tecrube_yili: float,
    hizmet_grubu: str,
    ogrenim_durumu: str,
) -> Optional[str]:
    """
    Tecrübe yılı, hizmet grubu ve öğrenim durumuna göre kademeyi belirler.

    :param tecrube_yili: Hesaplanmış tecrübe yılı.
    :param hizmet_grubu: Hizmet grubu kodu (ör. 'A/AG-5').
    :param ogrenim_durumu: Öğrenim seviyesi (ör. 'Lisans').
    :returns: Kademe string'i (ör. '3' veya '3-4'), bulunamazsa None.
    """
    if hizmet_grubu not in _DK_TABLOSU:
        return None

    for min_yil, max_yil, kademe_haritasi in _DK_TABLOSU[hizmet_grubu]:
        # Son aralık için max_yil=inf, o yüzden dahil kontrolü yapılır
        if min_yil <= tecrube_yili < max_yil or (
            max_yil == float("inf") and tecrube_yili >= min_yil
        ):
            kademe = kademe_haritasi.get(ogrenim_durumu)
            return kademe if kademe != "-" else None

    return None


def hesapla_dk(
    alanda_prim_gunu: float,
    ogrenim_durumu: str,
) -> DKSonuc:
    """
    Verilen parametrelerden D-K sonucunu hesaplar.

    :param alanda_prim_gunu: Çalışma alanında toplam prim günü sayısı.
    :param ogrenim_durumu: Öğrenim seviyesi.
    :returns: Hesaplanan D-K sonucu (:class:`DKSonuc`).
    :raises ValueError: Hesaplama mümkün değilse.
    """
    tecrube_yili = hesapla_tecrube_yili(alanda_prim_gunu)
    hizmet_grubu = belirle_hizmet_grubu(tecrube_yili)
    unvan = belirle_unvan(hizmet_grubu)
    kademe = belirle_kademe(tecrube_yili, hizmet_grubu, ogrenim_durumu)

    if kademe is None:
        raise ValueError(
            f"Kademe belirlenemedi: tecrübe={tecrube_yili:.2f} yıl, "
            f"grup={hizmet_grubu}, öğrenim={ogrenim_durumu}"
        )

    return DKSonuc(unvan=unvan, hizmet_grubu=hizmet_grubu, kademe=kademe)
