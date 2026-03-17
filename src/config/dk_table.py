"""
D-K (Derece-Kademe) Tablosu kuralları ve hesaplama mantığı.

Bu modül İlk Atama D-K Tablosu PDF'sinden çıkarılan
ünvan, derece ve kademe belirleme kurallarını içerir.

Tablo mantığı:
- Tecrübe yılı = Çalışma alanında toplam prim günü / 360
- Öğrenim durumu: Lisans, Tezsiz YL, Tezli YL, Doktora
- Ünvan/Seviye/Kademe bu iki parametreye göre belirlenir
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Any

from src.config.constants import (
    OGRENIM_LISANS,
    OGRENIM_TEZSIZ_YL,
    OGRENIM_TEZLI_YL,
    OGRENIM_DOKTORA,
    GUN_PER_YIL,
)

BASLANGIC = "baslangic"
BITIS = "bitis"
KADEMELER = "kademeler"


@dataclass
class DKSonuc:
    """D-K hesaplama sonucunu tutan veri sınıfı."""

    unvan: str
    """Belirlenen ünvan (ör. 'Uzman Yardımcısı')."""

    seviye: str
    """Belirlenen kıdem seviyesi (ör. '6', '5', '4', '3', '2')."""

    kademe: str
    """Belirlenen kademe veya kademe aralığı (ör. '5' veya '3-4')."""


@dataclass
class SeviyeKurali:
    """Bir kıdem seviyesinin tüm kurallarını tek bir yapıda tutar."""

    seviye: str
    unvan: str
    min_tecrube_yili: float
    dk_araliklari: list[dict[str, Any]]


# ---------------------------------------------------------------------------
# D-K Tablosu
# ---------------------------------------------------------------------------

_SEVIYELER: dict[str, SeviyeKurali] = {
    "6": SeviyeKurali(
        seviye="6",
        unvan="Uzman Yardımcısı",
        min_tecrube_yili=0.0,
        dk_araliklari=[
            # 0 - 2 yıl (2 dahil)
            {
                BASLANGIC: 0.0,
                BITIS: 2.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "5", BITIS: "6"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "5", BITIS: "6"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_DOKTORA: {BASLANGIC: "-", BITIS: "-"},
                },
            },
            # 2 - 3 yıl (2 hariç, 3 dahil)
            {
                BASLANGIC: 2.0,
                BITIS: 3.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "3", BITIS: "4"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "3", BITIS: "4"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "2", BITIS: "2"},
                    OGRENIM_DOKTORA: {BASLANGIC: "2", BITIS: "2"},
                },
            },
        ],
    ),
    "5": SeviyeKurali(
        seviye="5",
        unvan="Uzman",
        min_tecrube_yili=3.0,
        dk_araliklari=[
            # 3 - 5 yıl
            {
                BASLANGIC: 3.0,
                BITIS: 5.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "4", BITIS: "4"},
                    OGRENIM_DOKTORA: {BASLANGIC: "2", BITIS: "2"},
                },
            },
            # 6 - 8 yıl
            {
                BASLANGIC: 5.0,
                BITIS: 8.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "2", BITIS: "2"},
                    OGRENIM_DOKTORA: {BASLANGIC: "2", BITIS: "2"},
                },
            },
        ],
    ),
    "4": SeviyeKurali(
        seviye="4",
        unvan="Kıdemli Uzman",
        min_tecrube_yili=8.0,
        dk_araliklari=[
            # 8 - 9 yıl
            {
                BASLANGIC: 8.0,
                BITIS: 9.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "4", BITIS: "4"},
                    OGRENIM_DOKTORA: {BASLANGIC: "3", BITIS: "3"},
                },
            },
            # 10 - 12 yıl
            {
                BASLANGIC: 9.0,
                BITIS: 12.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_DOKTORA: {BASLANGIC: "3", BITIS: "3"},
                },
            },
        ],
    ),
    "3": SeviyeKurali(
        seviye="3",
        unvan="Başuzman",
        min_tecrube_yili=12.0,
        dk_araliklari=[
            # 12 - 14 yıl
            {
                BASLANGIC: 12.0,
                BITIS: 14.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "5", BITIS: "5"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "4", BITIS: "4"},
                    OGRENIM_DOKTORA: {BASLANGIC: "2", BITIS: "2"},
                },
            },
            # 15 - 16 yıl
            {
                BASLANGIC: 14.0,
                BITIS: 16.0,
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "2", BITIS: "2"},
                    OGRENIM_DOKTORA: {BASLANGIC: "2", BITIS: "2"},
                },
            },
        ],
    ),
    "2": SeviyeKurali(
        seviye="2",
        unvan="Kıdemli Başuzman",
        min_tecrube_yili=16.0,
        dk_araliklari=[
            # 16+ yıl
            {
                BASLANGIC: 16.0,
                BITIS: float("inf"),
                KADEMELER: {
                    OGRENIM_LISANS: {BASLANGIC: "4", BITIS: "4"},
                    OGRENIM_TEZSIZ_YL: {BASLANGIC: "3", BITIS: "4"},
                    OGRENIM_TEZLI_YL: {BASLANGIC: "3", BITIS: "3"},
                    OGRENIM_DOKTORA: {BASLANGIC: "3", BITIS: "3"},
                },
            },
        ],
    ),
}


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


def belirle_seviye(tecrube_yili: float) -> str:
    """
    Tecrübe yılına göre kıdem seviyesini belirler.

    :param tecrube_yili: Hesaplanmış tecrübe yılı.
    :returns: Seviye kodu (ör. '5').
    """
    siralı_seviyeler = sorted(
        _SEVIYELER.values(), key=lambda s: s.min_tecrube_yili, reverse=True
    )
    for kural in siralı_seviyeler:
        if tecrube_yili >= kural.min_tecrube_yili:
            return kural.seviye
    return "6"


def belirle_unvan(seviye: str) -> str:
    """
    Seviyeye göre ünvanı belirler.

    :param seviye: Seviye kodu.
    :returns: Ünvan adı.
    :raises ValueError: Bilinmeyen seviye verildiğinde.
    """
    if seviye not in _SEVIYELER:
        raise ValueError(f"Bilinmeyen seviye: {seviye}")
    return _SEVIYELER[seviye].unvan


def belirle_kademe(
    tecrube_yili: float,
    seviye: str,
    ogrenim_durumu: str,
) -> Optional[str]:
    """
    Tecrübe yılı, seviye ve öğrenim durumuna göre kademeyi belirler.

    :param tecrube_yili: Hesaplanmış tecrübe yılı.
    :param seviye: Seviye kodu (ör. '5').
    :param ogrenim_durumu: Öğrenim seviyesi (ör. 'Lisans').
    :returns: Kademe string'i (ör. '3' veya '3-4'), bulunamazsa None.
    """
    if seviye not in _SEVIYELER:
        return None

    kural = _SEVIYELER[seviye]
    for aralik in kural.dk_araliklari:
        min_yil = aralik[BASLANGIC]
        max_yil = aralik[BITIS]
        kademe_haritasi = aralik[KADEMELER]

        # Son aralık için max_yil=inf, o yüzden dahil kontrolü yapılır
        if min_yil <= tecrube_yili < max_yil or (
            max_yil == float("inf") and tecrube_yili >= min_yil
        ):
            kademe_dict = kademe_haritasi.get(ogrenim_durumu)
            if kademe_dict and kademe_dict.get(BASLANGIC) != "-":
                bas = kademe_dict[BASLANGIC]
                bit = kademe_dict[BITIS]
                return f"{bas}-{bit}" if bas != bit else str(bas)
            return None

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
    seviye = belirle_seviye(tecrube_yili)
    unvan = belirle_unvan(seviye)
    kademe = belirle_kademe(tecrube_yili, seviye, ogrenim_durumu)

    if kademe is None:
        raise ValueError(
            f"Kademe belirlenemedi: tecrübe={tecrube_yili:.2f} yıl, "
            f"seviye={seviye}, öğrenim={ogrenim_durumu}"
        )

    return DKSonuc(unvan=unvan, seviye=seviye, kademe=kademe)
