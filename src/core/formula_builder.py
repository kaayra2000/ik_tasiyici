"""
Excel formülleri üretici modülü.

Tüm fonksiyonlar Türkçe Excel formül adlarını kullanır
(kullanıcının varsayılan dil ayarına göre: ``EĞER``, ``TOPLA``, ``VE``).

Formül mantığı AGENT.md §Excel Formül Gereksinimleri bölümünden alınmıştır.
"""

from __future__ import annotations

from src.config.constants import (
    COL_BASLANGIC_TARIHI,
    COL_BITIS_TARIHI,
    COL_ALANINDA,
    COL_TOPLAM_PRIM,
    COL_ALANDA_PRIM,
    GUN_PER_YIL,
    TECRUBE_BASLANGIC_SATIR,
    TECRUBE_BITIS_SATIR,
)


# ---------------------------------------------------------------------------
# Prim günü formülleri (satır bazlı)
# ---------------------------------------------------------------------------


def prim_gunu_formulu(satir: int) -> str:
    """
    Belirtilen satır için toplam prim günü formülünü üretir.

    Başlangıç ve bitiş tarihi doluysa gün farkını hesaplar, aksi hâlde boş döner.

    :param satir: Hedef Excel satır numarası (1-indexed).
    :returns: Türkçe Excel formülü string'i.

    Örnek::

        prim_gunu_formulu(10)
        # =EĞER(VE(D10<>"";E10<>"");E10-D10;"")
    """
    d = f"{COL_BASLANGIC_TARIHI}{satir}"
    e = f"{COL_BITIS_TARIHI}{satir}"
    return f'=EĞER(VE({d}<>"";{e}<>"");{e}-{d};"")'


def alanda_prim_formulu(satir: int) -> str:
    """
    Belirtilen satır için çalışma alanında prim günü formülünü üretir.

    "Alanında" sütunu "E" ise o satırın toplam prim günü değerini alır.

    :param satir: Hedef Excel satır numarası (1-indexed).
    :returns: Türkçe Excel formülü string'i.

    Örnek::

        alanda_prim_formulu(10)
        # =EĞER(J10="E";K10;"")
    """
    j = f"{COL_ALANINDA}{satir}"
    k = f"{COL_TOPLAM_PRIM}{satir}"
    return f'=EĞER({j}="E";{k};"")'


# ---------------------------------------------------------------------------
# Toplam formülleri
# ---------------------------------------------------------------------------


def toplam_prim_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """
    Toplam prim günü aralık toplamı formülünü üretir.

    :param bitis_satir: Son tecrübe satırının numarası.
    :param baslangic_satir: İlk tecrübe satırının numarası.
    :returns: ``=TOPLA(K10:K{n})`` biçiminde formül.
    """
    return f"=TOPLA({COL_TOPLAM_PRIM}{baslangic_satir}:{COL_TOPLAM_PRIM}{bitis_satir})"


def toplam_alanda_prim_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """
    Çalışma alanında toplam prim günü aralık toplamı formülünü üretir.

    :param bitis_satir: Son tecrübe satırının numarası.
    :param baslangic_satir: İlk tecrübe satırının numarası.
    :returns: ``=TOPLA(L10:L{n})`` biçiminde formül.
    """
    return f"=TOPLA({COL_ALANDA_PRIM}{baslangic_satir}:{COL_ALANDA_PRIM}{bitis_satir})"


# ---------------------------------------------------------------------------
# Tecrübe yılı formülü
# ---------------------------------------------------------------------------


def tecrube_yili_formulu(alanda_toplam_hucre: str) -> str:
    """
    Tecrübe yılını hesaplayan formülü üretir.

    Alanda toplam prim gününü ``GUN_PER_YIL`` (360) e böler.

    :param alanda_toplam_hucre: Alanda toplam prim gününün bulunduğu hücre adresi
        (ör. ``"L27"``).
    :returns: Bölme formülü string'i.

    Örnek::

        tecrube_yili_formulu("L27")
        # =L27/360
    """
    return f"={alanda_toplam_hucre}/{GUN_PER_YIL}"


# ---------------------------------------------------------------------------
# En yüksek öğrenim belirleme formülü
# ---------------------------------------------------------------------------


def en_yuksek_ogrenim_formulu(
    doktora_hucre: str,
    doktora_alaninda_hucre: str,
    tezli_yl_hucre: str,
    tezli_yl_alaninda_hucre: str,
    tezsiz_yl_hucre: str,
    tezsiz_yl_alaninda_hucre: str,
    lisans_hucre: str,
    lisans_alaninda_hucre: str,
) -> str:
    """
    Alanında okunan en yüksek öğrenim seviyesini dönen formülü üretir.

    her öğrenim seviyesi için hem öğrenim hücresi dolu, hem de "alanında" E olmalıdır.

    :returns: İç içe EĞER formülü string'i.
    """
    return (
        f'=EĞER(VE({doktora_hucre}<>"";{doktora_alaninda_hucre}="E");"Doktora";'
        f'EĞER(VE({tezli_yl_hucre}<>"";{tezli_yl_alaninda_hucre}="E");"Tezli YL";'
        f'EĞER(VE({tezsiz_yl_hucre}<>"";{tezsiz_yl_alaninda_hucre}="E");"Tezsiz YL";'
        f'EĞER(VE({lisans_hucre}<>"";{lisans_alaninda_hucre}="E");"Lisans";""))))'
    )


# ---------------------------------------------------------------------------
# Ünvan / Hizmet Grubu / Kademe formülleri
# ---------------------------------------------------------------------------


def unvan_formulu(tecrube_yili_hucre: str) -> str:
    """
    Tecrübe yılına göre ünvanı belirleyen formülü üretir.

    Eşik değerleri AGENT.md §Ünvan/Derece/Kademe Hesaplama bölümünden alınmıştır.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre (ör. ``"N27"``).
    :returns: İç içe EĞER formülü string'i.
    """
    t = tecrube_yili_hucre
    return (
        f'=EĞER({t}>=16;"Kıdemli Başuzman";'
        f'EĞER({t}>=12;"Başuzman";'
        f'EĞER({t}>=8;"Kıdemli Uzman";'
        f'EĞER({t}>=3;"Uzman";"Uzman Yardımcısı"))))'
    )


def hizmet_grubu_formulu(tecrube_yili_hucre: str) -> str:
    """
    Tecrübe yılına göre hizmet grubunu belirleyen formülü üretir.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre.
    :returns: İç içe EĞER formülü string'i.
    """
    t = tecrube_yili_hucre
    return (
        f'=EĞER({t}>=16;"A/AG-2";'
        f'EĞER({t}>=12;"A/AG-3";'
        f'EĞER({t}>=8;"A/AG-4";'
        f'EĞER({t}>=3;"A/AG-5";"A/AG-6"))))'
    )


def kademe_formulu(
    tecrube_yili_hucre: str,
    ogrenim_hucre: str,
) -> str:
    """
    Tecrübe yılı ve öğrenim durumuna göre kademeyi belirleyen formülü üretir.

    D-K Tablosundaki matrise göre oluşturulmuştur (AGENT.md §Kademe Belirleme Matrisi).
    Formül önce hizmet grubunu (tecrübe aralığı), ardından öğrenim durumunu kontrol eder.

    :param tecrube_yili_hucre: Tecrübe yılı hücresi (ör. ``"N27"``).
    :param ogrenim_hucre: Öğrenim durumu hücresi (ör. ``"B5"``).
    :returns: İç içe EĞER formülü string'i.
    """
    t = tecrube_yili_hucre
    o = ogrenim_hucre

    # A/AG-2 (16+)
    ag2 = (
        f'EĞER({o}="Lisans";"4";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"3-4";'
        f'EĞER({o}="Tezli Yüksek Lisans";"3";"3")))'
    )

    # A/AG-3 (12-16)
    ag3_low = (
        f'EĞER({o}="Lisans";"5";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"5";'
        f'EĞER({o}="Tezli Yüksek Lisans";"4";"2")))'
    )
    ag3_high = (
        f'EĞER({o}="Lisans";"3";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"3";'
        f'EĞER({o}="Tezli Yüksek Lisans";"2";"2")))'
    )
    ag3 = f'EĞER({t}<14;{ag3_low};{ag3_high})'

    # A/AG-4 (8-12)
    ag4_low = (
        f'EĞER({o}="Lisans";"5";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"5";'
        f'EĞER({o}="Tezli Yüksek Lisans";"4";"3")))'
    )
    ag4_high = (
        f'EĞER({o}="Lisans";"3";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"3";'
        f'EĞER({o}="Tezli Yüksek Lisans";"3";"3")))'
    )
    ag4 = f'EĞER({t}<9;{ag4_low};{ag4_high})'

    # A/AG-5 (3-8)
    ag5_low = (
        f'EĞER({o}="Lisans";"5";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"5";'
        f'EĞER({o}="Tezli Yüksek Lisans";"4";"2")))'
    )
    ag5_high = (
        f'EĞER({o}="Lisans";"3";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"3";'
        f'EĞER({o}="Tezli Yüksek Lisans";"2";"2")))'
    )
    ag5 = f'EĞER({t}<5;{ag5_low};{ag5_high})'

    # A/AG-6 (0-3)
    ag6_low = (
        f'EĞER({o}="Lisans";"5-6";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"5-6";'
        f'EĞER({o}="Tezli Yüksek Lisans";"3";"")))'
    )
    ag6_high = (
        f'EĞER({o}="Lisans";"3-4";'
        f'EĞER({o}="Tezsiz Yüksek Lisans";"3-4";'
        f'EĞER({o}="Tezli Yüksek Lisans";"2";"")))'
    )
    ag6 = f'EĞER({t}<2;{ag6_low};{ag6_high})'

    return (
        f'=EĞER({t}>=16;{ag2};'
        f'EĞER({t}>=12;{ag3};'
        f'EĞER({t}>=8;{ag4};'
        f'EĞER({t}>=3;{ag5};{ag6}))))'
    )
