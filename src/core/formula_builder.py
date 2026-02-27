"""
Excel formülleri üretici modülü.

openpyxl / OOXML standardı İngilizce fonksiyon adları kullanır.
LibreOffice ve Excel, dosyayı açarken formülleri otomatik olarak
kullanıcının arayüz diline çevirir.

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
    :returns: OOXML uyumlu İngilizce Excel formülü string'i.

    Örnek::

        prim_gunu_formulu(11)
        # =IF(AND(E11<>"",F11<>""),F11-E11,"")
    """
    d = f"{COL_BASLANGIC_TARIHI}{satir}"
    e = f"{COL_BITIS_TARIHI}{satir}"
    return f'=IF(AND({d}<>"",{e}<>""),{e}-{d},"")'


def alanda_prim_formulu(satir: int) -> str:
    """
    Belirtilen satır için çalışma alanında prim günü formülünü üretir.

    "Alanında" sütunu "E" ise o satırın toplam prim günü değerini alır.

    :param satir: Hedef Excel satır numarası (1-indexed).
    :returns: OOXML uyumlu İngilizce Excel formülü string'i.

    Örnek::

        alanda_prim_formulu(11)
        # =IF(J11="E",K11,"")
    """
    j = f"{COL_ALANINDA}{satir}"
    k = f"{COL_TOPLAM_PRIM}{satir}"
    return f'=IF({j}="E",{k},"")'


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
    :returns: ``=SUM(K11:K{n})`` biçiminde formül.
    """
    return f"=SUM({COL_TOPLAM_PRIM}{baslangic_satir}:{COL_TOPLAM_PRIM}{bitis_satir})"


def toplam_alanda_prim_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """
    Çalışma alanında toplam prim günü aralık toplamı formülünü üretir.

    :param bitis_satir: Son tecrübe satırının numarası.
    :param baslangic_satir: İlk tecrübe satırının numarası.
    :returns: ``=SUM(L11:L{n})`` biçiminde formül.
    """
    return f"=SUM({COL_ALANDA_PRIM}{baslangic_satir}:{COL_ALANDA_PRIM}{bitis_satir})"


# ---------------------------------------------------------------------------
# Tecrübe yılı formülü
# ---------------------------------------------------------------------------


def tecrube_yili_formulu(alanda_toplam_hucre: str) -> str:
    """
    Tecrübe yılını hesaplayan formülü üretir.

    Alanda toplam prim gününü ``GUN_PER_YIL`` (360) e böler.

    :param alanda_toplam_hucre: Alanda toplam prim gününün bulunduğu hücre adresi
        (ör. ``"L19"``).
    :returns: Bölme formülü string'i.

    Örnek::

        tecrube_yili_formulu("L19")
        # =L19/360
    """
    return f"={alanda_toplam_hucre}/{GUN_PER_YIL}"


# ---------------------------------------------------------------------------
# En yüksek öğrenim belirleme formülü
# ---------------------------------------------------------------------------


def en_yuksek_ogrenim_formulu(
    baslangic_satir: int = 6,
    bitis_satir: int = 8,
    ad_sutun: str = "B",
    okul_sutun: str = "C",
    alaninda_sutun: str = "K",
) -> str:
    """
    Belirtilen aralıkta alanında okunan en yüksek öğrenim seviyesini (adını) döndüren Excel formülünü üretir.

    Daha yüksek satır numarasının (örneğin 8 - Doktora), daha düşük satır numarasına (örneğin 6 - Lisans)
    göre daha üst bir öğrenim seviyesini ifade ettiği varsayılır. Bu nedenle en yüksek satırdan aşağıya (öncelikli)
    doğru kontrol eden bir formül oluşturulur.
    Her öğrenim seviyesi için okul hücresi dolu olmalı ve "alanında" hücresi "E" olmalıdır.

    :param baslangic_satir: Öğrenim bilgilerinin başladığı satır (ör. ``6``).
    :param bitis_satir: Öğrenim bilgilerinin bittiği satır (ör. ``8``).
    :param ad_sutun: Öğrenim adının (ör. Lisans, Doktora) bulunduğu sütun harfi (ör. ``"B"``).
    :param okul_sutun: Okul adının bulunduğu sütun harfi (ör. ``"C"``).
    :param alaninda_sutun: Alanında olup olmadığını ("E"/"H") belirten sütun harfi (ör. ``"K"``).
    :returns: İç içe IF formülü string'i.
    """
    formul = '""'
    for satir in range(baslangic_satir, bitis_satir + 1):
        okul = f"{okul_sutun}{satir}"
        alaninda = f"{alaninda_sutun}{satir}"
        ad = f"{ad_sutun}{satir}"
        formul = f'IF(AND({okul}<>"",{alaninda}="E"),{ad},{formul})'

    return f"={formul}"


# ---------------------------------------------------------------------------
# Ünvan / Hizmet Grubu / Kademe formülleri
# ---------------------------------------------------------------------------


def unvan_formulu(tecrube_yili_hucre: str) -> str:
    """
    Tecrübe yılına göre ünvanı belirleyen formülü üretir.

    Eşik değerleri AGENT.md §Ünvan/Derece/Kademe Hesaplama bölümünden alınmıştır.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre (ör. ``"Z1"``).
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    return (
        f'=IF({t}>=16,"Kıdemli Başuzman",'
        f'IF({t}>=12,"Başuzman",'
        f'IF({t}>=8,"Kıdemli Uzman",'
        f'IF({t}>=3,"Uzman","Uzman Yardımcısı"))))'
    )


def hizmet_grubu_formulu(tecrube_yili_hucre: str) -> str:
    """
    Tecrübe yılına göre hizmet grubunu belirleyen formülü üretir.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre.
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    return (
        f'=IF({t}>=16,"A/AG-2",'
        f'IF({t}>=12,"A/AG-3",'
        f'IF({t}>=8,"A/AG-4",'
        f'IF({t}>=3,"A/AG-5","A/AG-6"))))'
    )


def kademe_formulu(
    tecrube_yili_hucre: str,
    ogrenim_hucre: str,
) -> str:
    """
    Tecrübe yılı ve öğrenim durumuna göre kademeyi belirleyen formülü üretir.

    D-K Tablosundaki matrise göre oluşturulmuştur (AGENT.md §Kademe Belirleme Matrisi).
    Formül önce hizmet grubunu (tecrübe aralığı), ardından öğrenim durumunu kontrol eder.

    :param tecrube_yili_hucre: Tecrübe yılı hücresi (ör. ``"Z1"``).
    :param ogrenim_hucre: Öğrenim durumu hücresi (ör. ``"C8"``).
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    o = ogrenim_hucre

    def branch(lisans: str, tezsiz: str, tezli: str, doktora: str) -> str:
        return (
            f'IF({o}="Lisans","{lisans}",'
            f'IF({o}="Tezsiz Yüksek Lisans","{tezsiz}",'
            f'IF({o}="Tezli Yüksek Lisans","{tezli}","{doktora}")))'
        )

    # A/AG-2 (16+)
    ag2 = branch("3", "3", "2", "2")

    # A/AG-3 (12-16)
    ag3_low = branch("5", "5", "4", "2")
    ag3_high = branch("3", "3", "2", "2")
    ag3 = f'IF({t}<15,{ag3_low},{ag3_high})'

    # A/AG-4 (8-12)
    ag4_low = branch("5", "5", "4", "2")
    ag4_high = branch("3", "3", "2", "2")
    ag4 = f'IF({t}<10,{ag4_low},{ag4_high})'

    # A/AG-5 (3-8)
    ag5_low = branch("5", "5", "4", "2")
    ag5_high = branch("3", "3", "2", "2")
    ag5 = f'IF({t}<6,{ag5_low},{ag5_high})'

    # A/AG-6 (0-3)
    ag6_low = branch("5", "4", "3", "2")
    ag6_high = branch("3", "3", "2", "2")
    ag6 = f'IF({t}<2,{ag6_low},{ag6_high})'

    return (
        f'=IF({t}>=16,{ag2},'
        f'IF({t}>=12,{ag3},'
        f'IF({t}>=8,{ag4},'
        f'IF({t}>=3,{ag5},{ag6}))))'
    )
