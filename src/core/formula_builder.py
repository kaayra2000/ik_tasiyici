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
    COL_EKSIK_GUN,
    OGRENIM_DOKTORA,
    OGRENIM_LISANS,
    OGRENIM_TEZLI_YL,
    OGRENIM_TEZSIZ_YL,
    TECRUBE_BASLANGIC_SATIR,
    TECRUBE_BITIS_SATIR,
)

# ---------------------------------------------------------------------------
# L28 toplam gün bazlı tecrübe yıl/ay/gün formülleri (takvim bazlı)
# ---------------------------------------------------------------------------


def _days360_toplam_ifadesi(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """
    Belirtilen satır aralığındaki DAYS360 toplamı için formül ifadesi üretir.

    Boş tarih hücrelerini hesaba katmamak için SUMPRODUCT ile filtre uygular.
    """
    baslangic_aralik = (
        f"{COL_BASLANGIC_TARIHI}{baslangic_satir}:{COL_BASLANGIC_TARIHI}{bitis_satir}"
    )
    bitis_aralik = (
        f"{COL_BITIS_TARIHI}{baslangic_satir}:{COL_BITIS_TARIHI}{bitis_satir}"
    )
    alaninda_aralik = f"{COL_ALANINDA}{baslangic_satir}:{COL_ALANINDA}{bitis_satir}"
    return (
        f'SUMPRODUCT(--({baslangic_aralik}<>""),--({bitis_aralik}<>""),'
        f'--({alaninda_aralik}="E"),'
        f"DAYS360({baslangic_aralik},{bitis_aralik},0))"
    )


def tecrube_360_yil_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """L28 toplam gün hücresinden takvim bazlı yıl değerini döndüren formülü üretir."""
    _ = (bitis_satir, baslangic_satir)
    toplam_hucre = "L28"
    baslangic_tarih = "DATE(2001,1,1)"
    return f'=IF({toplam_hucre}=0,"",' f"YEAR({baslangic_tarih}+{toplam_hucre})-2001)"


def tecrube_360_ay_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """L28 toplam gün hücresinden takvim bazlı ay değerini döndüren formülü üretir."""
    _ = (bitis_satir, baslangic_satir)
    toplam_hucre = "L28"
    baslangic_tarih = "DATE(2001,1,1)"
    return f'=IF({toplam_hucre}=0,"",' f"MONTH({baslangic_tarih}+{toplam_hucre})-1)"


def tecrube_360_gun_formulu(
    bitis_satir: int = TECRUBE_BITIS_SATIR,
    baslangic_satir: int = TECRUBE_BASLANGIC_SATIR,
) -> str:
    """L28 toplam gün hücresinden takvim bazlı gün değerini döndüren formülü üretir."""
    _ = (bitis_satir, baslangic_satir)
    toplam_hucre = "L28"
    baslangic_tarih = "DATE(2001,1,1)"
    return f'=IF({toplam_hucre}=0,"",DAY({baslangic_tarih}+{toplam_hucre})-1)'


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
    :returns: ``=SUM(L11:L{n})-SUM(M11:M{n})`` biçiminde formül.
    """
    alanda = f"{COL_ALANDA_PRIM}{baslangic_satir}:{COL_ALANDA_PRIM}{bitis_satir}"
    eksik = f"{COL_EKSIK_GUN}{baslangic_satir}:{COL_EKSIK_GUN}{bitis_satir}"
    return f"=SUM({alanda})-SUM({eksik})"


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

    Sıralama satır numarasına değil, öğrenim seviyesine göre yapılır.
    Öncelik: Doktora > Tezli Yüksek Lisans > Tezsiz Yüksek Lisans > Lisans.
    Her öğrenim seviyesi için okul hücresi dolu olmalı ve "alanında" hücresi "E" olmalıdır.
    Ayrıca ilgili satırın öğrenim adı hücresinin (B sütunu) hedef seviye ile eşleşmesi gerekir.

    :param baslangic_satir: Öğrenim bilgilerinin başladığı satır (ör. ``6``).
    :param bitis_satir: Öğrenim bilgilerinin bittiği satır (ör. ``8``).
    :param ad_sutun: Öğrenim adının (ör. Lisans, Doktora) bulunduğu sütun harfi (ör. ``"B"``).
    :param okul_sutun: Okul adının bulunduğu sütun harfi (ör. ``"C"``).
    :param alaninda_sutun: Alanında olup olmadığını ("E"/"H") belirten sütun harfi (ör. ``"K"``).
    :returns: İç içe IF formülü string'i.
    """

    def seviye_var_mi(seviye: str) -> str:
        kosullar = []
        for satir in range(baslangic_satir, bitis_satir + 1):
            okul = f"{okul_sutun}{satir}"
            alaninda = f"{alaninda_sutun}{satir}"
            ad = f"{ad_sutun}{satir}"
            kosullar.append(f'AND({ad}="{seviye}",{okul}<>"",{alaninda}="E")')
        if not kosullar:
            return "FALSE"
        return f'OR({",".join(kosullar)})'

    return (
        f'=IF({seviye_var_mi(OGRENIM_DOKTORA)},"{OGRENIM_DOKTORA}",'
        f'IF({seviye_var_mi(OGRENIM_TEZLI_YL)},"{OGRENIM_TEZLI_YL}",'
        f'IF({seviye_var_mi(OGRENIM_TEZSIZ_YL)},"{OGRENIM_TEZSIZ_YL}",'
        f'IF({seviye_var_mi(OGRENIM_LISANS)},"{OGRENIM_LISANS}",""))))'
    )


# ---------------------------------------------------------------------------
# Ünvan / Hizmet Grubu / Kademe formülleri
# ---------------------------------------------------------------------------


BRUT_UCRET_HARITASI: dict[str, float] = {
    # AG Hizmet Grubu
    "AG-1/6": 347991.73,
    "AG-1/5": 351432.73,
    "AG-1/4": 354908.10,
    "AG-1/3": 358418.23,
    "AG-1/2": 361963.45,
    "AG-1/1": 365544.19,
    "AG-2/6": 281908.77,
    "AG-2/5": 298589.66,
    "AG-2/4": 316271.37,
    "AG-2/3": 333452.13,
    "AG-2/2": 339054.57,
    "AG-2/1": 344584.72,
    "AG-3/6": 239949.31,
    "AG-3/5": 251751.97,
    "AG-3/4": 262905.55,
    "AG-3/3": 273265.98,
    "AG-3/2": 277306.69,
    "AG-3/1": 281407.93,
    "AG-4/6": 193332.48,
    "AG-4/5": 209434.76,
    "AG-4/4": 225877.97,
    "AG-4/3": 232537.52,
    "AG-4/2": 235967.18,
    "AG-4/1": 239448.36,
    "AG-5/6": 154169.78,
    "AG-5/5": 167694.60,
    "AG-5/4": 177522.64,
    "AG-5/3": 187288.41,
    "AG-5/2": 190039.34,
    "AG-5/1": 192831.54,
    "AG-6/6": 104077.85,
    "AG-6/5": 116099.83,
    "AG-6/4": 129564.61,
    "AG-6/3": 142131.67,
    "AG-6/2": 154573.14,
    "AG-6/1": 168134.22,
    # A Hizmet Grubu
    "A-1/6": 227298.27,
    "A-1/5": 229512.78,
    "A-1/4": 231749.78,
    "A-1/3": 234009.42,
    "A-1/2": 236291.86,
    "A-1/1": 238597.53,
    "A-2/6": 197970.47,
    "A-2/5": 204595.56,
    "A-2/4": 211452.58,
    "A-2/3": 218549.57,
    "A-2/2": 221907.62,
    "A-2/1": 225105.92,
    "A-3/6": 178020.46,
    "A-3/5": 183100.72,
    "A-3/4": 187461.22,
    "A-3/3": 191930.70,
    "A-3/2": 194679.48,
    "A-3/1": 197469.52,
    "A-4/6": 149625.05,
    "A-4/5": 158786.45,
    "A-4/4": 167792.84,
    "A-4/3": 172566.13,
    "A-4/2": 175024.46,
    "A-4/1": 177519.55,
    "A-5/6": 128901.11,
    "A-5/5": 134912.16,
    "A-5/4": 141223.71,
    "A-5/3": 146356.78,
    "A-5/2": 147733.52,
    "A-5/1": 149124.08,
    "A-6/6": 98845.86,
    "A-6/5": 106960.89,
    "A-6/4": 115806.08,
    "A-6/3": 122233.57,
    "A-6/2": 129046.76,
    "A-6/1": 135666.93,
}


def brut_ucret_formulu(kademe_hucre: str, tablo_araligi: str | None = None) -> str:
    """
    F3 benzeri derece/kademe anahtarına göre G3 brüt ücret formülünü üretir.

    Anahtar formatı: ``AG-2/4`` veya ``A-5/1``.
    Eşleşme bulunamazsa boş string döner.
    """
    if not BRUT_UCRET_HARITASI:
        return '=""'

    if not tablo_araligi:
        raise ValueError("tablo_araligi zorunludur")

    # SWITCH eski ofis surumlerinde #NAME? verebildigi icin VLOOKUP kullanilir.
    return f'=IFERROR(VLOOKUP({kademe_hucre},{tablo_araligi},2,FALSE),"")'


def unvan_formulu(
    tecrube_yili_hucre: str,
    hizmet_grubu_turu_hucre: str,
) -> str:
    """
    Tecrübe yılına ve hizmet grubu türüne göre ünvanı belirleyen formülü üretir.

    Eşik değerleri AGENT.md §Ünvan/Derece/Kademe Hesaplama bölümünden alınmıştır.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre (ör. ``"J29"``).
    :param hizmet_grubu_turu_hucre: Hizmet grubu türü seçiminin bulunduğu hücre
        (ör. ``"M3"``). ``"AG"`` ise 3+ yıl ünvanlarına ``" Araştırmacı"``
        eklenir; 3 yıl altı için yalnızca ``"Araştırmacı"`` döner.
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    g = hizmet_grubu_turu_hucre

    def varyant(unvan: str) -> str:
        return f'IF({g}="AG","{unvan} Araştırmacı","{unvan}")'

    return (
        f'=IF({t}>=16,{varyant("Kıdemli Başuzman")},'
        f'IF({t}>=12,{varyant("Başuzman")},'
        f'IF({t}>=8,{varyant("Kıdemli Uzman")},'
        f'IF({t}>=3,{varyant("Uzman")},IF({g}="AG","Araştırmacı","Uzman Yardımcısı")))))'
    )


def hizmet_grubu_formulu(
    tecrube_yili_hucre: str,
    hizmet_grubu_turu_hucre: str,
) -> str:
    """
    Tecrübe yılına ve seçilen türe göre hizmet grubunu belirleyen formülü üretir.

    :param tecrube_yili_hucre: Tecrübe yılı değerinin bulunduğu hücre.
    :param hizmet_grubu_turu_hucre: Hizmet grubu türü seçiminin bulunduğu hücre.
        Beklenen değerler ``"A"`` veya ``"AG"``.
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    g = hizmet_grubu_turu_hucre
    return (
        f'=IF(OR({g}="A",{g}="AG"),'
        f'IF({t}>=16,{g}&"-2",'
        f'IF({t}>=12,{g}&"-3",'
        f'IF({t}>=8,{g}&"-4",'
        f'IF({t}>=3,{g}&"-5",{g}&"-6")))),'
        f'""'
        f")"
    )


def kademe_formulu(
    tecrube_yili_hucre: str,
    ogrenim_hucre: str,
) -> str:
    """
    Tecrübe yılı ve öğrenim durumuna göre kademeyi belirleyen formülü üretir.

    D-K Tablosundaki matrise göre oluşturulmuştur (AGENT.md §Kademe Belirleme Matrisi).
    Formül önce hizmet grubunu (tecrübe aralığı), ardından öğrenim durumunu kontrol eder.

    :param tecrube_yili_hucre: Tecrübe yılı hücresi (ör. ``"J29"``).
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
    ag3 = f"IF({t}<15,{ag3_low},{ag3_high})"

    # A/AG-4 (8-12)
    ag4_low = branch("5", "5", "4", "2")
    ag4_high = branch("3", "3", "2", "2")
    ag4 = f"IF({t}<10,{ag4_low},{ag4_high})"

    # A/AG-5 (3-8)
    ag5_low = branch("5", "5", "4", "2")
    ag5_high = branch("3", "3", "2", "2")
    ag5 = f"IF({t}<6,{ag5_low},{ag5_high})"

    # A/AG-6 (0-3)
    ag6_low = branch("5", "4", "3", "2")
    ag6_high = branch("3", "3", "2", "2")
    ag6 = f"IF({t}<2,{ag6_low},{ag6_high})"

    return (
        f"=IF({t}>=16,{ag2},"
        f"IF({t}>=12,{ag3},"
        f"IF({t}>=8,{ag4},"
        f"IF({t}>=3,{ag5},{ag6}))))"
    )


def kademe_baslangic_formulu(
    tecrube_yili_hucre: str,
    ogrenim_hucre: str,
    derece_kademe_hucre: str | None = None,
) -> str:
    """
    Tecrübe yılı ve öğrenim durumuna göre kademe başlangıcını belirleyen formülü üretir.
    Kıdem tablosuna göre kademe başlangıcı değerleri tecrübe yılı ve eğitim seviyesine göre değişir.
    NOT: Kademe AZALDIKÇA yükselme olur (6'dan 5'e, 5'ten 4'e).

    :param tecrube_yili_hucre: Tecrübe yılı hücresi (ör. ``"J29"``).
    :param ogrenim_hucre: Öğrenim durumu hücresi (ör. ``"Z4"``).
    :param derece_kademe_hucre: Derece/kademe hücresi (ör. ``"F3"``).
        Verilirse formül, ``AG-4/`` benzeri ön eki bu hücreden alır ve
        sayısal kademeyle birleştirir.
    :returns: Kademe başlangıcı formülü string'i.
    """
    t = tecrube_yili_hucre
    o = ogrenim_hucre

    def branch(lisans: str, tezsiz: str, tezli: str, doktora: str) -> str:
        return (
            f'IF({o}="Lisans",{lisans},'
            f'IF({o}="Tezsiz Yüksek Lisans",{tezsiz},'
            f'IF({o}="Tezli Yüksek Lisans",{tezli},{doktora})))'
        )

    # 16+ yıl
    ag2 = branch("6", "6", "5", "4")

    # 15-16 yıl
    ag3_15_16 = branch("4", "4", "3", "4")

    # 12-15 yıl
    ag3_12_15 = branch("6", "6", "5", "4")

    # 10-12 yıl
    ag4_10_12 = branch("4", "4", "3", "4")

    # 8-10 yıl
    ag4_8_10 = branch("6", "6", "5", "4")

    # 6-8 yıl
    ag5_6_8 = branch("4", "4", "3", "4")

    # 3-6 yıl
    ag5_3_6 = branch("6", "6", "5", "4")

    # 2-3 yıl
    ag6_2_3 = branch("4", "4", "2", "2")

    # 0-2 yıl
    ag6_0_2 = branch("6", "5", "4", "4")

    kademe_ifadesi = (
        f"IF({t}>=16,{ag2},"
        f"IF({t}>=15,{ag3_15_16},"
        f"IF({t}>=12,{ag3_12_15},"
        f"IF({t}>=10,{ag4_10_12},"
        f"IF({t}>=8,{ag4_8_10},"
        f"IF({t}>=6,{ag5_6_8},"
        f"IF({t}>=3,{ag5_3_6},"
        f"IF({t}>=2,{ag6_2_3},{ag6_0_2}))))))))"
    )
    if not derece_kademe_hucre:
        return f"={kademe_ifadesi}"

    dk = derece_kademe_hucre
    on_ek = f'IFERROR(LEFT({dk},FIND("/",{dk})),"")'
    return f'=IF({dk}="",{kademe_ifadesi},{on_ek}&{kademe_ifadesi})'


def kademe_bitis_formulu(
    tecrube_yili_hucre: str,
    ogrenim_hucre: str,
    derece_kademe_hucre: str | None = None,
) -> str:
    """
    Tecrübe yılı ve öğrenim durumuna göre kademe bitişini belirleyen formülü üretir.
    Kıdem tablosundaki değerlere göre oluşturulmuştur.

    :param tecrube_yili_hucre: Tecrübe yılı hücresi (ör. ``"J29"``).
    :param ogrenim_hucre: Öğrenim durumu hücresi (ör. ``"Z4"``).
    :param derece_kademe_hucre: Derece/kademe hücresi (ör. ``"F3"``).
        Verilirse formül, ``AG-4/`` benzeri ön eki bu hücreden alır ve
        sayısal kademeyle birleştirir.
    :returns: İç içe IF formülü string'i.
    """
    t = tecrube_yili_hucre
    o = ogrenim_hucre

    def branch(lisans: str, tezsiz: str, tezli: str, doktora: str) -> str:
        return (
            f'IF({o}="Lisans",{lisans},'
            f'IF({o}="Tezsiz Yüksek Lisans",{tezsiz},'
            f'IF({o}="Tezli Yüksek Lisans",{tezli},{doktora})))'
        )

    # 16+ yıl
    ag2 = branch("3", "3", "2", "2")

    # 15-16 yıl
    ag3_15_16 = branch("3", "3", "2", "2")

    # 12-15 yıl
    ag3_12_15 = branch("5", "5", "4", "2")

    # 10-12 yıl
    ag4_10_12 = branch("3", "3", "2", "2")

    # 8-10 yıl
    ag4_8_10 = branch("5", "5", "4", "2")

    # 6-8 yıl
    ag5_6_8 = branch("3", "3", "2", "2")

    # 3-6 yıl
    ag5_3_6 = branch("5", "5", "4", "2")

    # 2-3 yıl
    ag6_2_3 = branch("3", "3", "2", "2")

    # 0-2 yıl
    ag6_0_2 = branch("5", "4", "3", "3")

    kademe_ifadesi = (
        f"IF({t}>=16,{ag2},"
        f"IF({t}>=15,{ag3_15_16},"
        f"IF({t}>=12,{ag3_12_15},"
        f"IF({t}>=10,{ag4_10_12},"
        f"IF({t}>=8,{ag4_8_10},"
        f"IF({t}>=6,{ag5_6_8},"
        f"IF({t}>=3,{ag5_3_6},"
        f"IF({t}>=2,{ag6_2_3},{ag6_0_2}))))))))"
    )
    if not derece_kademe_hucre:
        return f"={kademe_ifadesi}"

    dk = derece_kademe_hucre
    on_ek = f'IFERROR(LEFT({dk},FIND("/",{dk})),"")'
    return f'=IF({dk}="",{kademe_ifadesi},{on_ek}&{kademe_ifadesi})'
