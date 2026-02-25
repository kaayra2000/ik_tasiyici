"""
Doğrulama (validasyon) fonksiyonları.

Bu modül TC Kimlik Numarası ve diğer giriş verilerinin
doğrulanması için kullanılan fonksiyonları içerir.
"""

from __future__ import annotations


def validate_tckn(tckn: str) -> bool:
    """
    TC Kimlik Numarasının geçerliliğini doğrular.

    Kontrol kuralları:
    1. 11 haneli olmalı
    2. İlk hane 0 olamaz
    3. Sadece rakamlardan oluşmalı
    4. (tek haneler toplamı * 7 - çift haneler toplamı) mod 10 = 10. hane
    5. İlk 10 hanenin toplamının birler basamağı = 11. hane

    :param tckn: Doğrulanacak TC Kimlik Numarası (string).
    :returns: Geçerliyse True, değilse False.
    """
    if not isinstance(tckn, str):
        tckn = str(tckn)

    tckn = tckn.strip()

    if len(tckn) != 11 or not tckn.isdigit():
        return False

    if tckn[0] == "0":
        return False

    digits = [int(d) for d in tckn]

    # 10. hane kontrolü: (tek pozisyonlar toplamı * 7 - çift pozisyonlar toplamı) mod 10
    odd_sum = sum(digits[0:9:2])   # 1., 3., 5., 7., 9. haneler (0-indexed: 0,2,4,6,8)
    even_sum = sum(digits[1:8:2])  # 2., 4., 6., 8. haneler (0-indexed: 1,3,5,7)
    if (odd_sum * 7 - even_sum) % 10 != digits[9]:
        return False

    # 11. hane kontrolü: ilk 10 hanenin toplamının birler basamağı
    if sum(digits[:10]) % 10 != digits[10]:
        return False

    return True


def validate_ad_soyad(ad_soyad: str) -> bool:
    """
    Ad Soyad alanının geçerliliğini doğrular.

    :param ad_soyad: Doğrulanacak ad soyad bilgisi.
    :returns: Geçerliyse True, boş/None ise False.
    """
    if not ad_soyad or not isinstance(ad_soyad, str):
        return False
    return bool(ad_soyad.strip())


def validate_birim(birim: str) -> bool:
    """
    Birim alanının geçerliliğini doğrular.

    :param birim: Doğrulanacak birim/enstitü bilgisi.
    :returns: Geçerliyse True, boş/None ise False.
    """
    if not birim or not isinstance(birim, str):
        return False
    return bool(birim.strip())


def normalize_tckn(tckn: object) -> str:
    """
    TCKN değerini standart 11 haneli string formatına çevirir.

    Excel'den float olarak okunan TCKN değerlerini (ör. 12345678901.0)
    düzgün string'e dönüştürür.

    :param tckn: Dönüştürülecek TCKN değeri (string, int veya float).
    :returns: 11 haneli string TCKN.
    """
    if isinstance(tckn, float):
        tckn = int(tckn)
    return str(tckn).strip().zfill(11)
