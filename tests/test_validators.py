"""
TCKN ve diğer validasyon fonksiyonları için testler.
"""

import pytest
from src.core.validators import validate_tckn, validate_ad_soyad, validate_birim, normalize_tckn


class TestValidateTCKN:
    """validate_tckn fonksiyonu için test sınıfı."""

    # --- Geçerli TCKN'ler ---
    def test_gecerli_tckn_1(self):
        """Bilinen geçerli bir TCKN kabul edilmeli."""
        assert validate_tckn("10000000146") is True

    def test_gecerli_tckn_2(self):
        """Başka bir bilinen geçerli TCKN (algoritmik olarak türetildi)."""
        # odd=1+0+0+0+0+0=1, even=0+0+0+0+0=0, d10=(1*7-0)%10=7, d11=(1+7)%10=8
        assert validate_tckn("10000000078") is True

    def test_gecerli_tckn_3(self):
        """Üçüncü geçerli TCKN (algoritmik olarak türetildi)."""
        # odd=1+0+0+0+0+5=6, even=0+0+0+0+0=0, d10=(6*7-0)%10=2, d11=(1+5+2)%10=8
        assert validate_tckn("10000050028") is True

    # --- Geçersiz TCKN'ler ---
    def test_11_haneden_az(self):
        """10 haneli TCKN reddedilmeli."""
        assert validate_tckn("1234567890") is False

    def test_11_haneden_fazla(self):
        """12 haneli TCKN reddedilmeli."""
        assert validate_tckn("123456789012") is False

    def test_ilk_hane_sifir(self):
        """İlk hanesi 0 olan TCKN reddedilmeli."""
        assert validate_tckn("01234567890") is False

    def test_harf_iceren(self):
        """Harf içeren TCKN reddedilmeli."""
        assert validate_tckn("1234567890A") is False

    def test_bos_string(self):
        """Boş string reddedilmeli."""
        assert validate_tckn("") is False

    def test_bosluk_iceren(self):
        """Boşluk içeren TCKN reddedilmeli."""
        assert validate_tckn("1234 567890") is False

    def test_yanlis_10_hane(self):
        """10. hanesi yanlış olan TCKN reddedilmeli."""
        assert validate_tckn("10000000137") is False

    def test_yanlis_11_hane(self):
        """11. hanesi yanlış olan TCKN reddedilmeli."""
        assert validate_tckn("10000000145") is False

    def test_none_gibi_deger(self):
        """None benzeri (int 0) değer 11 haneye doldurulamaz, reddedilmeli."""
        # 0 -> "0" -> len < 11 veya ilk hane 0
        assert validate_tckn("0") is False

    def test_string_olarak_gecerli(self):
        """String formatında doğru TCKN kabul edilmeli."""
        assert validate_tckn("10000000146") is True


class TestNormalizeTCKN:
    """normalize_tckn fonksiyonu için test sınıfı."""

    def test_float_donusumu(self):
        """Float değer doğru 11 haneli string'e dönüşmeli."""
        assert normalize_tckn(10000000146.0) == "10000000146"

    def test_int_donusumu(self):
        """Integer değer doğru 11 haneli string'e dönüşmeli."""
        assert normalize_tckn(10000000146) == "10000000146"

    def test_string_bosluk_temizleme(self):
        """Başındaki/sonundaki boşluklar temizlenmeli."""
        assert normalize_tckn(" 10000000146 ") == "10000000146"

    def test_kisa_sayi_doldurma(self):
        """Kısa sayı 11 haneye sıfırla doldurulmalı."""
        assert normalize_tckn(1) == "00000000001"


class TestValidateAdSoyad:
    """validate_ad_soyad fonksiyonu için test sınıfı."""

    def test_gecerli_ad_soyad(self):
        assert validate_ad_soyad("Fatma KARACA") is True

    def test_bos_string(self):
        assert validate_ad_soyad("") is False

    def test_sadece_bosluk(self):
        assert validate_ad_soyad("   ") is False

    def test_none(self):
        assert validate_ad_soyad(None) is False


class TestValidateBirim:
    """validate_birim fonksiyonu için test sınıfı."""

    def test_gecerli_birim(self):
        assert validate_birim("Marmara Enstitüsü") is True

    def test_bos_string(self):
        assert validate_birim("") is False

    def test_none(self):
        assert validate_birim(None) is False
