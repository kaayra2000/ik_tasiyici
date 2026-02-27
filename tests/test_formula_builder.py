"""
formula_builder modülü için testler.

Her formül fonksiyonunun beklenen OOXML uyumlu İngilizce Excel formülünü
üretip üretmediği kontrol edilir.

Not: openpyxl OOXML standardı gereği İngilizce fonksiyon adları kullanır.
LibreOffice/Excel bunları kullanıcı diline otomatik çevirir.
"""

from __future__ import annotations

import pytest

from src.core.formula_builder import (
    alanda_prim_formulu,
    en_yuksek_ogrenim_formulu,
    hizmet_grubu_formulu,
    kademe_formulu,
    prim_gunu_formulu,
    tecrube_yili_formulu,
    toplam_alanda_prim_formulu,
    toplam_prim_formulu,
    unvan_formulu,
)


class TestPrimGunuFormulu:
    """prim_gunu_formulu fonksiyonu için testler."""

    def test_satir_10(self):
        sonuc = prim_gunu_formulu(10)
        assert sonuc == '=IF(AND(E10<>"",F10<>""),F10-E10,"")'

    def test_satir_25(self):
        sonuc = prim_gunu_formulu(25)
        assert sonuc == '=IF(AND(E25<>"",F25<>""),F25-E25,"")'

    def test_egitim_iceriyor(self):
        """Formül IF ve AND anahtar kelimelerini içermeli (OOXML standardı)."""
        sonuc = prim_gunu_formulu(15)
        assert "IF" in sonuc
        assert "AND" in sonuc


class TestAlandaPrimFormulu:
    """alanda_prim_formulu fonksiyonu için testler."""

    def test_satir_10(self):
        sonuc = alanda_prim_formulu(10)
        assert sonuc == '=IF(J10="E",K10,"")'

    def test_satir_20(self):
        sonuc = alanda_prim_formulu(20)
        assert sonuc == '=IF(J20="E",K20,"")'

    def test_e_harfi_iceriyor(self):
        """Formül 'E' koşulunu içermeli."""
        sonuc = alanda_prim_formulu(12)
        assert '"E"' in sonuc


class TestToplamPrimFormulu:
    """toplam_prim_formulu fonksiyonu için testler."""

    def test_varsayilan_aralik(self):
        sonuc = toplam_prim_formulu()
        # TECRUBE_BASLANGIC_SATIR=11, TECRUBE_BITIS_SATIR=18
        assert sonuc == "=SUM(K11:K18)"

    def test_ozel_aralik(self):
        sonuc = toplam_prim_formulu(bitis_satir=30, baslangic_satir=10)
        assert sonuc == "=SUM(K10:K30)"

    def test_topla_iceriyor(self):
        """OOXML standardında SUM kullanılır."""
        assert "SUM" in toplam_prim_formulu()


class TestToplamAlandaPrimFormulu:
    """toplam_alanda_prim_formulu fonksiyonu için testler."""

    def test_varsayilan_aralik(self):
        sonuc = toplam_alanda_prim_formulu()
        # TECRUBE_BASLANGIC_SATIR=11, TECRUBE_BITIS_SATIR=18
        assert sonuc == "=SUM(L11:L18)"

    def test_ozel_aralik(self):
        sonuc = toplam_alanda_prim_formulu(bitis_satir=20)
        assert sonuc == "=SUM(L11:L20)"


class TestTecrübeYiliFormulu:
    """tecrube_yili_formulu fonksiyonu için testler."""

    def test_l27_hucresi(self):
        sonuc = tecrube_yili_formulu("L27")
        assert sonuc == "=L27/360"

    def test_farkli_hucre(self):
        sonuc = tecrube_yili_formulu("L30")
        assert sonuc == "=L30/360"

    def test_360_ile_bolme(self):
        assert "/360" in tecrube_yili_formulu("L99")


class TestEnYuksekOgrenimFormulu:
    """en_yuksek_ogrenim_formulu fonksiyonu için testler."""

    def test_doktora_birinci_kontrol(self):
        """Doktora kontrolü formülde en başta gelmeli."""
        sonuc = en_yuksek_ogrenim_formulu(
            "C8", "K8", "C7", "K7", "C6", "K6"
        )
        # Doktora ilk EĞER'de olmalı
        assert '"Doktora"' in sonuc
        assert sonuc.index('"Doktora"') < sonuc.index('"Tezli Yüksek Lisans"')

    def test_lisans_son_kontrol(self):
        """Lisans en sondaki EĞER'de olmalı."""
        sonuc = en_yuksek_ogrenim_formulu(
            "C8", "K8", "C7", "K7", "C6", "K6"
        )
        assert '"Lisans"' in sonuc
        assert sonuc.index('"Lisans"') > sonuc.index('"Tezli Yüksek Lisans"')

    def test_hucre_referanslari_icerir(self):
        """Fonksiyon argümanlarındaki hücre adlarını içermeli."""
        sonuc = en_yuksek_ogrenim_formulu(
            "C8", "K8", "C7", "K7", "C6", "K6"
        )
        for hucre in ["C8", "K8", "C7", "K7", "C6", "K6"]:
            assert hucre in sonuc


class TestUnvanFormulu:
    """unvan_formulu fonksiyonu için testler."""

    def test_kidemli_basuzman_en_ust(self):
        sonuc = unvan_formulu("N28")
        assert '"Kıdemli Başuzman"' in sonuc
        # 16+ yıl en baştaki koşul olmalı
        assert sonuc.index("16") < sonuc.index("12")

    def test_uzman_yardimcisi_son(self):
        """Uzman Yardımcısı son seçenek olmalı."""
        sonuc = unvan_formulu("N28")
        assert '"Uzman Yardımcısı"' in sonuc
        # Diğer ünvanlardan sonra gelmeli
        assert sonuc.index('"Uzman Yardımcısı"') > sonuc.index('"Uzman"')

    def test_tum_unvanlar(self):
        """Formül tüm 5 ünvanı içermeli."""
        sonuc = unvan_formulu("N28")
        for unvan in [
            "Kıdemli Başuzman", "Başuzman", "Kıdemli Uzman", "Uzman", "Uzman Yardımcısı"
        ]:
            assert unvan in sonuc


class TestHizmetGrubuFormulu:
    """hizmet_grubu_formulu fonksiyonu için testler."""

    def test_ag2_en_ust(self):
        sonuc = hizmet_grubu_formulu("N28")
        assert '"A/AG-2"' in sonuc

    def test_ag6_son(self):
        sonuc = hizmet_grubu_formulu("N28")
        assert '"A/AG-6"' in sonuc
        # A/AG-6 en sona gelmeli
        assert sonuc.index('"A/AG-6"') > sonuc.index('"A/AG-5"')

    def test_tum_gruplar(self):
        """Formül tüm 5 hizmet grubunu içermeli."""
        sonuc = hizmet_grubu_formulu("N28")
        for grup in ["A/AG-2", "A/AG-3", "A/AG-4", "A/AG-5", "A/AG-6"]:
            assert grup in sonuc


class TestKademeFormulu:
    """kademe_formulu fonksiyonu için testler."""

    def test_formul_egitim_hucresine_bakar(self):
        """Formül öğrenim hücresini içermeli."""
        sonuc = kademe_formulu("N28", "C8")
        assert "C8" in sonuc

    def test_formul_tecrube_hucresine_bakar(self):
        """Formül tecrübe yılı hücresini içermeli."""
        sonuc = kademe_formulu("N28", "C8")
        assert "N28" in sonuc

    def test_lisans_kademesi_icerir(self):
        """Formül Lisans ile eşleşen kademe değerlerini içermeli."""
        sonuc = kademe_formulu("N28", "C8")
        assert "Lisans" in sonuc

    def test_doktora_ag6_tezli_yl_ile_ayni(self):
        """
        Doktora, A/AG-6 grubunda artık Tezli Yüksek Lisans ile aynı kademe değerlerini
        döndürür: t<2 için '3', 2≤t<3 için '2'.
        """
        sonuc = kademe_formulu("N28", "C8")
        assert "Tezli Yüksek Lisans" in sonuc
        # Son dal artık boş değil — "3" ve "2" içermeli
        assert '"3"' in sonuc
        assert '"2"' in sonuc

    def test_formul_egitim_string(self):
        """Formül eğitim türü kontrolü yapmalı."""
        sonuc = kademe_formulu("N28", "C8")
        assert "Tezli Yüksek Lisans" in sonuc
