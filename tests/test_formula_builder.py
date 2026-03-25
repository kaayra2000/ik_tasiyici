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
    brut_ucret_formulu,
    en_yuksek_ogrenim_formulu,
    hizmet_grubu_formulu,
    kademe_baslangic_formulu,
    kademe_bitis_formulu,
    kademe_formulu,
    prim_gunu_formulu,
    tecrube_360_ay_formulu,
    tecrube_360_gun_formulu,
    tecrube_360_yil_formulu,
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
        # TECRUBE_BASLANGIC_SATIR=13, TECRUBE_BITIS_SATIR=27
        assert sonuc == "=SUM(K13:K27)"

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
        # TECRUBE_BASLANGIC_SATIR=13, TECRUBE_BITIS_SATIR=27
        assert sonuc == "=SUM(L13:L27)-SUM(M13:M27)"

    def test_ozel_aralik(self):
        sonuc = toplam_alanda_prim_formulu(bitis_satir=20)
        assert sonuc == "=SUM(L13:L20)-SUM(M13:M20)"


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


class TestTecrube360Formulleri:
    """L28 toplam gün hücresi bazlı yıl/ay/gün formülü testleri."""

    def test_yil_formulu_varsayilan(self):
        sonuc = tecrube_360_yil_formulu()
        assert sonuc == '=IF(L28=0,"",INT(L28/360))'

    def test_ay_formulu_varsayilan(self):
        sonuc = tecrube_360_ay_formulu()
        assert sonuc == '=IF(L28=0,"",INT(MOD(L28,360)/30))'

    def test_gun_formulu_varsayilan(self):
        sonuc = tecrube_360_gun_formulu()
        assert sonuc == '=IF(L28=0,"",MOD(L28,30))'

    def test_l28_hucresine_baglidir(self):
        assert "L28" in tecrube_360_yil_formulu()
        assert "L28" in tecrube_360_ay_formulu()
        assert "L28" in tecrube_360_gun_formulu()


class TestEnYuksekOgrenimFormulu:
    """en_yuksek_ogrenim_formulu fonksiyonu için testler."""

    def test_doktora_ilk_kontrol(self):
        """Doktora her zaman en öncelikli kontrol olmalı."""
        sonuc = en_yuksek_ogrenim_formulu(6, 8, "B", "C", "K")
        assert sonuc.startswith('=IF(OR(AND(B6="Doktora",C6<>"",K6="E"),')

    def test_lisans_son_kontrol(self):
        """Lisans en düşük olduğu için en içteki EĞER'de olmalı."""
        sonuc = en_yuksek_ogrenim_formulu(6, 8, "B", "C", "K")
        assert 'IF(OR(AND(B6="Lisans",C6<>"",K6="E")' in sonuc
        assert sonuc.rfind('"Lisans"') > sonuc.rfind('"Tezsiz Yüksek Lisans"')

    def test_hucre_referanslari_icerir(self):
        """Fonksiyon argümanlarındaki tüm hücre adlarını içermeli."""
        sonuc = en_yuksek_ogrenim_formulu(6, 8, "B", "C", "K")
        for hucre in ["C8", "K8", "B8", "C7", "K7", "B7", "C6", "K6", "B6"]:
            assert hucre in sonuc
        for seviye in [
            "Doktora",
            "Tezli Yüksek Lisans",
            "Tezsiz Yüksek Lisans",
            "Lisans",
        ]:
            assert f'"{seviye}"' in sonuc


class TestUnvanFormulu:
    """unvan_formulu fonksiyonu için testler."""

    def test_kidemli_basuzman_en_ust(self):
        sonuc = unvan_formulu("N28", "M3")
        assert '"Kıdemli Başuzman Araştırmacı"' in sonuc
        assert '"Kıdemli Başuzman"' in sonuc
        # 16+ yıl en baştaki koşul olmalı
        assert sonuc.index("16") < sonuc.index("12")

    def test_uzman_yardimcisi_son(self):
        """Uzman Yardımcısı son seçenek olmalı."""
        sonuc = unvan_formulu("N28", "M3")
        assert '"Uzman Yardımcısı"' in sonuc
        assert 'M3="AG"' in sonuc
        # Diğer ünvanlardan sonra gelmeli
        assert sonuc.index('"Uzman Yardımcısı"') > sonuc.index('"Uzman"')

    def test_tum_unvanlar(self):
        """Formül tüm 5 ünvanı içermeli."""
        sonuc = unvan_formulu("N28", "M3")
        for unvan in [
            "Kıdemli Başuzman",
            "Başuzman",
            "Kıdemli Uzman",
            "Uzman",
            "Uzman Yardımcısı",
            "Araştırmacı",
        ]:
            assert unvan in sonuc

    def test_ag_uzman_yardimcisi_ozel_durumu(self):
        """AG ve 3 yıl altı için yalnızca Araştırmacı dönmeli."""
        sonuc = unvan_formulu("N28", "M3")
        assert 'IF(M3="AG","Araştırmacı","Uzman Yardımcısı")' in sonuc
        assert '"Uzman Yardımcısı Araştırmacı"' not in sonuc


class TestHizmetGrubuFormulu:
    """hizmet_grubu_formulu fonksiyonu için testler."""

    def test_ag2_en_ust(self):
        sonuc = hizmet_grubu_formulu("N28", "M3")
        assert 'M3&"-2"' in sonuc

    def test_ag6_son(self):
        sonuc = hizmet_grubu_formulu("N28", "M3")
        assert 'M3&"-6"' in sonuc
        # -6 en sona gelmeli
        assert sonuc.index('M3&"-6"') > sonuc.index('M3&"-5"')

    def test_tum_gruplar(self):
        """Formül tüm 5 hizmet grubunu içermeli."""
        sonuc = hizmet_grubu_formulu("N28", "M3")
        assert 'OR(M3="A",M3="AG")' in sonuc
        for grup in ['M3&"-2"', 'M3&"-3"', 'M3&"-4"', 'M3&"-5"', 'M3&"-6"']:
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


class TestBrutUcretFormulu:
    """brut_ucret_formulu fonksiyonu için testler."""

    def test_f3_referansi_icerir(self):
        sonuc = brut_ucret_formulu("F3", "$AA$1:$AB$72")
        assert "F3" in sonuc

    def test_ag_ve_a_anahtarlari_icerir(self):
        sonuc = brut_ucret_formulu("F3", "$AA$1:$AB$72")
        assert "$AA$1:$AB$72" in sonuc
        assert "VLOOKUP" in sonuc

    def test_vlookup_kullanir(self):
        sonuc = brut_ucret_formulu("F3", "$AA$1:$AB$72")
        assert sonuc == '=IFERROR(VLOOKUP(F3,$AA$1:$AB$72,2,FALSE),"")'


class TestKademeBaslangicVeBitisFormulu:
    """K30/L30 formullerinde derece on eki davranisini test eder."""

    def test_baslangic_formulu_f3_ile_on_ek_ekler(self):
        sonuc = kademe_baslangic_formulu("Z1", "Z4", "F3")
        assert 'LEFT(F3,FIND("/",F3))' in sonuc
        assert 'IF(F3=""' in sonuc
        assert "Z1" in sonuc
        assert "Z4" in sonuc

    def test_bitis_formulu_f3_ile_on_ek_ekler(self):
        sonuc = kademe_bitis_formulu("Z1", "Z4", "F3")
        assert 'LEFT(F3,FIND("/",F3))' in sonuc
        assert 'IF(F3=""' in sonuc
        assert "Z1" in sonuc
        assert "Z4" in sonuc

    def test_baslangic_formulu_f3_yokken_sayisal_doner(self):
        sonuc = kademe_baslangic_formulu("Z1", "Z4")
        assert 'LEFT(' not in sonuc
        assert 'FIND("/"' not in sonuc

    def test_bitis_formulu_f3_yokken_sayisal_doner(self):
        sonuc = kademe_bitis_formulu("Z1", "Z4")
        assert 'LEFT(' not in sonuc
        assert 'FIND("/"' not in sonuc
