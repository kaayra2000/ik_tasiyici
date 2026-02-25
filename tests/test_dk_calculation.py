"""
D-K tablosu hesaplama fonksiyonları için testler.
"""

import pytest
from src.config.dk_table import (
    hesapla_tecrube_yili,
    belirle_hizmet_grubu,
    belirle_unvan,
    belirle_kademe,
    hesapla_dk,
    DKSonuc,
)
from src.config.constants import (
    OGRENIM_LISANS,
    OGRENIM_TEZSIZ_YL,
    OGRENIM_TEZLI_YL,
    OGRENIM_DOKTORA,
    GUN_PER_YIL,
)


class TestHesaplaTecrübeYili:
    """hesapla_tecrube_yili fonksiyonu için testler."""

    def test_sifir_gun(self):
        assert hesapla_tecrube_yili(0) == 0.0

    def test_bir_yil(self):
        assert hesapla_tecrube_yili(360) == pytest.approx(1.0)

    def test_bes_yil(self):
        assert hesapla_tecrube_yili(1800) == pytest.approx(5.0)

    def test_onalti_yil(self):
        assert hesapla_tecrube_yili(5760) == pytest.approx(16.0)

    def test_kesirli_yil(self):
        assert hesapla_tecrube_yili(180) == pytest.approx(0.5)


class TestBelirleHizmetGrubu:
    """belirle_hizmet_grubu fonksiyonu için testler."""

    def test_sifir_yil(self):
        assert belirle_hizmet_grubu(0.0) == "A/AG-6"

    def test_iki_yil(self):
        assert belirle_hizmet_grubu(2.0) == "A/AG-6"

    def test_uc_yil(self):
        assert belirle_hizmet_grubu(3.0) == "A/AG-5"

    def test_alti_yil(self):
        assert belirle_hizmet_grubu(6.0) == "A/AG-5"

    def test_sekiz_yil(self):
        assert belirle_hizmet_grubu(8.0) == "A/AG-4"

    def test_oniki_yil(self):
        assert belirle_hizmet_grubu(12.0) == "A/AG-3"

    def test_onalti_yil(self):
        assert belirle_hizmet_grubu(16.0) == "A/AG-2"

    def test_yirmi_yil(self):
        assert belirle_hizmet_grubu(20.0) == "A/AG-2"


class TestBelirleUnvan:
    """belirle_unvan fonksiyonu için testler."""

    def test_ag6(self):
        assert belirle_unvan("A/AG-6") == "Uzman Yardımcısı"

    def test_ag5(self):
        assert belirle_unvan("A/AG-5") == "Uzman"

    def test_ag4(self):
        assert belirle_unvan("A/AG-4") == "Kıdemli Uzman"

    def test_ag3(self):
        assert belirle_unvan("A/AG-3") == "Başuzman"

    def test_ag2(self):
        assert belirle_unvan("A/AG-2") == "Kıdemli Başuzman"

    def test_bilinmeyen_grup(self):
        with pytest.raises(ValueError):
            belirle_unvan("A/AG-9")


class TestBelirleKademe:
    """belirle_kademe fonksiyonu için testler."""

    # A/AG-6 testleri
    def test_ag6_lisans_0_yil(self):
        assert belirle_kademe(0.5, "A/AG-6", OGRENIM_LISANS) == "5-6"

    def test_ag6_tezli_yl_0_yil(self):
        assert belirle_kademe(0.5, "A/AG-6", OGRENIM_TEZLI_YL) == "3"

    def test_ag6_lisans_2_5_yil(self):
        assert belirle_kademe(2.5, "A/AG-6", OGRENIM_LISANS) == "3-4"

    def test_ag6_doktora_returns_none(self):
        """A/AG-6 için Doktora kademesi yoktur."""
        assert belirle_kademe(0.5, "A/AG-6", OGRENIM_DOKTORA) is None

    # A/AG-5 testleri
    def test_ag5_lisans_3_yil(self):
        assert belirle_kademe(3.0, "A/AG-5", OGRENIM_LISANS) == "5"

    def test_ag5_doktora_4_yil(self):
        assert belirle_kademe(4.0, "A/AG-5", OGRENIM_DOKTORA) == "2"

    def test_ag5_tezli_yl_7_yil(self):
        assert belirle_kademe(7.0, "A/AG-5", OGRENIM_TEZLI_YL) == "2"

    # A/AG-4 testleri
    def test_ag4_lisans_8_yil(self):
        assert belirle_kademe(8.0, "A/AG-4", OGRENIM_LISANS) == "5"

    def test_ag4_doktora_10_yil(self):
        assert belirle_kademe(10.0, "A/AG-4", OGRENIM_DOKTORA) == "3"

    # A/AG-3 testleri
    def test_ag3_tezli_yl_13_yil(self):
        assert belirle_kademe(13.0, "A/AG-3", OGRENIM_TEZLI_YL) == "4"

    def test_ag3_lisans_15_yil(self):
        assert belirle_kademe(15.0, "A/AG-3", OGRENIM_LISANS) == "3"

    # A/AG-2 testleri
    def test_ag2_lisans_16_yil(self):
        assert belirle_kademe(16.0, "A/AG-2", OGRENIM_LISANS) == "4"

    def test_ag2_tezsiz_yl_20_yil(self):
        assert belirle_kademe(20.0, "A/AG-2", OGRENIM_TEZSIZ_YL) == "3-4"

    def test_ag2_doktora_18_yil(self):
        assert belirle_kademe(18.0, "A/AG-2", OGRENIM_DOKTORA) == "3"


class TestHesaplaDK:
    """hesapla_dk fonksiyonu entegrasyon testleri."""

    def test_uzman_yardimcisi_lisans(self):
        """1 yıl Lisans → A/AG-6, Uzman Yardımcısı, kademe 5-6."""
        sonuc = hesapla_dk(360.0, OGRENIM_LISANS)
        assert sonuc.unvan == "Uzman Yardımcısı"
        assert sonuc.hizmet_grubu == "A/AG-6"
        assert sonuc.kademe == "5-6"

    def test_uzman_tezli_yl(self):
        """4 yıl Tezli YL → A/AG-5, Uzman, kademe 4."""
        sonuc = hesapla_dk(4 * GUN_PER_YIL, OGRENIM_TEZLI_YL)
        assert sonuc.unvan == "Uzman"
        assert sonuc.hizmet_grubu == "A/AG-5"
        assert sonuc.kademe == "4"

    def test_kidemli_uzman_doktora(self):
        """9 yıl Doktora → A/AG-4, Kıdemli Uzman, kademe 3."""
        sonuc = hesapla_dk(9 * GUN_PER_YIL, OGRENIM_DOKTORA)
        assert sonuc.unvan == "Kıdemli Uzman"
        assert sonuc.hizmet_grubu == "A/AG-4"
        assert sonuc.kademe == "3"

    def test_basuzman_lisans(self):
        """13 yıl Lisans → A/AG-3, Başuzman, kademe 5."""
        sonuc = hesapla_dk(13 * GUN_PER_YIL, OGRENIM_LISANS)
        assert sonuc.unvan == "Başuzman"
        assert sonuc.hizmet_grubu == "A/AG-3"
        assert sonuc.kademe == "5"

    def test_kidemli_basuzman_doktora(self):
        """17 yıl Doktora → A/AG-2, Kıdemli Başuzman, kademe 3."""
        sonuc = hesapla_dk(17 * GUN_PER_YIL, OGRENIM_DOKTORA)
        assert sonuc.unvan == "Kıdemli Başuzman"
        assert sonuc.hizmet_grubu == "A/AG-2"
        assert sonuc.kademe == "3"

    def test_doktora_ag6_raises(self):
        """A/AG-6 kademesinde Doktora için hesaplama hata vermeli."""
        with pytest.raises(ValueError):
            hesapla_dk(180.0, OGRENIM_DOKTORA)
