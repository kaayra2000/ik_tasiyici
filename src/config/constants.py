"""
Uygulama genelinde kullanılan sabit değerler.

Bu modül sütun adları, varsayılan değerler ve diğer
sabit tanımları içerir.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Kaynak Excel sütun adları
# ---------------------------------------------------------------------------

#: TC Kimlik Numarası sütunu
COL_TCKN: str = "TCKN"

#: Ad Soyad sütunu
COL_AD_SOYAD: str = "AD SOYAD"

#: Birim sütunu
COL_BIRIM: str = "BİRİMİ"

# ---------------------------------------------------------------------------
# Çıktı dosyası
# ---------------------------------------------------------------------------

#: Çıktı Excel dosyasının varsayılan adı
OUTPUT_FILENAME: str = "DK_cikti.xlsx"

#: Excel sayfa adı için maksimum uzunluk sınırı
MAX_SHEET_NAME_LEN: int = 31

#: Kullanılacak excel şablon dosyası
TEMPLATE_PATH: str = "docs/cikti_taslagi_dolu.xlsx"

# ---------------------------------------------------------------------------
# Mesleki tecrübe satır aralığı (karar tutanağı şablonunda)
# ---------------------------------------------------------------------------

#: Tecrübe satırlarının başladığı satır numarası (1-indexed).
#: Satır 10 başlık satırı olduğundan veriler 11'den başlar.
TECRUBE_BASLANGIC_SATIR: int = 11

#: Tecrübe satırlarının bittiği satır numarası (varsayılan)
TECRUBE_BITIS_SATIR: int = 18

# ---------------------------------------------------------------------------
# Sütun harfleri (karar tutanağı şablonunda)
# ---------------------------------------------------------------------------

#: Başlangıç tarihi sütunu
COL_BASLANGIC_TARIHI: str = "E"

#: Bitiş tarihi sütunu
COL_BITIS_TARIHI: str = "F"

#: Alanında (E/H) sütunu
COL_ALANINDA: str = "J"

#: Toplam prim günü sütunu
COL_TOPLAM_PRIM: str = "K"

#: Çalışma alanında prim günü sütunu
COL_ALANDA_PRIM: str = "L"

# ---------------------------------------------------------------------------
# Öğrenim durumu değerleri
# ---------------------------------------------------------------------------

OGRENIM_LISANS: str = "Lisans"
OGRENIM_TEZSIZ_YL: str = "Tezsiz Yüksek Lisans"
OGRENIM_TEZLI_YL: str = "Tezli Yüksek Lisans"
OGRENIM_DOKTORA: str = "Doktora"

OGRENIM_SEVIYELERI: list[str] = [
    OGRENIM_LISANS,
    OGRENIM_TEZSIZ_YL,
    OGRENIM_TEZLI_YL,
    OGRENIM_DOKTORA,
]

# ---------------------------------------------------------------------------
# Tecrübe yılı eşikleri
# ---------------------------------------------------------------------------

#: 1 yılı oluşturan prim günü sayısı
GUN_PER_YIL: int = 360
