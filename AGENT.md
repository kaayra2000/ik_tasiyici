# AGENT.md - DK Tutanak Oluşturucu Projesi

## 🎯 Proje Amacı

Bu proje, 2026 yılında işe başlayan personel için **Derece-Kademe (D-K) Karar Tutanakları** oluşturmayı otomatize eden bir Python uygulamasıdır. Uygulama, kaynak Excel dosyasından personel bilgilerini alır, kullanıcının gireceği mesleki tecrübe ve öğrenim bilgileri için formüllü bir şablon oluşturur ve D-K tablosuna göre ünvan/derece/kademe hesaplama formüllerini otomatik yerleştirir.

---

## 📁 Girdi Dosyaları

### 1. `coklu_girdi.xlsx`
- **Amaç**: Personel temel bilgilerinin kaynağı
- **Kullanılacak Sütunlar**:
  - `TCKN` → TC Kimlik No
  - `AD SOYAD` → Ad Soyad
  - `BİRİMİ` → Enstitü/Birim bilgisi
- **Yaklaşık Kayıt Sayısı**: ~300 (değişken)

### 2. `cikti_ornegi.xlsx`
- **Amaç**: Çıktı dosyası şablonu (Karar Tutanağı formatı)
- **İşaretleme Sistemi**:
  - `o` → Otomatik doldurulacak (kaynak dosyadan)
  - `h` → Hesaplanacak (Excel formülü ile)
  - `e` → Elle girilecek (kullanıcı tarafından)

### 3. `kidem_tablosu.pdf`
- **Amaç**: Ünvan, Derece, Kademe belirleme kuralları
- **Parametreler**:
  - Öğrenim Durumu (Lisans, Tezsiz YL, Tezli YL, Doktora)
  - Kurum Dışı Tecrübe Süresi (yıl)
  - Alanında mı? (E/H)

---

## 📤 Çıktı Dosyası

### `DK_Tutanaklari_2026.xlsx`
- **Format**: Tek Excel dosyası, her personel için ayrı sayfa (sheet)
- **Sayfa Adı Formatı**: `{Ad Soyad} - {TCKN}`
  - Örnek: `Fatma KARACA - 33755123380`
- **Her Sayfa İçeriği**: `cikti_ornegi.xlsx` formatında Karar Tutanağı

---

## 🔄 İş Akışı

```
┌─────────────────────────────────────────────────────────────────┐
│  1. Kullanıcı PyQt GUI'yi başlatır                              │
├─────────────────────────────────────────────────────────────────┤
│  2. "coklu_girdi.xlsx" dosyasını seçer     │
├─────────────────────────────────────────────────────────────────┤
│  3. Uygulama her personel için:                                 │
│     a) TCKN, Ad Soyad, Birim bilgilerini okur                   │
│     b) Şablon formatında yeni sayfa oluşturur                   │
│     c) Otomatik alanları (o) doldurur                           │
│     d) Hesaplama formüllerini (h) yerleştirir                   │
│     e) Elle girilecek alanları (e) boş bırakır                  │
├─────────────────────────────────────────────────────────────────┤
│  4. "DK_Tutanaklari_2026.xlsx" dosyası oluşturulur              │
├─────────────────────────────────────────────────────────────────┤
│  5. Kullanıcı Excel'i açar ve elle girilecek alanları doldurur  │
│     - Öğrenim bilgileri                                         │
│     - Mesleki tecrübeler                                        │
│     - Brüt ücret                                                │
│     - Alanında (E/H) değerlendirmeleri                          │
├─────────────────────────────────────────────────────────────────┤
│  6. Formüller otomatik hesaplar:                                │
│     - Toplam Prim Günü                                          │
│     - Çalışma Alanında Prim Günü                                │
│     - Ünvan, Derece, Kademe                                     │
└─────────────────────────────────────────────────────────────────┘
```

---

## 📊 Excel Formül Gereksinimleri

### 1. Prim Günü Hesaplamaları

#### Toplam Prim Günü (Her Satır İçin)
```excel
=EĞER(VE(D{row}<>"";E{row}<>"");E{row}-D{row};"")
```
- `D{row}` = Başlangıç tarihi
- `E{row}` = Bitiş tarihi
- Sonuç: Gün sayısı (days farkı)

#### Toplam Prim Günü Toplamı
```excel
=TOPLA(K10:K{son_satır})
```

#### Çalışma Alanında Prim Günü (Her Satır İçin)
```excel
=EĞER(J{row}="E";K{row};"")
```
- `J{row}` = Alanında (E/H) sütunu
- `K{row}` = O satırın Toplam Prim Günü

#### Çalışma Alanında Prim Günü Toplamı
```excel
=TOPLA(L10:L{son_satır})
```

### 2. Tecrübe Süresi Hesaplama (Yıl)
```excel
=L{toplam_satır}/360
```

### 3. En Yüksek Alanında Öğrenim Belirleme
```excel
=EĞER(VE(öğrenim_doktora<>"";doktora_alanında="E");"Doktora";
  EĞER(VE(öğrenim_tezli_yl<>"";tezli_yl_alanında="E");"Tezli YL";
    EĞER(VE(öğrenim_tezsiz_yl<>"";tezsiz_yl_alanında="E");"Tezsiz YL";
      EĞER(VE(öğrenim_lisans<>"";lisans_alanında="E");"Lisans";""))))
```

### 4. Ünvan/Derece/Kademe Hesaplama

D-K Tablosuna göre formül mantığı (en alttan yukarı doğru):

```excel
// Tecrübe Yılı = Çalışma Alanında Toplam Prim Günü / 360

// ÜNVAN FORMÜLÜ
=EĞER(tecrübe_yıl>=16;"Kıdemli Başuzman";
  EĞER(tecrübe_yıl>=12;"Başuzman";
    EĞER(tecrübe_yıl>=8;"Kıdemli Uzman";
      EĞER(tecrübe_yıl>=3;"Uzman";
        "Uzman Yardımcısı"))))

// DERECE FORMÜLÜ (Hizmet Grubu)
=EĞER(tecrübe_yıl>=16;"A/AG-2";
  EĞER(tecrübe_yıl>=12;"A/AG-3";
    EĞER(tecrübe_yıl>=8;"A/AG-4";
      EĞER(tecrübe_yıl>=3;"A/AG-5";
        "A/AG-6"))))

// KADEME FORMÜLÜ (Öğrenim durumu ve tecrübeye göre detaylı)
// Bu formül D-K tablosundaki matrise göre oluşturulacak
```

#### Kademe Belirleme Matrisi (PDF'den)

### A/AG-6 - Uzman Yardımcısı/Araştırmacı (0-3 Yıl)

| Tecrübe Süresi | Lisans | Tezsiz Yüksek Lisans | Tezli Yüksek Lisans | Doktora |
|---|---|---|---|---|
| 2 Yıl ve Daha Az | 5-6 | 5-6 | 3 | - |
| 2-3 Yıl Arası | 3-4 | 3-4 | 2 | - |

---

### A/AG-5 - Uzman/Uzman Araştırmacı (3-8 Yıl)

| Tecrübe Süresi | Lisans | Tezsiz Yüksek Lisans | Tezli Yüksek Lisans | Doktora |
|---|---|---|---|---|
| 3-5 Yıl | 5 | 5 | 4 | 2 |
| 6-8 Yıl | 3 | 3 | 2 | 2 |

---

### A/AG-4 - Kıdemli Uzman/Kıdemli Uzman Araştırmacı (8-12 Yıl)

| Tecrübe Süresi | Lisans | Tezsiz Yüksek Lisans | Tezli Yüksek Lisans | Doktora |
|---|---|---|---|---|
| 8-9 Yıl | 5 | 5 | 4 | 3 |
| 10-12 Yıl | 3 | 3 | 3 | 3 |

---

### A/AG-3 - Başuzman/Başuzman Araştırmacı (12-16 Yıl)

| Tecrübe Süresi | Lisans | Tezsiz Yüksek Lisans | Tezli Yüksek Lisans | Doktora |
|---|---|---|---|---|
| 12-14 Yıl | 5 | 5 | 4 | 2 |
| 15-16 Yıl | 3 | 3 | 2 | 2 |

---

### A/AG-2 - Kıdemli Başuzman/Kıdemli Başuzman Araştırmacı (16+ Yıl)

| Tecrübe Süresi | Lisans | Tezsiz Yüksek Lisans | Tezli Yüksek Lisans | Doktora |
|---|---|---|---|---|
| 16 Yıl ve üstü | 4 | 3-4 | 3 | 3 |


---

## ✅ Validasyon Kuralları

### TC Kimlik No Validasyonu
```python
def validate_tckn(tckn: str) -> bool:
    """
    TC Kimlik No doğrulama kuralları:
    1. 11 haneli olmalı
    2. İlk hane 0 olamaz
    3. İlk 10 hanenin toplamının birler basamağı = 11. hane
    4. (1,3,5,7,9. haneler toplamı * 7 - 2,4,6,8. haneler toplamı) mod 10 = 10. hane
    5. İlk 10 hanenin toplamı mod 10 = 11. hane
    """
    if len(tckn) != 11 or not tckn.isdigit():
        return False
    if tckn[0] == '0':
        return False
    
    digits = [int(d) for d in tckn]
    
    # Kural 4: 10. hane kontrolü
    odd_sum = sum(digits[0:9:2])  # 1,3,5,7,9. haneler
    even_sum = sum(digits[1:8:2])  # 2,4,6,8. haneler
    if (odd_sum * 7 - even_sum) % 10 != digits[9]:
        return False
    
    # Kural 5: 11. hane kontrolü
    if sum(digits[:10]) % 10 != digits[10]:
        return False
    
    return True
```

---

## 🛠️ Teknik Gereksinimler

### Python Sürümü
- Python 3.9+

### Bağımlılıklar
```
openpyxl>=3.1.0      # Excel okuma/yazma
pandas>=2.0.0        # Veri işleme
PyQt6>=6.5.0         # GUI framework
```

### Proje Yapısı
```
dk-tutanak-olusturucu/
├── AGENT.md
├── README.md
├── pyproject.toml
├── src/
│   ├── __init__.py
│   ├── main.py              # Ana giriş noktası
│   ├── gui/
│   │   ├── __init__.py
│   │   └── main_window.py   # PyQt ana pencere
│   ├── core/
│   │   ├── __init__.py
│   │   ├── excel_reader.py  # Kaynak Excel okuyucu
│   │   ├── excel_writer.py  # Çıktı Excel oluşturucu
│   │   ├── formula_builder.py  # Excel formül oluşturucu
│   │   └── validators.py    # TCKN ve diğer validasyonlar
│   └── config/
│       ├── __init__.py
│       ├── constants.py     # Sabit değerler
│       └── dk_table.py      # D-K tablosu kuralları
├── templates/
│   └── karar_tutanagi_template.xlsx
├── tests/
│   ├── __init__.py
│   ├── test_validators.py
│   ├── test_formula_builder.py
│   └── test_dk_calculation.py
└── docs/
    └── (örnek dosyalar)
```

---

## 📋 Geliştirme Görevleri

### Faz 1: Temel Altyapı
- [x] Proje yapısını oluştur
- [x] `pyproject.toml` hazırla
- [x] TCKN validasyon fonksiyonunu yaz ve test et
- [x] D-K tablosu kurallarını `dk_table.py`'a aktar

### Faz 2: Excel İşlemleri
- [x] Kaynak Excel okuyucu (`excel_reader.py`)
  - [x] TCKN, Ad Soyad, Birim sütunlarını oku
  - [x] Boş/hatalı satırları filtrele
- [x] Şablon oluşturucu (`excel_writer.py`)
  - [x] Karar Tutanağı formatını oluştur
  - [x] Otomatik alanları doldur
- [x] Formül oluşturucu (`formula_builder.py`)
  - [x] Prim günü formülleri
  - [x] Toplam formülleri
  - [x] Ünvan/Derece/Kademe formülleri

### Faz 3: GUI Geliştirme
- [ ] PyQt6 ana pencere tasarımı
  - [ ] Dosya seçme butonu
  - [ ] İlerleme çubuğu
  - [ ] Log/durum alanı
- [ ] Dosya seçme dialog'u
- [ ] İşlem başlatma/iptal butonları
- [ ] Hata mesajları gösterimi

### Faz 4: Entegrasyon ve Test
- [ ] Tüm modülleri entegre et
- [ ] Unit testleri yaz
- [ ] ~300 kayıtlık test verisiyle performans testi
- [ ] Edge case'leri test et (boş alanlar, geçersiz TCKN vb.)

### Faz 5: Dokümantasyon ve Paketleme
- [ ] README.md hazırla
- [ ] Kullanım kılavuzu yaz
- [ ] Executable oluştur (PyInstaller ile)

---

## ⚠️ Dikkat Edilecek Noktalar

1. **Excel Formülleri Türkçe Olmalı**: `=TOPLA()`, `=EĞER()` gibi Türkçe fonksiyon adları kullanılmalı (kullanıcının Excel dil ayarına göre)

2. **Sayfa Adı Limiti**: Excel sayfa adları max 31 karakter olabilir. `{Ad Soyad} - {TCKN}` formatı uzun olabilir, gerekirse kısaltma yapılmalı

3. **Tarih Formatı**: Excel'de tarih hücreleri doğru formatta olmalı ki formüller çalışsın

4. **Formül Hücre Referansları**: Mesleki tecrübe satır sayısı değişken olabilir, formüller dinamik aralık kullanmalı

5. **Boş Satır Kontrolü**: Formüller boş satırlarda hata vermemeli (`EĞER` ile kontrol)

---

## 🔗 Referans Dosyalar

| Dosya | Açıklama |
|-------|----------|
| `coklu_girdi.xlsx` | Personel listesi kaynağı |
| `cikti_ornegi.xlsx` | Çıktı şablonu |
| `kidem_tablosu.pdf` | Ünvan/Derece/Kademe kuralları |

---

## 📝 Notlar

- Brüt ücret hesaplaması yapılmayacak, kullanıcı elle girecek
- "Alanında" değerlendirmesi kullanıcı tarafından belirlenecek
- Formüller Excel içinde çalışacak (Python sabit değer yazmayacak)
- Çıktı tek dosyada, her personel ayrı sayfada olacak
```