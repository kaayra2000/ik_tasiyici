#!/usr/bin/env python3
"""
Faz 2 modüllerini (excel_reader, excel_writer) test etmek için
kullanılabilecek basit bir deneme betiğidir.

Kullanımı:
$ source .venv/bin/activate
$ python demo.py
"""

from pathlib import Path
from src.core.excel_reader import oku_personel_listesi
from src.core.excel_writer import olustur_dk_dosyasi

def main():
    # Kaynak dosya yolu (kendi yolunuza göre güncelleyebilirsiniz)
    input_file = "docs/coklu_girdi.xlsx"
    
    print(f"[{input_file}] dosyası okunuyor...")
    try:
        personeller = oku_personel_listesi(input_file)
        print(f"Başarılı: {len(personeller)} geçerli personel okundu.")
        
        if personeller:
            print("DK tutanakları Excel dosyası oluşturuluyor...")
            output_path = olustur_dk_dosyasi(personeller)
            print(f"Başarılı: Çıktı dosyası oluşturuldu -> {output_path}")
        else:
            print("Uyarı: Geçerli personel bulunamadığı için dosya oluşturulmadı.")
            
    except FileNotFoundError as e:
        print(f"Hata: Kaynak dosya bulunamadı. Lütfen 'docs' klasöründe doğru dosyanın olduğundan emin olun.\n({e})")
    except Exception as e:
        print(f"Beklenmeyen bir hata oluştu: {e}")

if __name__ == "__main__":
    main()
