#!/usr/bin/env python3
"""
Faz 2 modüllerini (excel_reader, excel_writer) test etmek için
kullanılabilecek basit bir deneme betiğidir.

Kullanımı:
$ source .venv/bin/activate
$ python demo.py
"""

from pathlib import Path
from src.core.excel_reader import oku_personel_listesi_raporlu
from src.core.excel_writer import olustur_dk_dosyasi_raporlu


def main():
    # Kaynak dosya yolu (kendi yolunuza göre güncelleyebilirsiniz)
    input_file = "docs/coklu_girdi.xlsx"

    print(f"[{input_file}] dosyası okunuyor...")
    try:
        rapor = oku_personel_listesi_raporlu(input_file)
        personeller = rapor.personeller
        print(f"Başarılı: {len(personeller)} geçerli personel okundu.")

        if personeller:
            print("DK tutanakları Excel dosyası oluşturuluyor...")
            tutanak_rapor = olustur_dk_dosyasi_raporlu(personeller)
            print(f"Başarılı: Çıktı dosyası oluşturuldu -> {tutanak_rapor.output_path}")
        else:
            print("Uyarı: Geçerli personel bulunamadığı için dosya oluşturulmadı.")

    except FileNotFoundError as e:
        print(
            f"Hata: Kaynak dosya bulunamadı. Lütfen 'docs' klasöründe doğru dosyanın olduğundan emin olun.\n({e})"
        )
    except Exception as e:
        print(f"Beklenmeyen bir hata oluştu: {e}")


if __name__ == "__main__":
    main()
