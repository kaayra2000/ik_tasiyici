import os
from pathlib import Path
import PyInstaller.__main__

def main():
    """
    Uygulamayı PyInstaller ile derler.
    """
    print("Derleme başlatılıyor...")
    project_root = Path(__file__).parent.parent
    
    # Derleme dizinini değiştir
    os.chdir(project_root)
    
    PyInstaller.__main__.run([
        'src/main.py',
        '--name=DK_Tutanak_Olusturucu',
        '--windowed',               # Konsol penceresini gizle
        '--noconfirm',              # Sormadan üzerine yaz
        '--clean',                  # Geçici dosyaları temizle
        '--add-data=src/gui/style.qss:src/gui', # QSS dosyasını ekle
        # '--icon=assets/icon.ico', # İkon varsa eklenebilir
    ])
    
    print("Derleme tamamlandı! Çıktılar 'dist' klasöründe.")

if __name__ == "__main__":
    main()
