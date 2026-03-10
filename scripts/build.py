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
    
    import tomllib
    
    with open('pyproject.toml', 'rb') as f:
        config = tomllib.load(f)
        version = config['project']['version']
        
    # Sürüm numarasını çalışma anında okunabilmesi için bir dosyaya yaz
    version_file = project_root / 'version.txt'
    version_file.write_text(version, encoding='utf-8')
        
    PyInstaller.__main__.run([
        'src/main.py',
        '--name=DK_Tutanak_Olusturucu',
        '--windowed',               # Konsol penceresini gizle
        '--onefile',                # Tek bir çalıştırılabilir dosya oluştur
        '--noconfirm',              # Sormadan üzerine yaz
        '--clean',                  # Geçici dosyaları temizle
        '--add-data=src/gui/style.qss:src/gui', # QSS dosyasını ekle
        '--add-data=src/assets/tubitak_logo.png:src/assets', # Logo dosyasını ekle
        '--add-data=version.txt:.', # Versiyon dosyasını ana dizine ekle
        '--icon=src/assets/tubitak_logo.png', # EXE ikonunu ayarla
    ])
    
    print("Derleme tamamlandı! Çıktılar 'dist' klasöründe.")

if __name__ == "__main__":
    main()
