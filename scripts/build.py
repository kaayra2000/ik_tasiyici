import os
import tomllib
import subprocess
from pathlib import Path
import PyInstaller.__main__


def get_version(project_root: Path) -> str:
    """pyproject.toml dosyasından sürüm bilgisini okur ve version.txt'ye yazar."""
    with open(project_root / "pyproject.toml", "rb") as f:
        config = tomllib.load(f)
        version = config["project"]["version"]

    version_file = project_root / "version.txt"
    version_file.write_text(version, encoding="utf-8")
    return version


def get_release_notes(project_root: Path, version: str) -> None:
    """Git tag üzerinden sürüm notlarını çeker ve release_notes.txt'ye yazar."""
    release_notes_file = project_root / "release_notes.txt"
    try:
        result = subprocess.run(
            ["git", "tag", "-l", "--format=%(contents)", f"v{version}"],
            capture_output=True,
            text=True,
            check=True,
        )
        notes = result.stdout.strip()
        if not notes:
            notes = "Sürüm notu bulunamadı."
    except Exception:
        notes = "Sürüm notu bulunamadı."
    release_notes_file.write_text(notes, encoding="utf-8")


def run_pyinstaller(exe_name: str) -> None:
    """PyInstaller'ı çalıştırarak derleme işlemini gerçekleştirir."""
    PyInstaller.__main__.run(
        [
            "src/main.py",
            f"--name={exe_name}",
            "--windowed",  # Konsol penceresini gizle
            "--onefile",  # Tek bir çalıştırılabilir dosya oluştur
            "--noconfirm",  # Sormadan üzerine yaz
            "--clean",  # Geçici dosyaları temizle
            "--add-data=src/gui/style.qss:src/gui",  # QSS dosyasını ekle
            "--add-data=src/assets/tubitak_logo.png:src/assets",  # Logo dosyasını ekle
            "--add-data=version.txt:.",  # Versiyon dosyasını ana dizine ekle
            "--add-data=release_notes.txt:.",  # Sürüm notu dosyasını ana dizine ekle
            "--icon=src/assets/tubitak_logo.png",  # EXE ikonunu ayarla
        ]
    )


def main() -> None:
    """
    Uygulamayı PyInstaller ile derler.
    """
    print("Derleme başlatılıyor...")
    project_root = Path(__file__).parent.parent

    # Derleme dizinini değiştir
    os.chdir(project_root)

    version = get_version(project_root)
    get_release_notes(project_root, version)

    exe_name = f"DK_Tutanak_Olusturucu_v{version}"
    run_pyinstaller(exe_name)

    print("Derleme tamamlandı! Çıktılar 'dist' klasöründe.")


if __name__ == "__main__":
    main()
