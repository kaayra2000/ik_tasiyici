"""
Uygulama ayarlarını yöneten modül.

QSettings üzerinde Facade sağlayarak ayar okuma/yazma işlemlerini
tek bir sorumluluk altında toplar (SRP).
"""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QSettings


class SettingsManager:
    """Uygulama ayarlarını yöneten Facade sınıfı.

    :param org: Organizasyon adı (QSettings için).
    :param app: Uygulama adı (QSettings için).
    """

    KEY_INPUT_PATH = "last_input_path"
    KEY_TEMPLATE_PATH = "last_template_path"
    KEY_OUTPUT_PATH = "last_output_path"

    def __init__(
        self, org: str = "IK_Tasiyici", app: str = "TutanakOlusturucu"
    ) -> None:
        self._settings = QSettings(org, app)

    # ------------------------------------------------------------------
    # Genel okuma / yazma
    # ------------------------------------------------------------------

    def get(self, key: str, default: str = "") -> str:
        """Belirtilen anahtarın değerini döner.

        :param key: Ayar anahtarı.
        :param default: Anahtar yoksa dönecek varsayılan değer.
        :returns: Kaydedilmiş değer ya da *default*.
        """
        return self._settings.value(key, default)

    def set(self, key: str, value: str) -> None:
        """Belirtilen anahtara değer yazar.

        :param key: Ayar anahtarı.
        :param value: Kaydedilecek değer.
        """
        self._settings.setValue(key, value)

    # ------------------------------------------------------------------
    # Yardımcı doğrulama
    # ------------------------------------------------------------------

    def get_existing_file(self, key: str) -> str:
        """Kaydedilmiş yol geçerli bir dosyaysa döner, değilse boş string.

        :param key: Ayar anahtarı.
        :returns: Dosya yolu ya da ``""``.
        """
        path = self.get(key)
        if path and Path(path).is_file():
            return path
        return ""

    def get_parent_dir(self, key: str) -> str:
        """Kaydedilmiş yolun üst dizinini döner (dialog başlangıcı için).

        :param key: Ayar anahtarı.
        :returns: Üst dizin yolu ya da ``""``.
        """
        path = self.get(key)
        if path:
            parent = Path(path).parent
            if parent.is_dir():
                return str(parent)
        return ""
