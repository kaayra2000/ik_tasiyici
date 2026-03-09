"""
Tutanak oluşturma iş mantığını kapsülleyen servis katmanı.

Facade tasarım kalıbı ile core katmanına basit bir arayüz sunar.
DIP: MainWindow artık doğrudan core fonksiyonlarına bağımlı değildir.
"""

from __future__ import annotations

from pathlib import Path
from typing import List

from src.config.constants import DEFAULT_VERSION
from src.core.excel_reader import (
    Personel,
    PersonelOkumaRaporu,
    oku_personel_listesi_raporlu,
)
from src.core.excel_writer import olustur_dk_dosyasi


class TutanakService:
    """Tutanak oluşturma iş mantığı Facade'ı.

    Core katmanındaki ``excel_reader`` ve ``excel_writer`` modüllerini
    GUI katmanından soyutlar.
    """

    def __init__(self) -> None:
        self._son_personel_okuma_raporu: PersonelOkumaRaporu | None = None

    def personel_oku(self, input_path: str) -> List[Personel]:
        """Kaynak Excel dosyasından personel listesini okur.

        :param input_path: Kaynak dosya yolu.
        :returns: Geçerli personel nesnelerinin listesi.
        :raises FileNotFoundError: Dosya bulunamazsa.
        :raises ValueError: Zorunlu sütunlar eksikse.
        """
        self._son_personel_okuma_raporu = oku_personel_listesi_raporlu(input_path)
        return self._son_personel_okuma_raporu.personeller

    def son_personel_okuma_raporu(self) -> PersonelOkumaRaporu | None:
        """Son okuma işlemine ait detaylı raporu döndürür."""
        return self._son_personel_okuma_raporu

    def son_personel_okuma_uyarilari(self) -> list[str]:
        """Son okuma sırasında atlanan satırlar için log mesajları üretir."""
        if self._son_personel_okuma_raporu is None:
            return []
        return [
            satir_reddi.log_mesaji
            for satir_reddi in self._son_personel_okuma_raporu.reddedilen_satirlar
        ]

    def tutanak_olustur(
        self,
        personeller: List[Personel],
        template_path: str,
        output_path: str,
        version: str = DEFAULT_VERSION,
    ) -> Path:
        """DK tutanak dosyasını oluşturur.

        :param personeller: İşlenecek personel listesi.
        :param template_path: Çıktı taslağı dosya yolu.
        :param output_path: Çıktı dosyasının tam yolu.
        :param version: Çıktı versiyonu (ör. ``"v1"``).
        :returns: Oluşturulan dosyanın tam yolu.
        """
        output_path_obj = Path(output_path)
        return olustur_dk_dosyasi(
            personeller=personeller,
            cikti_dizini=output_path_obj.parent,
            dosya_adi=output_path_obj.name,
            template_path=template_path,
            version=version,
        )
