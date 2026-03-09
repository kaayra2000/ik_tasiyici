"""
Tutanak oluşturma iş mantığını kapsülleyen servis katmanı.

Facade tasarım kalıbı ile core katmanına basit bir arayüz sunar.
DIP: TutanakWindow artık doğrudan core fonksiyonlarına bağımlı değildir.
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
from src.core.excel_writer import (
    TutanakOlusturmaRaporu,
    olustur_dk_dosyasi_raporlu,
)


class TutanakService:
    """Tutanak oluşturma iş mantığı Facade'ı.

    Core katmanındaki ``excel_reader`` ve ``excel_writer`` modüllerini
    GUI katmanından soyutlar.
    """

    def __init__(self) -> None:
        self._son_personel_okuma_raporu: PersonelOkumaRaporu | None = None
        self._son_tutanak_olusturma_raporu: TutanakOlusturmaRaporu | None = None

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
        self._son_tutanak_olusturma_raporu = olustur_dk_dosyasi_raporlu(
            personeller=personeller,
            cikti_dizini=output_path_obj.parent,
            dosya_adi=output_path_obj.name,
            template_path=template_path,
            version=version,
        )
        return self._son_tutanak_olusturma_raporu.output_path

    def son_tutanak_olusturma_uyarilari(self) -> list[str]:
        """Son tutanak oluşturma denemesindeki uyarıları döndürür."""
        if self._son_tutanak_olusturma_raporu is None:
            return []
        return list(self._son_tutanak_olusturma_raporu.warning_messages)

    def son_tutanak_olusturma_raporu(self) -> TutanakOlusturmaRaporu | None:
        """Son tutanak oluşturma raporunu döndürür."""
        return self._son_tutanak_olusturma_raporu
