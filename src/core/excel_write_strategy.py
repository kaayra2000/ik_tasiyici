"""
Excel çıktı stratejileri için soyut arayüz.

Strategy Pattern ile farklı şablon versiyonlarını destekler.
Her versiyon bu arayüzü implemente eden bir sınıf olarak tanımlanır.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

from src.core.excel_reader import Personel


class ExcelWriteStrategy(ABC):
    """Soyut Excel yazma stratejisi.

    Her çıktı versiyonu bu sınıfı miras alarak ``sayfa_doldur``
    metodunu implemente etmelidir.
    """

    @abstractmethod
    def sayfa_doldur(self, ws, personel: Personel) -> None:
        """Kopyalanmış şablon çalışma sayfasına personel verisini doldurur.

        :param ws: openpyxl çalışma sayfası nesnesi.
        :param personel: Doldurulacak personel verisi.
        """
        ...
