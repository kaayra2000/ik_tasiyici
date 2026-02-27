"""
Excel yazma stratejisi fabrikası.

Factory Method ile versiyon string'ine göre doğru stratejiyi döndürür.
Yeni bir versiyon eklemek için sadece bu dosyadaki ``_STRATEGIES``
sözlüğüne yeni sınıfı kaydetmek yeterlidir.
"""

from __future__ import annotations

from src.core.excel_write_strategy import ExcelWriteStrategy
from src.core.excel_write_strategy_v1 import ExcelWriteStrategyV1


# Desteklenen strateji sınıfları
_STRATEGIES: dict[str, type[ExcelWriteStrategy]] = {
    "v1": ExcelWriteStrategyV1,
}


class ExcelWriterFactory:
    """Versiyon bazlı Excel yazma stratejisi fabrikası."""

    @staticmethod
    def create(version: str) -> ExcelWriteStrategy:
        """Belirtilen versiyona uygun strateji nesnesini oluşturur.

        :param version: Versiyon tanımlayıcısı (ör. ``"v1"``).
        :returns: İlgili :class:`ExcelWriteStrategy` alt sınıfı örneği.
        :raises ValueError: Desteklenmeyen versiyon verilirse.
        """
        strategy_cls = _STRATEGIES.get(version)
        if strategy_cls is None:
            desteklenen = ", ".join(sorted(_STRATEGIES.keys()))
            raise ValueError(
                f"Desteklenmeyen çıktı versiyonu: '{version}'. "
                f"Desteklenen versiyonlar: {desteklenen}"
            )
        return strategy_cls()
