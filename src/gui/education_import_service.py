"""Mezuniyet içe aktarma iş mantığını GUI'den soyutlayan servis katmanı."""

from __future__ import annotations

from src.core.education_importer import (
    EducationImporter,
    EducationImportResult,
)


class EducationImportService:
    """GUI için mezuniyet içe aktarma facade'ı."""

    def __init__(self, importer: EducationImporter | None = None) -> None:
        self._importer = importer or EducationImporter()

    def import_education(
        self,
        source_path: str,
        target_path: str,
    ) -> EducationImportResult:
        """Kaynak mezuniyet dosyasını seçili tutanağa işler."""
        return self._importer.import_education(
            source_path=source_path,
            target_path=target_path,
        )

    def son_import_uyarilari(self) -> list[str]:
        """Son mezuniyet içe aktarma denemesindeki uyarıları döndürür."""
        return self._importer.last_warning_messages()
