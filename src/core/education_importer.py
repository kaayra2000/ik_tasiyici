"""Mezuniyet bilgilerinin mevcut tutanaklara aktarımını yöneten modül."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from shutil import copy2

import openpyxl
import pandas as pd

from src.config.constants import (
    OGRENIM_DOKTORA,
    OGRENIM_LISANS,
    OGRENIM_TEZLI_YL,
    OGRENIM_TEZSIZ_YL,
)
from src.core.validators import normalize_tckn, validate_tckn

_SOURCE_COL_TCKN = "TC KIMLIK NO"
_SOURCE_COL_AD = "AD"
_SOURCE_COL_MEZUNIYET_TARIHI = "MEZUNIYET TARIHI"
_SOURCE_COL_UNIVERSITE = "UNIVERSITE"
_SOURCE_COL_FAKULTE = "ENSMYOFAK"
_SOURCE_COL_PROGRAM = "PROGRAM"
_MISSING_RECORD_MARKER = "MEZUN KAYDI BULUNAMADI"

_DEFAULT_EDUCATION_ROWS = [6, 7, 8, 9, 10]
_EDUCATION_LEVEL_PRIORITY = {
    OGRENIM_LISANS: 0,
    OGRENIM_TEZSIZ_YL: 1,
    OGRENIM_TEZLI_YL: 2,
    OGRENIM_DOKTORA: 3,
}


@dataclass(frozen=True)
class EducationRecord:
    """Tek bir mezuniyet kaydını temsil eder."""

    tckn: str
    level: str
    school_text: str
    department_text: str
    graduation_date: str = ""

    @property
    def fingerprint(self) -> tuple[str, str, str]:
        """Kaydı eşsizleştirmek için kullanılacak parmak izi."""
        return (
            self.level.strip().casefold(),
            self.school_text.strip().casefold(),
            self.department_text.strip().casefold(),
        )


@dataclass
class EducationImportResult:
    """İçe aktarma işleminin özetini taşır."""

    backup_path: Path
    matched_sheet_count: int = 0
    updated_sheet_count: int = 0
    appended_record_count: int = 0
    skipped_record_count: int = 0
    unmatched_tckns: list[str] = field(default_factory=list)
    warning_messages: list[str] = field(default_factory=list)


class EducationImporter:
    """Kaynak Excel'deki mezuniyet kayıtlarını tutanak dosyasına işler."""

    def __init__(self) -> None:
        self._last_warning_messages: list[str] = []

    def import_education(
        self,
        source_path: str | Path,
        target_path: str | Path,
    ) -> EducationImportResult:
        """Kaynak mezuniyet dosyasını hedef tutanak workbook'una aktarır."""
        source_path = Path(source_path)
        target_path = Path(target_path)
        self._last_warning_messages = []

        records_by_tckn = self._read_source_records(source_path)
        if not records_by_tckn:
            raise ValueError(
                "Kaynak dosyada işlenecek geçerli mezuniyet kaydı bulunamadı."
            )

        backup_path = self._create_backup(target_path)
        workbook = openpyxl.load_workbook(target_path)
        result = EducationImportResult(backup_path=backup_path)
        matched_tckns: set[str] = set()

        try:
            for worksheet in workbook.worksheets:
                tckn = self._match_tckn(worksheet.title, records_by_tckn)
                if tckn is None:
                    continue

                matched_tckns.add(tckn)
                result.matched_sheet_count += 1

                appended_count, skipped_count, warning_messages = (
                    self._apply_records_to_sheet(
                        worksheet,
                        records_by_tckn[tckn],
                    )
                )
                result.appended_record_count += appended_count
                result.skipped_record_count += skipped_count
                self._last_warning_messages.extend(warning_messages)
                if appended_count:
                    result.updated_sheet_count += 1

            result.unmatched_tckns = sorted(
                tckn for tckn in records_by_tckn if tckn not in matched_tckns
            )
            self._last_warning_messages.extend(
                self._build_unmatched_messages(records_by_tckn, result.unmatched_tckns)
            )
            result.warning_messages = list(self._last_warning_messages)

            if result.appended_record_count:
                self._save_workbook(workbook, target_path)
        finally:
            workbook.close()

        return result

    def last_warning_messages(self) -> list[str]:
        """Son içe aktarma denemesindeki uyarıları döndürür."""
        return list(self._last_warning_messages)

    def _read_source_records(
        self,
        source_path: Path,
    ) -> dict[str, list[EducationRecord]]:
        """Kaynak Excel'i okuyup TCKN bazlı sözlüğe dönüştürür."""
        if not source_path.is_file():
            raise FileNotFoundError(
                f"Kaynak mezuniyet dosyası bulunamadı: {source_path}"
            )

        dataframe = pd.read_excel(source_path, dtype=str)
        self._validate_source_columns(dataframe)

        records_by_tckn: dict[str, list[EducationRecord]] = {}
        for index, row in dataframe.iterrows():
            record, warning_message = self._row_to_record(row, int(index) + 2)
            if record is None:
                if warning_message:
                    self._last_warning_messages.append(warning_message)
                continue
            records_by_tckn.setdefault(record.tckn, []).append(record)

        for tckn, records in records_by_tckn.items():
            records_by_tckn[tckn] = sorted(
                records,
                key=self._record_sort_key,
            )

        return records_by_tckn

    @staticmethod
    def _validate_source_columns(dataframe: pd.DataFrame) -> None:
        """Kaynak dosyada zorunlu sütunların bulunduğunu doğrular."""
        required_columns = {
            _SOURCE_COL_TCKN,
            _SOURCE_COL_AD,
            _SOURCE_COL_MEZUNIYET_TARIHI,
            _SOURCE_COL_UNIVERSITE,
            _SOURCE_COL_FAKULTE,
            _SOURCE_COL_PROGRAM,
        }
        missing_columns = required_columns - set(dataframe.columns)
        if missing_columns:
            raise ValueError(
                f"Kaynak mezuniyet dosyasında zorunlu sütunlar eksik: {sorted(missing_columns)}"
            )

    def _row_to_record(
        self,
        row: pd.Series,
        excel_row_no: int,
    ) -> tuple[EducationRecord | None, str | None]:
        """Tek bir kaynak satırını işleyip geçerliyse EducationRecord döndürür."""
        raw_tckn = self._clean_text(row.get(_SOURCE_COL_TCKN))
        raw_name = self._clean_text(row.get(_SOURCE_COL_AD))
        university = self._clean_text(row.get(_SOURCE_COL_UNIVERSITE))
        faculty = self._clean_text(row.get(_SOURCE_COL_FAKULTE))
        program = self._clean_text(row.get(_SOURCE_COL_PROGRAM))
        raw_graduation_date = self._clean_text(row.get(_SOURCE_COL_MEZUNIYET_TARIHI))

        if not any(
            [raw_tckn, raw_name, university, faculty, program, raw_graduation_date]
        ):
            return None, None

        if not raw_tckn:
            return None, self._build_source_warning_message(
                excel_row_no,
                "TC KIMLIK NO boş",
                raw_tckn=raw_tckn,
                raw_name=raw_name,
                university=university,
                program=program,
            )

        if not raw_name:
            return None, self._build_source_warning_message(
                excel_row_no,
                "AD alanı boş",
                raw_tckn=raw_tckn,
                raw_name=raw_name,
                university=university,
                program=program,
            )

        if _MISSING_RECORD_MARKER in raw_name.upper():
            return None, self._build_source_warning_message(
                excel_row_no,
                "Kayıtta mezun kaydı bulunamadı ibaresi var",
                raw_tckn=raw_tckn,
                raw_name=raw_name,
                university=university,
                program=program,
            )

        tckn = normalize_tckn(raw_tckn)
        if not validate_tckn(tckn):
            return None, self._build_source_warning_message(
                excel_row_no,
                f"Geçersiz TCKN: {tckn}",
                raw_tckn=tckn,
                raw_name=raw_name,
                university=university,
                program=program,
            )

        graduation_date = self._format_date(row.get(_SOURCE_COL_MEZUNIYET_TARIHI))

        school_text = university
        department_text = self._build_department_text(program)
        if not school_text:
            return None, self._build_source_warning_message(
                excel_row_no,
                "Okul bilgisi üretilemedi",
                raw_tckn=tckn,
                raw_name=raw_name,
                university=university,
                program=program,
            )

        return (
            EducationRecord(
                tckn=tckn,
                level=self._infer_level(program, faculty),
                school_text=school_text,
                department_text=department_text,
                graduation_date=graduation_date,
            ),
            None,
        )

    @staticmethod
    def _record_sort_key(record: EducationRecord) -> tuple[int, str, str]:
        """Kayıtları şablondaki sıra ile uyumlu dizmek için anahtar üretir."""
        return (
            _EDUCATION_LEVEL_PRIORITY.get(record.level, 99),
            record.graduation_date,
            record.school_text,
        )

    @staticmethod
    def _clean_text(value: object) -> str:
        """Hücre değerini güvenli biçimde normalize eder."""
        if value is None or pd.isna(value):
            return ""
        return " ".join(str(value).strip().split())

    def _format_date(self, value: object) -> str:
        """Mezuniyet tarihini ``dd.mm.YYYY`` biçimine çevirir."""
        text = self._clean_text(value)
        if not text:
            return ""

        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
        if pd.isna(parsed):
            return text
        return parsed.strftime("%d.%m.%Y")

    @staticmethod
    def _build_department_text(program: str) -> str:
        """Bölüm hücresine yazılacak metni üretir.

        Program adında " - " varsa yalnızca önceki kısım alınır;
        bu sayede "ÖRGÜN ÖĞRETİM" gibi ekler kırpılır.
        Ayrıca (YL) ve (DR) gibi ibareler temizlenir.
        """
        text = program
        if " - " in text:
            text = text.split(" - ")[0].strip()

        text = text.replace("(YL)", "").replace("(DR)", "").strip()
        return text

    @staticmethod
    def _infer_level(program: str, faculty: str) -> str:
        """Program metninden şablondaki öğrenim seviyesini tahmin eder."""
        search_text = f"{program} {faculty}".upper()
        if (
            "DOKTORA" in search_text
            or "SANATTA YETERLİK" in search_text
            or "(DR)" in search_text
        ):
            return OGRENIM_DOKTORA
        if (
            "TEZSİZ YÜKSEK LİSANS" in search_text
            or "TEZSIZ YUKSEK LISANS" in search_text
        ):
            return OGRENIM_TEZSIZ_YL
        if (
            "YÜKSEK LİSANS" in search_text
            or "YUKSEK LISANS" in search_text
            or "MASTER" in search_text
            or "(YL)" in search_text
        ):
            return OGRENIM_TEZLI_YL
        return OGRENIM_LISANS

    @staticmethod
    def _create_backup(target_path: Path) -> Path:
        """Hedef dosyanın zaman damgalı bir yedeğini alır."""
        if not target_path.is_file():
            raise FileNotFoundError(f"Hedef tutanak dosyası bulunamadı: {target_path}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        backup_name = f"{target_path.stem}_eski_{timestamp}{target_path.suffix}"
        backup_path = target_path.with_name(backup_name)
        copy2(target_path, backup_path)
        return backup_path

    @staticmethod
    def _match_tckn(
        sheet_title: str,
        records_by_tckn: dict[str, list[EducationRecord]],
    ) -> str | None:
        """Sayfa adı içinde geçen TCKN'yi bulur."""
        for tckn in records_by_tckn:
            if tckn == sheet_title or tckn in sheet_title:
                return tckn
        return None

    def _apply_records_to_sheet(
        self,
        worksheet,
        records: list[EducationRecord],
    ) -> tuple[int, int, list[str]]:
        """Kayıtları ilk boş eğitim satırlarına yazar."""
        education_rows = self._locate_education_rows(worksheet)
        empty_rows: list[int] = []
        existing_fingerprints: set[tuple[str, str, str]] = set()

        for row in education_rows:
            level = self._clean_text(worksheet[f"B{row}"].value)
            school = self._clean_text(worksheet[f"C{row}"].value)
            department = self._clean_text(worksheet[f"E{row}"].value)

            if school:
                # Normalize existing school cell by stripping any appended
                # graduation date (separated by ' - ') so duplicates are
                # detected even when a date is present in the cell.
                school_key = school.split(" - ")[0].casefold()
                existing_fingerprints.add(
                    (
                        level.casefold(),
                        school_key,
                        department.casefold(),
                    )
                )
                continue

            if level or department:
                continue

            empty_rows.append(row)

        appended_count = 0
        skipped_count = 0
        warning_messages: list[str] = []
        for record in records:
            if record.fingerprint in existing_fingerprints:
                skipped_count += 1
                warning_messages.append(
                    self._build_sheet_skip_message(
                        worksheet.title,
                        record,
                        "Hedef sayfada aynı eğitim kaydı zaten var",
                    )
                )
                continue

            if not empty_rows:
                skipped_count += 1
                warning_messages.append(
                    self._build_sheet_skip_message(
                        worksheet.title,
                        record,
                        "Boş eğitim satırı kalmadı",
                    )
                )
                continue

            row = empty_rows.pop(0)
            worksheet[f"B{row}"] = record.level
            worksheet[f"C{row}"] = record.school_text
            worksheet[f"E{row}"] = record.department_text
            if record.graduation_date:
                worksheet[f"I{row}"] = record.graduation_date
            existing_fingerprints.add(record.fingerprint)
            appended_count += 1

        return appended_count, skipped_count, warning_messages

    def _locate_education_rows(self, worksheet) -> list[int]:
        """Şablondaki eğitim satırlarını dinamik olarak bulur."""
        school_header_row = None
        experience_header_row = None

        for row in range(1, min(worksheet.max_row, 40) + 1):
            school_value = self._clean_text(worksheet[f"C{row}"].value).upper()
            left_value = self._clean_text(worksheet[f"B{row}"].value).upper()

            if school_header_row is None and school_value == "OKUL":
                school_header_row = row
            if "MESLEKİ TECRÜBELER" in left_value:
                experience_header_row = row
                break

        if (
            school_header_row is not None
            and experience_header_row is not None
            and experience_header_row > school_header_row + 1
        ):
            return list(range(school_header_row + 1, experience_header_row))

        return list(_DEFAULT_EDUCATION_ROWS)

    @staticmethod
    def _save_workbook(workbook, target_path: Path) -> None:
        """Workbook'u hedef dosya üzerine güvenli biçimde kaydeder."""
        try:
            workbook.save(target_path)
        except PermissionError as exc:
            raise PermissionError(
                f"Hedef tutanak dosyası kaydedilemedi. Dosya açık olabilir: {target_path}"
            ) from exc

    @staticmethod
    def _build_source_warning_message(
        excel_row_no: int,
        reason: str,
        raw_tckn: str,
        raw_name: str,
        university: str,
        program: str,
    ) -> str:
        """Geçersiz kaynak satırı için kullanıcıya dönük log mesajı üretir."""
        return (
            f"Kaynak satır {excel_row_no} atlandı: {reason}. "
            f"TCKN='{raw_tckn or '-'}', "
            f"AD='{raw_name or '-'}', "
            f"ÜNİVERSİTE='{university or '-'}', "
            f"PROGRAM='{program or '-'}'"
        )

    @staticmethod
    def _build_sheet_skip_message(
        sheet_title: str,
        record: EducationRecord,
        reason: str,
    ) -> str:
        """Hedef sayfaya yazılamayan eğitim kaydı için log mesajı üretir."""
        return (
            f"'{sheet_title}' sayfasında kayıt atlandı: {reason}. "
            f"TCKN='{record.tckn}', "
            f"SEVİYE='{record.level}', "
            f"OKUL='{record.school_text}', "
            f"BÖLÜM='{record.department_text or '-'}'"
        )

    @staticmethod
    def _build_unmatched_messages(
        records_by_tckn: dict[str, list[EducationRecord]],
        unmatched_tckns: list[str],
    ) -> list[str]:
        """Hedefte karşılığı bulunamayan TCKN'ler için log mesajı üretir."""
        messages: list[str] = []
        for tckn in unmatched_tckns:
            messages.append(
                f"Hedefte eşleşen sayfa bulunamadı: TCKN='{tckn}' "
                f"({len(records_by_tckn.get(tckn, []))} kayıt)"
            )
        return messages
