"""Mezuniyet ve meslek bilgilerinin mevcut tutanaklara işlendiği pencere."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.core.education_importer import EducationImportResult
from src.gui.education_import_service import EducationImportService
from src.gui.file_selection_widget import DialogType, FileSelectionWidget
from src.gui.log_widget import LogWidget
from src.gui.settings_manager import SettingsManager
from src.config.constants import make_tubitak_title


# ---------------------------------------------------------------------------
# Arka-plan iş parçacığı (QThread tabanlı worker)
# ---------------------------------------------------------------------------


class _ImportWorker(QThread):
    """Mezuniyet içe aktarmayı arka planda yürütür."""

    finished = pyqtSignal(object)   # EducationImportResult
    error = pyqtSignal(str, bool)   # (mesaj, izin_hatası_mı)

    def __init__(
        self,
        service: EducationImportService,
        source_path: str,
        target_path: str,
    ) -> None:
        super().__init__()
        self._service = service
        self._source_path = source_path
        self._target_path = target_path

    def run(self) -> None:  # noqa: D102
        try:
            result = self._service.import_education(
                source_path=self._source_path,
                target_path=self._target_path,
            )
            self.finished.emit(result)
        except PermissionError as exc:  # noqa: BLE001
            self.error.emit(str(exc), True)
        except Exception as exc:  # noqa: BLE001
            self.error.emit(str(exc), False)


class EducationImportWindow(QMainWindow):
    """Mezuniyet aktarım akışının GUI orkestratörü."""

    def __init__(
        self,
        settings: SettingsManager | None = None,
        service: EducationImportService | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle(make_tubitak_title("Mezuniyet Bilgisi Ekle"))
        self.setMinimumSize(680, 420)

        self._settings = settings or SettingsManager()
        self._service = service or EducationImportService()

        # Worker referansı (GC'den korumak için)
        self._import_worker: _ImportWorker | None = None

        self._init_ui()
        self._load_settings()

    def _init_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self._target_selector = FileSelectionWidget(
            label_text="Hedef Tutanak:",
            button_text="Tutanak Seç",
            dialog_title="Veri yazılacak hedef tutanak dosyasını seçin",
            dialog_type=DialogType.OPEN,
            file_filter="Excel Files (*.xlsx *.xls)",
        )
        self._target_selector.file_selected.connect(self._on_target_selected)
        main_layout.addWidget(self._target_selector)

        self._source_selector = FileSelectionWidget(
            label_text="Kaynak Veri:",
            button_text="Kaynak Seç",
            dialog_title="Mezuniyet bilgileri Excel dosyasını seçin",
            dialog_type=DialogType.OPEN,
            file_filter="Excel Files (*.xlsx *.xls)",
        )
        self._source_selector.file_selected.connect(self._on_source_selected)
        main_layout.addWidget(self._source_selector)

        self._log_widget = LogWidget(log_name="mezuniyet_aktarimi")
        main_layout.addWidget(self._log_widget)

        # -- İlerleme Çubuğu --
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)   # belirsiz mod (pulse)
        self._progress_bar.setTextVisible(False)
        self._progress_bar.setFixedHeight(16)
        self._progress_bar.setVisible(False)
        main_layout.addWidget(self._progress_bar)

        self._start_button = QPushButton("İşlemi Başlat")
        self._start_button.setObjectName("startButton")
        self._start_button.setMinimumHeight(40)
        self._start_button.clicked.connect(self._start_import)
        main_layout.addWidget(self._start_button)

    def _load_settings(self) -> None:
        """Son seçilen dosya veya klasörleri geri yükler."""
        self._restore_open_selector(
            self._target_selector,
            SettingsManager.KEY_EDUCATION_TARGET_PATH,
            "Son kullanılan hedef tutanak dosyası yüklendi.",
            "Son kullanılan hedef tutanak klasörü yüklendi.",
        )
        self._restore_open_selector(
            self._source_selector,
            SettingsManager.KEY_EDUCATION_SOURCE_PATH,
            "Son kullanılan kaynak mezuniyet dosyası yüklendi.",
            "Son kullanılan kaynak mezuniyet klasörü yüklendi.",
        )

    def _restore_open_selector(
        self,
        selector: FileSelectionWidget,
        key: str,
        success_message: str,
        fallback_message: str,
    ) -> None:
        """Seçiciye daha önce kaydedilen dosya ya da klasörü uygular."""
        file_path = self._settings.get_existing_file(key)
        if file_path:
            selector.set_path(file_path)
            self.log(success_message)
            return

        parent_dir = self._settings.get_parent_dir(key)
        if parent_dir:
            selector.set_dialog_path(parent_dir)
            self.log(fallback_message)

    def _on_target_selected(self, path: str) -> None:
        self._settings.set(SettingsManager.KEY_EDUCATION_TARGET_PATH, path)
        self.log(f"Hedef tutanak seçildi: {path}")

    def _on_source_selected(self, path: str) -> None:
        self._settings.set(SettingsManager.KEY_EDUCATION_SOURCE_PATH, path)
        self.log(f"Kaynak mezuniyet dosyası seçildi: {path}")

    def log(self, message: str) -> None:
        """Log alanına mesaj yazar."""
        self._log_widget.log(message)

    def _get_import_warnings(self) -> list[str]:
        """Servisten güvenli biçimde içe aktarma uyarıları alır."""
        warning_getter = getattr(self._service, "son_import_uyarilari", None)
        if not callable(warning_getter):
            return []

        warnings = warning_getter()
        if not isinstance(warnings, list):
            return []
        return [warning for warning in warnings if isinstance(warning, str)]

    def _log_summary(
        self,
        *,
        status: str,
        result: EducationImportResult | None = None,
        warning_count: int = 0,
        error_message: str | None = None,
    ) -> None:
        """İçe aktarma özetini log alanının en altına yazar."""
        rows: list[tuple[str, object]] = [
            ("Durum", status),
            ("Ayrıntı/uyarı kaydı", warning_count),
        ]
        if result is not None:
            rows.extend(
                [
                    ("Yedek oluşturuldu", result.backup_path),
                    ("Eşleşen sayfa sayısı", result.matched_sheet_count),
                    ("Güncellenen sayfa sayısı", result.updated_sheet_count),
                    ("Eklenen eğitim kaydı", result.appended_record_count),
                    ("Atlanan kayıt sayısı", result.skipped_record_count),
                    ("Hedefte bulunamayan TCKN sayısı", len(result.unmatched_tckns)),
                ]
            )
            if result.unmatched_tckns:
                joined_tckns = ", ".join(result.unmatched_tckns)
                rows.append(("Hedefte bulunamayan TCKN'ler", joined_tckns))
        if error_message:
            rows.append(("Hata", error_message))
        self._log_widget.log_summary_block(rows)

    # ------------------------------------------------------------------
    # İlerleme çubuğu yardımcısı
    # ------------------------------------------------------------------

    def _set_busy(self, busy: bool) -> None:
        """İşlem süresince butonu devre dışı bırakır ve progress bar'ı gösterir."""
        self._start_button.setEnabled(not busy)
        self._progress_bar.setVisible(busy)

    # ------------------------------------------------------------------
    # İş mantığı orkestresyonu
    # ------------------------------------------------------------------

    def _start_import(self) -> None:
        """Doğrulama sonrası içe aktarma sürecini başlatır."""
        target_path = self._target_selector.get_path()
        if not target_path:
            QMessageBox.warning(
                self,
                "Uyarı",
                "Lütfen veri yazılacak hedef tutanak dosyasını seçin.",
            )
            return

        source_path = self._source_selector.get_path()
        if not source_path:
            QMessageBox.warning(
                self,
                "Uyarı",
                "Lütfen mezuniyet bilgileri kaynağını seçin.",
            )
            return

        if not Path(target_path).is_file():
            QMessageBox.warning(
                self,
                "Hata",
                "Seçilen hedef tutanak dosyası bulunamıyor. Lütfen geçerli bir dosya seçin.",
            )
            return

        if not Path(source_path).is_file():
            QMessageBox.warning(
                self,
                "Hata",
                "Seçilen kaynak mezuniyet dosyası bulunamıyor. Lütfen geçerli bir dosya seçin.",
            )
            return

        self.log("-" * 40)
        self.log("Mezuniyet aktarımı başlatılıyor...")

        self._set_busy(True)

        self._import_worker = _ImportWorker(self._service, source_path, target_path)
        self._import_worker.finished.connect(self._on_import_finished)
        self._import_worker.error.connect(self._on_import_error)
        self._import_worker.start()

    def _on_import_finished(self, result: object) -> None:
        """İçe aktarma başarıyla tamamlandığında çağrılır."""
        self._set_busy(False)
        import_result: EducationImportResult = result  # type: ignore[assignment]
        warnings = self._get_import_warnings()
        self._log_widget.log_detail_block("İçe aktarma ayrıntıları:", warnings)
        self._log_summary(
            status="Başarılı",
            result=import_result,
            warning_count=len(warnings),
        )
        QMessageBox.information(
            self,
            "Başarılı",
            "Mezuniyet bilgileri hedef tutanağa işlendi.",
        )

    def _on_import_error(self, error_message: str, is_permission_error: bool) -> None:
        """İçe aktarma hata verdiğinde çağrılır."""
        self._set_busy(False)
        self.log(f"HATA: {error_message}")
        warnings = self._get_import_warnings()
        self._log_widget.log_detail_block("İçe aktarma ayrıntıları:", warnings)
        self._log_summary(
            status="Başarısız",
            warning_count=len(warnings),
            error_message=error_message,
        )
        if is_permission_error:
            QMessageBox.critical(
                self,
                "Dosya Kullanımda",
                f"{error_message}\n\nDosyayı kapatıp tekrar deneyin.",
            )
        else:
            QMessageBox.critical(
                self,
                "Hata",
                f"Beklenmeyen bir hata oluştu:\n{error_message}",
            )
