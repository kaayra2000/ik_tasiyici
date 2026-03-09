"""
Ana pencere modülü.

Uygulamanın ana penceresini tanımlayan ince orkestratör sınıfını içerir.
Tüm sorumluluklar ilgili bileşenlere devredilmiştir (SRP).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from PyQt6.QtCore import QThread, QUrl, pyqtSignal
import os
from PyQt6.QtGui import QAction, QActionGroup, QDesktopServices
from PyQt6.QtWidgets import (
    QMainWindow,
    QMenu,
    QMenuBar,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.config.constants import DEFAULT_VERSION, SUPPORTED_VERSIONS, make_tubitak_title

from src.gui.file_selection_widget import DialogType, FileSelectionWidget
from src.gui.log_widget import LogWidget
from src.gui.settings_manager import SettingsManager
from src.gui.tutanak_service import TutanakService


# ---------------------------------------------------------------------------
# Arka-plan iş parçacıkları (QThread tabanlı worker'lar)
# ---------------------------------------------------------------------------


class _PersonelOkuWorker(QThread):
    """Personel okuma işlemini arka planda yürütür."""

    finished = pyqtSignal(list)   # List[Personel]
    error = pyqtSignal(str)

    def __init__(self, service: TutanakService, input_path: str) -> None:
        super().__init__()
        self._service = service
        self._input_path = input_path

    def run(self) -> None:  # noqa: D102
        try:
            personeller = self._service.personel_oku(self._input_path)
            self.finished.emit(personeller)
        except Exception as exc:  # noqa: BLE001
            self.error.emit(str(exc))


class _TutanakOlusturWorker(QThread):
    """Tutanak oluşturma işlemini arka planda yürütür."""

    finished = pyqtSignal(object)   # Path
    error = pyqtSignal(str)

    def __init__(
        self,
        service: TutanakService,
        personeller: list,
        template_path: str,
        output_path: str,
        version: str,
    ) -> None:
        super().__init__()
        self._service = service
        self._personeller = personeller
        self._template_path = template_path
        self._output_path = output_path
        self._version = version

    def run(self) -> None:  # noqa: D102
        try:
            result_path = self._service.tutanak_olustur(
                personeller=self._personeller,
                template_path=self._template_path,
                output_path=self._output_path,
                version=self._version,
            )
            self.finished.emit(result_path)
        except Exception as exc:  # noqa: BLE001
            self.error.emit(str(exc))


@dataclass
class TutanakProcessState:
    """Tutanak oluşturma sürecinin geçici durum verilerini tutar."""
    valid_personnel_count: int = 0
    selected_version: str | None = None
    personeller: list = field(default_factory=list)
    personel_details_logged: bool = False
    pending_template_file: str = ""
    pending_output_path: str = ""


class TutanakWindow(QMainWindow):
    """DK Tutanak Oluşturucu ana penceresi.

    Sorumluluklar ilgili bileşenlere devredilmiştir:

    * **Dosya seçimi** → :class:`FileSelectionWidget`
    * **Ayar yönetimi** → :class:`SettingsManager`
    * **Log gösterimi** → :class:`LogWidget`
    * **İş mantığı** → :class:`TutanakService`
    """

    def __init__(
        self,
        settings: SettingsManager | None = None,
        service: TutanakService | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle(make_tubitak_title("DK Tutanak Oluşturucu"))
        self.setMinimumSize(600, 400)

        # Bağımlılıkları dışarıdan al ya da varsayılanları oluştur (DIP)
        self._settings = settings or SettingsManager()
        self._service = service or TutanakService()

        # Arka plan worker referansları (GC'den korumak için)
        self._okuma_worker: _PersonelOkuWorker | None = None
        self._olusturma_worker: _TutanakOlusturWorker | None = None

        # İşlem sırasında kullanılacak geçici durum nesnesi
        self._process_state = TutanakProcessState()

        self._init_menu_bar()
        self._init_ui()
        self._load_settings()

    # ------------------------------------------------------------------
    # Menü çubuğu
    # ------------------------------------------------------------------

    def _init_menu_bar(self) -> None:
        """Üst menü çubuğunu oluşturur.

        Her alt menü kendi oluşturma metoduna delege edilir (SRP).
        Yeni bir menü eklemek için sadece yeni bir metod yazılıp
        buraya eklenmesi yeterlidir (OCP).
        """
        menu_bar: QMenuBar = self.menuBar()
        ayarlar_menu: QMenu = menu_bar.addMenu("Ayarlar")

        self._init_version_menu(ayarlar_menu)

    def _init_version_menu(self, parent_menu: QMenu) -> None:
        """Çıktı sürümü alt menüsünü oluşturur.

        :param parent_menu: Alt menünün ekleneceği üst menü.
        """
        version_menu = QMenu("Çıktı Sürümü", self)
        parent_menu.addMenu(version_menu)

        self._version_action_group = QActionGroup(self)
        self._version_action_group.setExclusive(True)
        self._version_actions: dict[str, QAction] = {}

        for key, label in SUPPORTED_VERSIONS.items():
            action = QAction(label, self, checkable=True)
            action.setData(key)
            self._version_action_group.addAction(action)
            version_menu.addAction(action)
            self._version_actions[key] = action

        default_action = self._version_actions.get(DEFAULT_VERSION)
        if default_action:
            default_action.setChecked(True)

        self._version_action_group.triggered.connect(self._on_version_changed)

    def _get_selected_version(self) -> str:
        """Menüden seçili çıktı versiyonunu döndürür."""
        checked = self._version_action_group.checkedAction()
        if checked:
            return checked.data()
        return DEFAULT_VERSION

    # ------------------------------------------------------------------
    # UI oluşturma
    # ------------------------------------------------------------------

    def _init_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # -- Girdi Dosyası Seçimi --
        self._input_selector = FileSelectionWidget(
            label_text="Kaynak Excel:",
            button_text="Girdi Seç",
            dialog_title="Kaynak (Girdi) Excel Dosyasını Seç",
            dialog_type=DialogType.OPEN,
            file_filter="Excel Files (*.xlsx *.xls)",
        )
        self._input_selector.file_selected.connect(self._on_input_selected)
        main_layout.addWidget(self._input_selector)

        # -- Çıktı Taslağı Seçimi --
        self._template_selector = FileSelectionWidget(
            label_text="Çıktı Taslağı:",
            button_text="Taslak Seç",
            dialog_title="Çıktı Taslağı Excel Dosyasını Seç",
            dialog_type=DialogType.OPEN,
            file_filter="Excel Files (*.xlsx *.xls)",
        )
        self._template_selector.file_selected.connect(
            self._on_template_selected
        )
        main_layout.addWidget(self._template_selector)

        # -- Çıktı Kayıt Yeri Seçimi --
        self._output_selector = FileSelectionWidget(
            label_text="Kayıt Yeri:",
            button_text="Kayıt Yeri Seç",
            dialog_title="Çıktı Dosyasını Kaydet",
            dialog_type=DialogType.SAVE,
            file_filter="Excel Files (*.xlsx)",
        )
        self._output_selector.file_selected.connect(self._on_output_selected)
        main_layout.addWidget(self._output_selector)

        # -- Log Alanı --
        self._log_widget = LogWidget(log_name="ana_pencere")
        main_layout.addWidget(self._log_widget)

        # -- İlerleme Çubuğu --
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)   # belirsiz mod (pulse)
        self._progress_bar.setTextVisible(False)
        self._progress_bar.setFixedHeight(16)
        self._progress_bar.setVisible(False)
        main_layout.addWidget(self._progress_bar)

        # -- Başlat Butonu --
        self._start_button = QPushButton("Tutanakları Oluştur")
        self._start_button.setObjectName("startButton")
        self._start_button.setMinimumHeight(40)
        self._start_button.setStyleSheet(
            "font-weight: bold; font-size: 14px;"
        )
        self._start_button.clicked.connect(self._start_processing)
        main_layout.addWidget(self._start_button)

    # ------------------------------------------------------------------
    # Ayar yönetimi
    # ------------------------------------------------------------------

    def _load_settings(self) -> None:
        """Son kaydedilen yolları yükler."""
        self._restore_open_selector(
            self._input_selector,
            SettingsManager.KEY_INPUT_PATH,
            "Son kullanılan girdi dosyası yüklendi.",
            "Son kullanılan girdi klasörü yüklendi.",
        )
        self._restore_open_selector(
            self._template_selector,
            SettingsManager.KEY_TEMPLATE_PATH,
            "Son kullanılan özel çıktı taslağı yüklendi.",
            "Son kullanılan taslak klasörü yüklendi.",
        )

        output_path = self._settings.get(SettingsManager.KEY_OUTPUT_PATH)
        if output_path:
            self._output_selector.set_path(output_path)
            self.log("Son kullanılan kayıt yeri yüklendi.")

        saved_version = self._settings.get(
            SettingsManager.KEY_OUTPUT_VERSION, DEFAULT_VERSION
        )
        action = self._version_actions.get(saved_version)
        if action:
            action.setChecked(True)

    def _restore_open_selector(
        self,
        selector: FileSelectionWidget,
        key: str,
        success_message: str,
        fallback_message: str,
    ) -> None:
        """Açma tipi seçiciler için dosya ya da son klasörü geri yükler."""
        file_path = self._settings.get_existing_file(key)
        if file_path:
            selector.set_path(file_path)
            self.log(success_message)
            return

        parent_dir = self._settings.get_parent_dir(key)
        if parent_dir:
            selector.set_dialog_path(parent_dir)
            self.log(fallback_message)

    # ------------------------------------------------------------------
    # Sinyal slotları
    # ------------------------------------------------------------------

    def _on_input_selected(self, path: str) -> None:
        self._settings.set(SettingsManager.KEY_INPUT_PATH, path)
        self.log(f"Girdi dosyası seçildi: {path}")

    def _on_template_selected(self, path: str) -> None:
        self._settings.set(SettingsManager.KEY_TEMPLATE_PATH, path)
        self.log(f"Özel çıktı taslağı seçildi ve kaydedildi: {path}")

    def _on_output_selected(self, path: str) -> None:
        self._settings.set(SettingsManager.KEY_OUTPUT_PATH, path)
        self.log(f"Çıktı kayıt yeri belirlendi: {path}")

    def _on_version_changed(self, action: QAction) -> None:
        version = action.data()
        if version:
            self._settings.set(SettingsManager.KEY_OUTPUT_VERSION, version)
            self.log(f"Çıktı versiyonu değiştirildi: {SUPPORTED_VERSIONS.get(version, version)}")

    # ------------------------------------------------------------------
    # Log delegasyonu
    # ------------------------------------------------------------------

    def log(self, message: str) -> None:
        """Log alanına mesaj yazar.

        :param message: Yazılacak mesaj.
        """
        self._log_widget.log(message)

    def _get_service_messages(self, getter_name: str) -> list[str]:
        """Servisten güvenli biçimde string mesaj listesi alır."""
        getter = getattr(self._service, getter_name, None)
        if not callable(getter):
            return []

        messages = getter()
        if not isinstance(messages, list):
            return []
        return [message for message in messages if isinstance(message, str)]

    def _get_service_report(self, getter_name: str):
        """Servisten opsiyonel rapor nesnesi alır."""
        getter = getattr(self._service, getter_name, None)
        if not callable(getter):
            return None
        return getter()

    def _log_processing_summary(
        self,
        *,
        status: str,
        valid_personnel_count: int,
        version: str | None = None,
        result_path: str | Path | None = None,
        error_message: str | None = None,
    ) -> None:
        """İşlemin son özetini log alanının en altına yazar."""
        personel_warnings = self._get_service_messages("son_personel_okuma_uyarilari")
        tutanak_warnings = self._get_service_messages("son_tutanak_olusturma_uyarilari")
        tutanak_report = self._get_service_report("son_tutanak_olusturma_raporu")

        added_sheet_count = getattr(tutanak_report, "added_sheet_count", None)
        skipped_existing_count = getattr(tutanak_report, "skipped_existing_count", None)
        if not isinstance(added_sheet_count, int):
            added_sheet_count = None
        if not isinstance(skipped_existing_count, int):
            skipped_existing_count = len(tutanak_warnings)

        self._log_widget.log_summary_block(
            [
                ("Durum", status),
                ("Geçerli personel", valid_personnel_count),
                ("Geçersiz/atlanan kaynak satır", len(personel_warnings)),
                ("Çıktı versiyonu", version),
                ("Yeni eklenen sayfa", added_sheet_count),
                ("Mevcut olduğu için atlanan kayıt", skipped_existing_count),
                ("Çıktı dosyası", result_path),
                ("Hata", error_message),
            ]
        )

    # ------------------------------------------------------------------
    # İlerleme çubuğu yardımcıları
    # ------------------------------------------------------------------

    def _set_busy(self, busy: bool) -> None:
        """İşlem süresince butonu devre dışı bırakır ve progress bar'ı gösterir."""
        self._start_button.setEnabled(not busy)
        self._progress_bar.setVisible(busy)

    # ------------------------------------------------------------------
    # İş mantığı orkestresyonu – personel okuma aşaması
    # ------------------------------------------------------------------

    def _start_processing(self) -> None:
        """Doğrulama kontrolleri yapar ve personel okuma işlemini başlatır."""
        input_file = self._input_selector.get_path()
        if not input_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir girdi dosyası seçin.")
            return

        template_file = self._template_selector.get_path()
        if not template_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir çıktı taslağı seçin.")
            return

        if not Path(template_file).is_file():
            QMessageBox.warning(
                self,
                "Hata",
                "Seçilen taslak dosyası bulunamıyor. Lütfen geçerli bir dosya seçin.",
            )
            return

        output_path = self._output_selector.get_path()
        if not output_path:
            QMessageBox.warning(
                self, "Uyarı", "Lütfen bir çıktı kayıt yeri seçin."
            )
            return

        self.log("-" * 40)
        self.log("İşlem başlatılıyor...")
        self.log("Personel listesi okunuyor...")

        # Yeni işlem için durumu sıfırla
        self._process_state = TutanakProcessState(
            pending_template_file=template_file,
            pending_output_path=output_path,
        )

        self._set_busy(True)

        self._okuma_worker = _PersonelOkuWorker(self._service, input_file)
        self._okuma_worker.finished.connect(self._on_personel_oku_finished)
        self._okuma_worker.error.connect(self._on_personel_oku_error)
        # Run synchronously under pytest for deterministic tests.
        if os.environ.get("PYTEST_CURRENT_TEST"):
            self._okuma_worker.run()
        else:
            self._okuma_worker.start()

    def _on_personel_oku_finished(self, personeller: list) -> None:
        """Personel okuma başarıyla tamamlandığında çağrılır."""
        self._process_state.valid_personnel_count = len(personeller)
        self._log_widget.log_detail_block(
            "Personel okuma ayrıntıları:",
            self._get_service_messages("son_personel_okuma_uyarilari"),
        )
        self._process_state.personel_details_logged = True

        if not personeller:
            self._set_busy(False)
            self.log("Uyarı: İşlenecek personel bulunamadı.")
            self._log_processing_summary(
                status="İşlem yapılmadı",
                valid_personnel_count=0,
                error_message="İşlenecek geçerli personel kaydı bulunamadı.",
            )
            QMessageBox.information(
                self,
                "Bilgi",
                "İşlenecek geçerli personel kaydı bulunamadı.\n"
                "Detaylar için işlem sonuçları alanına bakın.",
            )
            return

        self._process_state.personeller = personeller
        self._process_state.selected_version = self._get_selected_version()
        self.log(f"DK tutanakları oluşturuluyor (versiyon: {self._process_state.selected_version})...")

        self._olusturma_worker = _TutanakOlusturWorker(
            self._service,
            personeller,
            self._process_state.pending_template_file,
            self._process_state.pending_output_path,
            self._process_state.selected_version,
        )
        self._olusturma_worker.finished.connect(self._on_tutanak_olustur_finished)
        self._olusturma_worker.error.connect(self._on_tutanak_olustur_error)
        # Run synchronously under pytest for deterministic tests.
        if os.environ.get("PYTEST_CURRENT_TEST"):
            self._olusturma_worker.run()
        else:
            self._olusturma_worker.start()

    def _on_personel_oku_error(self, error_message: str) -> None:
        """Personel okuma hata verdiğinde çağrılır."""
        self._set_busy(False)
        self.log(f"HATA: {error_message}")
        if not self._process_state.personel_details_logged:
            self._log_widget.log_detail_block(
                "Personel okuma ayrıntıları:",
                self._get_service_messages("son_personel_okuma_uyarilari"),
            )
        self._log_widget.log_detail_block(
            "Tutanak oluşturma ayrıntıları:",
            self._get_service_messages("son_tutanak_olusturma_uyarilari"),
        )
        self._log_processing_summary(
            status="Başarısız",
            valid_personnel_count=self._process_state.valid_personnel_count,
            version=self._process_state.selected_version,
            error_message=error_message,
        )
        QMessageBox.critical(
            self,
            "Hata",
            f"Beklenmeyen bir hata oluştu:\n{error_message}",
        )

    # ------------------------------------------------------------------
    # Tutanak oluşturma aşaması
    # ------------------------------------------------------------------

    def _on_tutanak_olustur_finished(self, result_path: object) -> None:
        """Tutanak oluşturma başarıyla tamamlandığında çağrılır."""
        self._set_busy(False)
        self._log_widget.log_detail_block(
            "Tutanak oluşturma ayrıntıları:",
            self._get_service_messages("son_tutanak_olusturma_uyarilari"),
        )
        self._open_generated_output(result_path)  # type: ignore[arg-type]
        self._log_processing_summary(
            status="Başarılı",
            valid_personnel_count=self._process_state.valid_personnel_count,
            version=self._process_state.selected_version,
            result_path=result_path,  # type: ignore[arg-type]
        )
        QMessageBox.information(
            self,
            "Başarılı",
            f"İşlem tamamlandı!\nDosya kaydedildi:\n{result_path}",
        )

    def _on_tutanak_olustur_error(self, error_message: str) -> None:
        """Tutanak oluşturma hata verdiğinde çağrılır."""
        self._set_busy(False)
        self.log(f"HATA: {error_message}")
        self._log_widget.log_detail_block(
            "Tutanak oluşturma ayrıntıları:",
            self._get_service_messages("son_tutanak_olusturma_uyarilari"),
        )
        self._log_processing_summary(
            status="Başarısız",
            valid_personnel_count=self._process_state.valid_personnel_count,
            version=self._process_state.selected_version,
            error_message=error_message,
        )
        QMessageBox.critical(
            self,
            "Hata",
            f"Beklenmeyen bir hata oluştu:\n{error_message}",
        )

    def _open_generated_output(self, result_path: str | Path) -> None:
        """Oluşturulan dosyayı ve bulunduğu klasörü açar."""
        output_file = Path(result_path).resolve()
        output_dir = output_file.parent

        self._open_local_path(output_dir, "çıktı klasörü")
        self._open_local_path(output_file, "oluşturulan tutanak")

    def _open_local_path(self, path: Path, description: str) -> bool:
        """Yerel bir dosya ya da klasörü işletim sistemine devrederek açar."""
        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        if opened:
            self.log(f"{description.capitalize()} açıldı: {path}")
        else:
            self.log(f"Uyarı: {description} otomatik açılamadı: {path}")
        return opened
