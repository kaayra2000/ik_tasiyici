"""
Ana pencere modülü.

Uygulamanın ana penceresini tanımlayan ince orkestratör sınıfını içerir.
Tüm sorumluluklar ilgili bileşenlere devredilmiştir (SRP).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.config.constants import DEFAULT_VERSION, SUPPORTED_VERSIONS

from src.gui.file_selection_widget import DialogType, FileSelectionWidget
from src.gui.log_widget import LogWidget
from src.gui.settings_manager import SettingsManager
from src.gui.tutanak_service import TutanakService


class MainWindow(QMainWindow):
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
        self.setWindowTitle("DK Tutanak Oluşturucu")
        self.setMinimumSize(600, 400)

        # Bağımlılıkları dışarıdan al ya da varsayılanları oluştur (DIP)
        self._settings = settings or SettingsManager()
        self._service = service or TutanakService()

        self._init_ui()
        self._load_settings()

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

        # -- Çıktı Versiyonu Seçimi --
        version_layout = QHBoxLayout()
        version_label = QLabel("Çıktı Versiyonu:")
        self._version_combo = QComboBox()
        self._version_combo.setObjectName("versionCombo")
        for key, label in SUPPORTED_VERSIONS.items():
            self._version_combo.addItem(label, key)
        self._version_combo.currentIndexChanged.connect(
            self._on_version_changed
        )
        version_layout.addWidget(version_label)
        version_layout.addWidget(self._version_combo, 1)
        main_layout.addLayout(version_layout)

        # -- Log Alanı --
        self._log_widget = LogWidget()
        main_layout.addWidget(self._log_widget)

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
        input_path = self._settings.get_existing_file(
            SettingsManager.KEY_INPUT_PATH
        )
        if input_path:
            self._input_selector.set_path(input_path)
            self.log("Son kullanılan girdi dosyası yüklendi.")

        template_path = self._settings.get_existing_file(
            SettingsManager.KEY_TEMPLATE_PATH
        )
        if template_path:
            self._template_selector.set_path(template_path)
            self.log("Son kullanılan özel çıktı taslağı yüklendi.")

        output_path = self._settings.get(SettingsManager.KEY_OUTPUT_PATH)
        if output_path:
            self._output_selector.set_path(output_path)
            self.log("Son kullanılan kayıt yeri yüklendi.")

        saved_version = self._settings.get(
            SettingsManager.KEY_OUTPUT_VERSION, DEFAULT_VERSION
        )
        idx = self._version_combo.findData(saved_version)
        if idx >= 0:
            self._version_combo.setCurrentIndex(idx)

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

    def _on_version_changed(self, index: int) -> None:
        version = self._version_combo.currentData()
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

    # ------------------------------------------------------------------
    # İş mantığı orkestresyonu
    # ------------------------------------------------------------------

    def _start_processing(self) -> None:
        """Doğrulama kontrolleri yapar ve tutanak oluşturma sürecini başlatır."""
        input_file = self._input_selector.get_path()
        if not input_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir girdi dosyası seçin.")
            return

        template_file = self._template_selector.get_path()
        if not template_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir çıktı taslağı seçin.")
            return

        from pathlib import Path

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

        try:
            self.log("Personel listesi okunuyor...")
            personeller = self._service.personel_oku(input_file)
            self.log(f"Başarılı: {len(personeller)} personel okundu.")

            if not personeller:
                self.log("Uyarı: İşlenecek personel bulunamadı.")
                QMessageBox.information(
                    self,
                    "Bilgi",
                    "İşlenecek geçerli personel kaydı bulunamadı.\n"
                    "(Birim veya isim boş olabilir)",
                )
                return

            selected_version = self._version_combo.currentData() or DEFAULT_VERSION
            self.log(f"DK tutanakları oluşturuluyor (versiyon: {selected_version})...")
            result_path = self._service.tutanak_olustur(
                personeller=personeller,
                template_path=template_file,
                output_path=output_path,
                version=selected_version,
            )

            self.log(f"İşlem tamamlandı!\nÇıktı dosyası: {result_path}")
            QMessageBox.information(
                self,
                "Başarılı",
                f"İşlem tamamlandı!\nDosya kaydedildi:\n{result_path}",
            )

        except Exception as e:
            self.log(f"HATA: {str(e)}")
            QMessageBox.critical(
                self,
                "Hata",
                f"Beklenmeyen bir hata oluştu:\n{str(e)}",
            )
