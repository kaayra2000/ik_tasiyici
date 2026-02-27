import os
import sys
from pathlib import Path

from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QCheckBox,
    QTextEdit,
    QFileDialog,
    QMessageBox,
)

from src.core.excel_reader import oku_personel_listesi
from src.core.excel_writer import olustur_dk_dosyasi


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DK Tutanak Oluşturucu")
        self.setMinimumSize(600, 400)

        # QSettings nesnesi oluştur
        self.settings = QSettings("IK_Tasiyici", "TutanakOlusturucu")

        # Arayüz elemanlarını oluştur
        self._init_ui()

        # Ayarları yükle
        self._load_settings()

    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # -- Girdi Dosyası Seçimi --
        input_layout = QHBoxLayout()
        self.input_label = QLabel("Kaynak Excel:")
        self.input_line_edit = QLineEdit()
        self.input_line_edit.setReadOnly(True)
        self.input_button = QPushButton("Girdi Seç")
        self.input_button.clicked.connect(self._select_input_file)

        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_line_edit)
        input_layout.addWidget(self.input_button)
        main_layout.addLayout(input_layout)

        # -- Çıktı Taslağı Seçimi --
        template_layout = QHBoxLayout()
        self.template_label = QLabel("Çıktı Taslağı:")
        self.template_line_edit = QLineEdit()
        self.template_line_edit.setReadOnly(True)
        self.template_button = QPushButton("Taslak Seç")
        self.template_button.clicked.connect(self._select_template_file)

        # Varsayılan Taslak Checkbox
        self.default_template_checkbox = QCheckBox("Varsayılan taslağı kullan")
        self.default_template_checkbox.stateChanged.connect(self._toggle_template_selection)

        template_layout.addWidget(self.template_label)
        template_layout.addWidget(self.template_line_edit)
        template_layout.addWidget(self.template_button)

        main_layout.addWidget(self.default_template_checkbox)
        main_layout.addLayout(template_layout)

        # -- Log Alanı --
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        main_layout.addWidget(QLabel("İşlem Logları:"))
        main_layout.addWidget(self.log_text_edit)

        # -- Başlat Butonu --
        self.start_button = QPushButton("Tutanakları Oluştur")
        self.start_button.setMinimumHeight(40)
        self.start_button.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.start_button.clicked.connect(self._start_processing)
        main_layout.addWidget(self.start_button)

    def _load_settings(self):
        """Uygulama açılırken son kaydedilen taslak yolunu yükler."""
        last_template_path = self.settings.value("last_template_path", "")
        
        # Yol var ve dosya mevcutsa kullanıcının en son seçtiği özel taslağı getir
        if last_template_path and Path(last_template_path).is_file():
            self.template_line_edit.setText(last_template_path)
            self.default_template_checkbox.setChecked(False)
            self._toggle_template_selection(Qt.CheckState.Unchecked.value)
            self.log("Son kullanılan çıktı taslağı yüklendi.")
        else:
            # Dosya yoksa veya hiç seçilmemişse varsayılanı işaretle
            self.default_template_checkbox.setChecked(True)
            self.template_line_edit.clear()
            self._toggle_template_selection(Qt.CheckState.Checked.value)

    def _toggle_template_selection(self, state):
        """Varsayılan taslak işaretlendiğinde özel taslak seçimini deaktif eder."""
        is_checked = self.default_template_checkbox.isChecked()
        self.template_button.setDisabled(is_checked)
        
        if is_checked:
            self.template_line_edit.setStyleSheet("color: gray;")
        else:
            self.template_line_edit.setStyleSheet("")

    def _select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Kaynak (Girdi) Excel Dosyasını Seç",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.input_line_edit.setText(file_path)
            self.log(f"Girdi dosyası seçildi: {file_path}")

    def _select_template_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Çıktı Taslağı Excel Dosyasını Seç",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.template_line_edit.setText(file_path)
            # Ayarı kaydet
            self.settings.setValue("last_template_path", file_path)
            self.log(f"Özel çıktı taslağı seçildi ve kaydedildi: {file_path}")

    def log(self, message: str):
        """Log alanına mesaj yazar."""
        self.log_text_edit.append(message)
        # Otomatik scroll
        vertical_scrollbar = self.log_text_edit.verticalScrollBar()
        vertical_scrollbar.setValue(vertical_scrollbar.maximum())

    def _start_processing(self):
        input_file = self.input_line_edit.text().strip()
        if not input_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir girdi dosyası seçin.")
            return

        is_default_template = self.default_template_checkbox.isChecked()
        template_file = self.template_line_edit.text().strip()

        if not is_default_template and not template_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen özel bir çıktı taslağı seçin veya 'Varsayılan taslağı kullan' seçeneğini işaretleyin.")
            return
            
        if not is_default_template and not Path(template_file).is_file():
            QMessageBox.warning(self, "Hata", "Seçilen özel taslak dosyası bulunamıyor. Lütfen geçerli bir dosya seçin.")
            return

        self.log("-" * 40)
        self.log("İşlem başlatılıyor...")
        
        try:
            self.log("Personel listesi okunuyor...")
            personeller = oku_personel_listesi(input_file)
            self.log(f"Başarılı: {len(personeller)} personel okundu.")
            
            if not personeller:
                self.log("Uyarı: İşlenecek personel bulunamadı.")
                QMessageBox.information(self, "Bilgi", "İşlenecek geçerli personel kaydı bulunamadı.\n(Birim veya isim boş olabilir)")
                return

            self.log("DK tutanakları oluşturuluyor...")
            
            # Parametreleri hazırla
            template_param = None if is_default_template else template_file
            
            # Ana dizini çıktı dizini olarak alalım (girdi dosyasının olduğu yere de kaydedilebilir)
            output_dir = Path(input_file).parent
            
            output_path = olustur_dk_dosyasi(
                personeller=personeller,
                cikti_dizini=output_dir,
                template_path=template_param
            )
            
            self.log(f"İşlem tamamlandı!\nÇıktı dosyası: {output_path}")
            QMessageBox.information(self, "Başarılı", f"İşlem tamamlandı!\nDosya kaydedildi:\n{output_path}")

        except Exception as e:
            self.log(f"HATA: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Beklenmeyen bir hata oluştu:\n{str(e)}")

