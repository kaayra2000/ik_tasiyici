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

        template_layout.addWidget(self.template_label)
        template_layout.addWidget(self.template_line_edit)
        template_layout.addWidget(self.template_button)

        main_layout.addLayout(template_layout)

        # -- Çıktı Kayıt Yeri Seçimi --
        output_layout = QHBoxLayout()
        self.output_label = QLabel("Kayıt Yeri:")
        self.output_line_edit = QLineEdit()
        self.output_line_edit.setReadOnly(True)
        self.output_button = QPushButton("Kayıt Yeri Seç")
        self.output_button.clicked.connect(self._select_output_file)
        
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_line_edit)
        output_layout.addWidget(self.output_button)
        main_layout.addLayout(output_layout)

        # -- Log Alanı --
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        main_layout.addWidget(QLabel("İşlem Sonuçları:"))
        main_layout.addWidget(self.log_text_edit)

        # -- Başlat Butonu --
        self.start_button = QPushButton("Tutanakları Oluştur")
        self.start_button.setMinimumHeight(40)
        self.start_button.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.start_button.clicked.connect(self._start_processing)
        main_layout.addWidget(self.start_button)

    def _load_settings(self):
        """Uygulama açılırken son kaydedilen yolları yükler."""
        last_input_path = self.settings.value("last_input_path", "")
        last_template_path = self.settings.value("last_template_path", "")
        last_output_path = self.settings.value("last_output_path", "")
        
        # Girdi Dosyası
        if last_input_path and Path(last_input_path).is_file():
            self.input_line_edit.setText(last_input_path)
            self.log("Son kullanılan girdi dosyası yüklendi.")
        
        # Çıktı Taslağı
        if last_template_path and Path(last_template_path).is_file():
            self.template_line_edit.setText(last_template_path)
            self.log("Son kullanılan özel çıktı taslağı yüklendi.")
            
        # Kayıt Yeri
        if last_output_path:
            self.output_line_edit.setText(last_output_path)
            self.log("Son kullanılan kayıt yeri yüklendi.")

    def _select_input_file(self):
        current_path = self.input_line_edit.text()
        if not current_path:
            last_path = self.settings.value("last_input_path", "")
            if last_path:
                current_path = str(Path(last_path).parent)
            else:
                current_path = ""

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Kaynak (Girdi) Excel Dosyasını Seç",
            current_path,
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.input_line_edit.setText(file_path)
            self.settings.setValue("last_input_path", file_path)
            self.log(f"Girdi dosyası seçildi: {file_path}")

    def _select_template_file(self):
        current_path = self.template_line_edit.text()
        if not current_path:
            last_path = self.settings.value("last_template_path", "")
            if last_path:
                current_path = str(Path(last_path).parent)
            else:
                current_path = ""
                
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Çıktı Taslağı Excel Dosyasını Seç",
            current_path,
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.template_line_edit.setText(file_path)
            # Ayarı kaydet
            self.settings.setValue("last_template_path", file_path)
            self.log(f"Özel çıktı taslağı seçildi ve kaydedildi: {file_path}")

    def _select_output_file(self):
        # Eğer önceden seçilmiş bir yol varsa oradan başla
        current_path = self.output_line_edit.text()
        if not current_path:
            last_path = self.settings.value("last_output_path", "")
            if last_path:
                current_path = last_path
            else:
                current_path = ""
                
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Çıktı Dosyasını Kaydet",
            current_path,
            "Excel Files (*.xlsx)"
        )
        if file_path:
            # .xlsx uzantısı yoksa ekle
            if not file_path.endswith(".xlsx"):
                file_path += ".xlsx"
                
            self.output_line_edit.setText(file_path)
            # Seçilen tam yolu ayarlara kaydet
            self.settings.setValue("last_output_path", file_path)
            self.log(f"Çıktı kayıt yeri belirlendi: {file_path}")

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

        template_file = self.template_line_edit.text().strip()

        if not template_file:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir çıktı taslağı seçin.")
            return
            
        if not Path(template_file).is_file():
            QMessageBox.warning(self, "Hata", "Seçilen taslak dosyası bulunamıyor. Lütfen geçerli bir dosya seçin.")
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
            template_param = template_file
            
            # Çıktı yolu kontrolü
            output_full_path = self.output_line_edit.text().strip()
            
            if not output_full_path:
                QMessageBox.warning(self, "Uyarı", "Lütfen bir çıktı kayıt yeri seçin.")
                return

            output_path_obj = Path(output_full_path)
            output_dir = output_path_obj.parent
            output_filename = output_path_obj.name
            
            output_path = olustur_dk_dosyasi(
                personeller=personeller,
                cikti_dizini=output_dir,
                dosya_adi=output_filename,
                template_path=template_param
            )
            
            self.log(f"İşlem tamamlandı!\nÇıktı dosyası: {output_path}")
            QMessageBox.information(self, "Başarılı", f"İşlem tamamlandı!\nDosya kaydedildi:\n{output_path}")

        except Exception as e:
            self.log(f"HATA: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Beklenmeyen bir hata oluştu:\n{str(e)}")

