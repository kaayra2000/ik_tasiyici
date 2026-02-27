"""
Yeniden kullanılabilir dosya seçim bileşeni.

Label + ReadOnly QLineEdit + QPushButton kalıbını tek bir widget'ta
kapsülleyerek DRY ve SRP ihlallerini giderir.
"""

from __future__ import annotations

from enum import Enum, auto

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QWidget,
)


class DialogType(Enum):
    """Dosya diyaloğu türü."""

    OPEN = auto()
    SAVE = auto()


class FileSelectionWidget(QWidget):
    """Label + ReadOnly LineEdit + Button dosya seçim bileşeni.

    :param label_text: Sol taraftaki etiket metni.
    :param button_text: Buton üzerindeki metin.
    :param dialog_title: Dosya diyaloğu başlığı.
    :param dialog_type: ``DialogType.OPEN`` veya ``DialogType.SAVE``.
    :param file_filter: Dosya filtresi (ör. ``"Excel Files (*.xlsx)"``).
    :param parent: Üst widget.

    Signals:
        file_selected(str): Dosya seçildiğinde yayınlanır.
    """

    file_selected = pyqtSignal(str)

    def __init__(
        self,
        label_text: str,
        button_text: str,
        dialog_title: str,
        dialog_type: DialogType = DialogType.OPEN,
        file_filter: str = "Excel Files (*.xlsx *.xls)",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._dialog_title = dialog_title
        self._dialog_type = dialog_type
        self._file_filter = file_filter

        self._init_ui(label_text, button_text)

    # ------------------------------------------------------------------
    # UI oluşturma
    # ------------------------------------------------------------------

    def _init_ui(self, label_text: str, button_text: str) -> None:
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._label = QLabel(label_text)
        self._line_edit = QLineEdit()
        self._line_edit.setReadOnly(True)
        self._button = QPushButton(button_text)
        self._button.setObjectName("actionButton")
        self._button.clicked.connect(self._open_dialog)

        layout.addWidget(self._label)
        layout.addWidget(self._line_edit)
        layout.addWidget(self._button)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_path(self) -> str:
        """Seçilen dosya yolunu döner."""
        return self._line_edit.text().strip()

    def set_path(self, path: str) -> None:
        """Dosya yolunu programatik olarak ayarlar.

        :param path: Gösterilecek dosya yolu.
        """
        self._line_edit.setText(path)

    # ------------------------------------------------------------------
    # Slot (orkestratör — SRP: yalnızca koordinasyon)
    # ------------------------------------------------------------------

    def _open_dialog(self) -> None:
        """Diyaloğu oluşturur, yapılandırır, konumlandırır ve sonucu işler."""
        dialog = self._build_dialog()
        self._configure_dialog(dialog)
        self._center_dialog(dialog)
        self._handle_result(dialog)

    # ------------------------------------------------------------------
    # Diyalog yardımcıları (her metot tek bir sorumluluğa sahip — SRP)
    # ------------------------------------------------------------------

    def _build_dialog(self) -> QFileDialog:
        """Temel QFileDialog nesnesini oluşturur.

        Native diyalog kapatılır; böylece pozisyonlandırma mümkün olur.
        """
        dialog = QFileDialog(
            self,
            self._dialog_title,
            self.get_path(),
            self._file_filter,
        )
        dialog.setOption(QFileDialog.Option.DontUseNativeDialog)
        return dialog

    def _configure_dialog(self, dialog: QFileDialog) -> None:
        """Diyaloğu türüne göre yapılandırır (OCP: yeni tür için yalnızca burası genişler).

        :param dialog: Yapılandırılacak diyalog nesnesi.
        """
        _MODE_MAP = {
            DialogType.OPEN: QFileDialog.AcceptMode.AcceptOpen,
            DialogType.SAVE: QFileDialog.AcceptMode.AcceptSave,
        }
        dialog.setAcceptMode(_MODE_MAP[self._dialog_type])
        if self._dialog_type == DialogType.SAVE:
            dialog.setDefaultSuffix("xlsx")

    def _center_dialog(self, dialog: QFileDialog) -> None:
        """Diyaloğu mevcut ekranın tam merkezine konumlandırır.

        :param dialog: Konumlandırılacak diyalog nesnesi.
        """
        if not self.screen():
            return
        qr = dialog.frameGeometry()
        qr.moveCenter(self.screen().availableGeometry().center())
        dialog.move(qr.topLeft())

    def _handle_result(self, dialog: QFileDialog) -> None:
        """Kullanıcının diyalogu kabul etmesi durumunda yolu günceller.

        :param dialog: Sonucu işlenecek diyalog nesnesi.
        """
        if dialog.exec() != QFileDialog.DialogCode.Accepted:
            return
        selected = dialog.selectedFiles()
        if selected:
            self._line_edit.setText(selected[0])
            self.file_selected.emit(selected[0])
