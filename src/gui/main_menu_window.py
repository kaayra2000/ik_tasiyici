"""Uygulama açılışındaki ana menü penceresi."""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.gui.education_import_window import EducationImportWindow
from src.gui.tutanak_window import TutanakWindow
from src.main import get_logo_path


class MainMenuWindow(QMainWindow):
    """İki ana akış arasında geçiş sağlayan giriş ekranı."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Personel Asistan")
        self.setMinimumSize(720, 480)
        self._active_window: QMainWindow | None = None
        self._init_ui()

    def _init_ui(self) -> None:
        central_widget = QWidget()
        central_widget.setObjectName("mainMenuRoot")
        self.setCentralWidget(central_widget)

        outer_layout = QVBoxLayout(central_widget)
        outer_layout.addStretch()

        menu_card = QFrame()
        menu_card.setObjectName("menuCard")
        menu_card.setMaximumWidth(720)
        card_layout = QVBoxLayout(menu_card)
        card_layout.setSpacing(18)

        # -- TÜBİTAK Logosu --
        logo_label = self._build_logo_label()
        if logo_label is not None:
            card_layout.addWidget(logo_label)

        title_label = QLabel("Personel Asistan")
        title_label.setObjectName("menuTitle")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(title_label)

        subtitle_label = QLabel(
            "İş akışını seçin ve ilgili modülle devam edin."
        )
        subtitle_label.setObjectName("menuSubtitle")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(subtitle_label)

        button_row = QHBoxLayout()
        button_row.setSpacing(16)

        self._tutanak_button = self._build_menu_button(
            "Tutanak Oluştur",
            "menuPrimaryButton",
        )
        self._tutanak_button.clicked.connect(self._open_tutanak_window)
        button_row.addWidget(self._tutanak_button)

        self._education_button = self._build_menu_button(
            "Mezuniyet/Meslek Bilgisi Ekle",
            "menuSecondaryButton",
        )
        self._education_button.clicked.connect(self._open_education_import_window)
        button_row.addWidget(self._education_button)

        card_layout.addLayout(button_row)

        outer_layout.addWidget(
            menu_card,
            alignment=Qt.AlignmentFlag.AlignHCenter,
        )
        outer_layout.addStretch()

    @staticmethod
    def _build_logo_label() -> QLabel | None:
        """TÜBİTAK logo etiketini oluşturur; dosya yoksa None döner."""
        logo_path = get_logo_path()
        if not logo_path.exists():
            return None
        label = QLabel()
        label.setObjectName("menuLogo")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pixmap = QPixmap(str(logo_path))
        label.setPixmap(
            pixmap.scaledToHeight(64, Qt.TransformationMode.SmoothTransformation)
        )
        return label

    @staticmethod
    def _build_menu_button(text: str, object_name: str) -> QPushButton:
        """Ana menüde kullanılacak büyük aksiyon butonunu üretir."""
        button = QPushButton(text)
        button.setObjectName(object_name)
        button.setMinimumHeight(96)
        return button

    def _open_tutanak_window(self) -> None:
        """Tutanak oluşturma ekranını açar."""
        self._open_child_window(TutanakWindow())

    def _open_education_import_window(self) -> None:
        """Mezuniyet içe aktarma ekranını açar."""
        self._open_child_window(EducationImportWindow())

    def _open_child_window(self, window: QMainWindow) -> None:
        """Seçilen modül penceresini açar ve menüyü gizler."""
        if self._active_window is not None:
            self._active_window.raise_()
            self._active_window.activateWindow()
            return

        self._active_window = window
        self._active_window.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self._active_window.destroyed.connect(self._on_child_closed)
        self.hide()
        self._active_window.show()

    def _on_child_closed(self) -> None:
        """Alt pencere kapanınca menüyü tekrar görünür yapar."""
        self._active_window = None
        self.show()
