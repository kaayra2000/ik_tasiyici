"""
Salt-okunur log gösterim bileşeni.

Loglama sorumluluğunu MainWindow'dan ayırarak SRP sağlar.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from PyQt6.QtWidgets import QLabel, QTextEdit, QVBoxLayout, QWidget


class LogWidget(QWidget):
    """Salt-okunur log alanı bileşeni.

    :param title: Alan başlığı.
    :param parent: Üst widget.
    """

    def __init__(
        self, title: str = "İşlem Sonuçları:", parent: QWidget | None = None
    ) -> None:
        super().__init__(parent)
        self._init_ui(title)

    def _init_ui(self, title: str) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._label = QLabel(title)
        self._text_edit = QTextEdit()
        self._text_edit.setReadOnly(True)

        layout.addWidget(self._label)
        layout.addWidget(self._text_edit)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def log(self, message: str) -> None:
        """Log alanına mesaj ekler ve otomatik aşağı kaydırır.

        :param message: Yazılacak mesaj.
        """
        self._text_edit.append(message)
        scrollbar = self._text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def log_detail_block(self, title: str, messages: Iterable[str]) -> None:
        """Detay mesajlarını başlık altında blok halinde yazar."""
        lines = [message for message in messages if isinstance(message, str) and message]
        if not lines:
            return

        self.log(title)
        for line in lines:
            self.log(line)

    def log_summary_block(
        self,
        rows: Iterable[tuple[str, Any]],
        title: str = "Özet:",
    ) -> None:
        """Özet satırlarını tek blok halinde yazar."""
        normalized_rows = [
            (label, value)
            for label, value in rows
            if isinstance(label, str) and label and value is not None
        ]

        self.log(title)
        for label, value in normalized_rows:
            self.log(f"{label}: {value}")

    def clear(self) -> None:
        """Log alanını temizler."""
        self._text_edit.clear()
