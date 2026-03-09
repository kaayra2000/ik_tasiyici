"""
Salt-okunur log gösterim bileşeni.

Loglama sorumluluğunu MainWindow'dan ayırarak SRP sağlar.
"""

from __future__ import annotations

from collections.abc import Iterable
import re
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
        self._markdown_blocks: list[str] = []
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
        stripped = message.strip()
        if not stripped:
            return

        if re.fullmatch(r"-{3,}", stripped):
            self._append_markdown_block("---")
            return

        label, content = self._split_labeled_message(message)
        if label:
            markdown = f"- **{self._escape_markdown(label)}:** {self._escape_markdown(content)}"
        else:
            markdown = f"- {self._escape_markdown(message)}"
        self._append_markdown_block(markdown)

    def log_detail_block(self, title: str, messages: Iterable[str]) -> None:
        """Detay mesajlarını başlık altında blok halinde yazar."""
        lines = [message for message in messages if isinstance(message, str) and message]
        if not lines:
            return

        heading = self._escape_markdown(title.strip())
        body = "\n".join(f"- {self._escape_markdown(line)}" for line in lines)
        self._append_markdown_block(f"### {heading}\n\n{body}")

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

        body = "\n".join(
            f"- **{self._escape_markdown(label)}:** {self._escape_markdown(str(value))}"
            for label, value in normalized_rows
        )
        self._append_markdown_block(
            f"### {self._escape_markdown(title.strip())}\n\n{body}"
        )

    def clear(self) -> None:
        """Log alanını temizler."""
        self._markdown_blocks.clear()
        self._text_edit.clear()

    def _append_markdown_block(self, markdown: str) -> None:
        """Markdown bloğunu belgeye ekleyip görünümü yeniler."""
        if not markdown.strip():
            return

        self._markdown_blocks.append(markdown.strip())
        self._text_edit.setMarkdown("\n\n".join(self._markdown_blocks))
        scrollbar = self._text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @staticmethod
    def _escape_markdown(text: str) -> str:
        """Serbest metni güvenli markdown metnine çevirir."""
        escaped = str(text).replace("\\", "\\\\")
        for char in ("`", "*", "_", "{", "}", "[", "]", "(", ")", "#", "+", "!", "|", ">"):
            escaped = escaped.replace(char, f"\\{char}")
        return escaped.replace("\n", "  \n")

    @staticmethod
    def _split_labeled_message(message: str) -> tuple[str, str]:
        """`ETİKET: içerik` biçimindeki mesajları ayırır."""
        if ":" not in message:
            return "", ""

        label, content = message.split(":", 1)
        label = label.strip()
        content = content.strip()
        if not label or not content:
            return "", ""
        if " " in label:
            return "", ""
        return label, content
