"""
Salt-okunur log gösterim bileşeni.

Loglama sorumluluğunu TutanakWindow'dan ayırarak SRP sağlar.
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
import re
from typing import Any

from PyQt6.QtCore import QStandardPaths
from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src.config.constants import APP_NAME, LOG_DIR_NAME


class LogWidget(QWidget):
    """Salt-okunur log alanı bileşeni.

    :param title: Alan başlığı.
    :param parent: Üst widget.
    """

    def __init__(
        self,
        title: str = "İşlem Sonuçları:",
        parent: QWidget | None = None,
        log_name: str | None = None,
    ) -> None:
        super().__init__(parent)
        self._markdown_blocks: list[str] = []
        self._log_file_path = self._resolve_log_file_path(log_name or title)

        if self._log_file_path:
            self._append_to_file(
                f"---\n**[Yeni Oturum - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]**"
            )

        self._init_ui(title)

    def _init_ui(self, title: str) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)

        self._label = QLabel(title)

        path_str = str(self._log_file_path.parent) if self._log_file_path else ""
        self._path_label = QLabel(
            f'<a href="file://{path_str}" style="color: gray; text-decoration: none;">[Klasör: {path_str}]</a>'
            if path_str
            else ""
        )
        self._path_label.setOpenExternalLinks(True)

        self._clear_button = QPushButton("Kayıtları Temizle")
        self._clear_button.clicked.connect(self.clear)

        top_layout.addWidget(self._label)
        top_layout.addSpacing(10)
        top_layout.addWidget(self._path_label)
        top_layout.addStretch()
        top_layout.addWidget(self._clear_button)

        self._text_edit = QTextEdit()
        self._text_edit.setReadOnly(True)

        layout.addLayout(top_layout)
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
        lines = [
            message for message in messages if isinstance(message, str) and message
        ]
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

    def log_file_path(self) -> Path | None:
        """Log dosyası yolunu döndürür."""
        return self._log_file_path

    def _append_markdown_block(self, markdown: str) -> None:
        """Markdown bloğunu belgeye ekleyip görünümü yeniler."""
        stripped = markdown.strip()
        if not stripped:
            return

        self._markdown_blocks.append(stripped)
        markdown_document = "\n\n".join(self._markdown_blocks)
        self._text_edit.setMarkdown(markdown_document)
        self._append_to_file(stripped)

        scrollbar = self._text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _append_to_file(self, content: str) -> None:
        """Yeni markdown içeriğini dosyaya ekler."""
        if self._log_file_path is None:
            return

        try:
            with open(self._log_file_path, "a", encoding="utf-8") as file:
                file.write(content + "\n\n")
        except OSError:
            # Dosya yazımı başarısız olursa UI akışını bozmayız.
            pass

    @classmethod
    def _resolve_log_file_path(cls, raw_name: str) -> Path | None:
        """Oturuma ait log dosyası yolunu çözümler."""
        log_name = cls._sanitize_log_name(raw_name)

        for location in (
            QStandardPaths.StandardLocation.AppLocalDataLocation,
            QStandardPaths.StandardLocation.AppDataLocation,
            QStandardPaths.StandardLocation.TempLocation,
        ):
            directory = cls._resolve_log_directory(location)
            if directory is None:
                continue

            try:
                directory.mkdir(parents=True, exist_ok=True)
            except OSError:
                continue

            return directory / f"{log_name}.md"

        return None

    @classmethod
    def _resolve_log_directory(
        cls,
        location: QStandardPaths.StandardLocation,
    ) -> Path | None:
        """Qt standart dizinlerinden log klasörünü türetir."""
        location_path = QStandardPaths.writableLocation(location)
        if not location_path:
            return None

        base_path = Path(location_path)
        if location == QStandardPaths.StandardLocation.TempLocation:
            return base_path / APP_NAME / LOG_DIR_NAME
        return base_path / LOG_DIR_NAME

    @staticmethod
    def _sanitize_log_name(raw_name: str) -> str:
        """Log dosyası adını platform bağımsız güvenli hale getirir."""
        normalized = re.sub(r"[^\w.-]+", "_", str(raw_name), flags=re.UNICODE)
        normalized = normalized.strip("._")
        return normalized or "log"

    @staticmethod
    def _escape_markdown(text: str) -> str:
        """Serbest metni güvenli markdown metnine çevirir."""
        escaped = str(text).replace("\\", "\\\\")
        for char in (
            "`",
            "*",
            "_",
            "{",
            "}",
            "[",
            "]",
            "(",
            ")",
            "#",
            "+",
            "!",
            "|",
            ">",
        ):
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
