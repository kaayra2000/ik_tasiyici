"""LogWidget birim testleri."""

import os
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.log_widget import LogWidget


@pytest.fixture(scope="session")
def qapp():
    """Test oturumu için tek bir QApplication örneği sağlar."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture()
def widget(qapp):
    """Test için LogWidget örneği üretir."""
    log_widget = LogWidget()
    yield log_widget
    log_widget.close()


class TestLogWidget:
    """LogWidget blok loglama davranış testleri."""

    def test_log_renders_plain_messages_as_markdown_list_item(
        self,
        widget: LogWidget,
    ):
        """Tekil log satırı markdown madde işareti olarak yazılmalı."""
        widget.log("İşlem başlatılıyor...")

        markdown = widget._text_edit.toMarkdown()
        assert "- İşlem başlatılıyor..." in markdown
        assert "İşlem başlatılıyor..." in widget._text_edit.toPlainText()

    def test_log_detail_block_writes_title_and_messages(self, widget: LogWidget):
        """Detay bloğu başlık ve tüm satırları yazmalı."""
        widget.log_detail_block(
            "Kaynak ayrıntıları:",
            ["Satır 2 atlandı", "Satır 3 atlandı"],
        )

        lines = widget._text_edit.toPlainText().splitlines()
        assert lines == [
            "Kaynak ayrıntıları:",
            "Satır 2 atlandı",
            "Satır 3 atlandı",
        ]
        markdown = widget._text_edit.toMarkdown()
        assert "### Kaynak ayrıntıları:" in markdown
        assert "- Satır 2 atlandı" in markdown

    def test_log_summary_block_skips_none_values(self, widget: LogWidget):
        """Özet bloğu None alanları yazmamalı."""
        widget.log_summary_block(
            [
                ("Durum", "Başarılı"),
                ("Hata", None),
                ("Eklenen kayıt", 2),
            ]
        )

        lines = widget._text_edit.toPlainText().splitlines()
        assert lines == [
            "Özet:",
            "Durum: Başarılı",
            "Eklenen kayıt: 2",
        ]
        markdown = widget._text_edit.toMarkdown()
        assert "### Özet:" in markdown
        assert "- **Durum:** Başarılı" in markdown

    @patch("src.gui.log_widget.QStandardPaths.writableLocation")
    def test_log_persists_markdown_to_writable_location(
        self,
        mock_writable_location,
        qapp,
        tmp_path,
    ):
        """Log içeriği Qt'nin verdiği yazılabilir dizine markdown olarak kaydedilmeli."""
        mock_writable_location.return_value = str(tmp_path)
        widget = LogWidget(log_name="test_oturumu")

        try:
            widget.log("İşlem başlatılıyor...")
            widget.log_summary_block([("Durum", "Başarılı")])

            log_file = widget.log_file_path()
            assert log_file is not None
            assert log_file.exists()
            assert log_file.suffix == ".md"
            assert log_file.parent == tmp_path / "logs"
            assert "- İşlem başlatılıyor..." in log_file.read_text(encoding="utf-8")
            assert "### Özet:" in log_file.read_text(encoding="utf-8")
        finally:
            widget.close()
