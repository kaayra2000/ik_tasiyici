"""LogWidget birim testleri."""

import os

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
