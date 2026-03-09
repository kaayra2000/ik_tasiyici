"""MainMenuWindow birim testleri."""

import os
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PyQt6.QtWidgets import QApplication

from src.gui.main_menu_window import MainMenuWindow


@pytest.fixture(scope="session")
def qapp():
    """Test oturumu için tek bir QApplication örneği sağlar."""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class TestMainMenuWindow:
    """MainMenuWindow davranış testleri."""

    @patch.object(MainMenuWindow, "_open_tutanak_window")
    def test_tutanak_button_opens_generation_flow(
        self,
        mock_open,
        qapp,
    ):
        """Tutanak butonu ilgili akışı tetiklemeli."""
        window = MainMenuWindow()

        window._tutanak_button.click()

        mock_open.assert_called_once()
        window.close()

    @patch.object(MainMenuWindow, "_open_education_import_window")
    def test_education_button_opens_import_flow(
        self,
        mock_open,
        qapp,
    ):
        """Mezuniyet butonu ilgili akışı tetiklemeli."""
        window = MainMenuWindow()

        window._education_button.click()

        mock_open.assert_called_once()
        window.close()

