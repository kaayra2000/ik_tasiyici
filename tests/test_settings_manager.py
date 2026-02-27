"""SettingsManager birim testleri."""

from unittest.mock import MagicMock, patch

import pytest

from src.gui.settings_manager import SettingsManager


class TestSettingsManager:
    """SettingsManager sınıfı testleri."""

    @pytest.fixture()
    def manager(self):
        """Her test için temiz bir SettingsManager oluşturur."""
        with patch("src.gui.settings_manager.QSettings") as mock_cls:
            mock_settings = MagicMock()
            mock_cls.return_value = mock_settings
            mgr = SettingsManager("TestOrg", "TestApp")
            mgr._mock = mock_settings  # test erişimi için
            yield mgr

    def test_get_varsayilan_deger(self, manager):
        """Anahtar yoksa varsayılan değer dönmeli."""
        manager._mock.value.return_value = ""
        assert manager.get("olmayan_key", "default") == ""

    def test_set_ve_get(self, manager):
        """Kaydedilen değer okunabilmeli."""
        manager.set("test_key", "/bir/yol")
        manager._mock.setValue.assert_called_once_with("test_key", "/bir/yol")

    def test_get_existing_file_dosya_var(self, manager, tmp_path):
        """Geçerli dosya yolu dönmeli."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        manager._mock.value.return_value = str(test_file)
        result = manager.get_existing_file("key")
        assert result == str(test_file)

    def test_get_existing_file_dosya_yok(self, manager):
        """Dosya yoksa boş string dönmeli."""
        manager._mock.value.return_value = "/olmayan/dosya.xlsx"
        result = manager.get_existing_file("key")
        assert result == ""

    def test_get_existing_file_bos_deger(self, manager):
        """Boş değer için boş string dönmeli."""
        manager._mock.value.return_value = ""
        result = manager.get_existing_file("key")
        assert result == ""

    def test_get_parent_dir_gecerli(self, manager, tmp_path):
        """Üst dizin geçerliyse döndürmeli."""
        test_file = tmp_path / "alt" / "dosya.xlsx"
        test_file.parent.mkdir(parents=True, exist_ok=True)
        manager._mock.value.return_value = str(test_file)
        result = manager.get_parent_dir("key")
        assert result == str(test_file.parent)

    def test_get_parent_dir_bos(self, manager):
        """Boş değer için boş string dönmeli."""
        manager._mock.value.return_value = ""
        result = manager.get_parent_dir("key")
        assert result == ""

    def test_anahtar_sabitleri(self):
        """Sınıf sabitleri doğru tanımlanmış olmalı."""
        assert SettingsManager.KEY_INPUT_PATH == "last_input_path"
        assert SettingsManager.KEY_TEMPLATE_PATH == "last_template_path"
        assert SettingsManager.KEY_OUTPUT_PATH == "last_output_path"
