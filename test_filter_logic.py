from pathlib import Path

from PySide6.QtWidgets import QApplication

from gyatt_o_tune.core.io import TuneLoader
from gyatt_o_tune.ui.main_window import MainWindow


def _get_app() -> QApplication:
    app = QApplication.instance()
    return app if app is not None else QApplication([])


def test_filter_logic() -> None:
    _get_app()
    window = MainWindow()

    test_tune = Path("tuning_data/Final.msq")
    assert test_tune.exists(), "Test tune file not found at tuning_data/Final.msq"

    tune_loader = TuneLoader()
    window.tune_data = tune_loader.load(test_tune)

    for table_name, table in window.tune_data.tables.items():
        is_1d = window._is_1d_table(table_name)
        expected_1d = table.rows == 1 or table.cols == 1
        assert is_1d == expected_1d, f"Table {table_name}: got {is_1d}, expected {expected_1d}"

    # Force deterministic initial state regardless of persisted QSettings.
    window.only_show_favorited_tables_action.setChecked(False)
    window._on_only_show_favorited_toggled()
    assert window.show_1d_tables_action.isEnabled(), "1D filter should be enabled when not showing only favorites"

    window.only_show_favorited_tables_action.setChecked(True)
    window._on_only_show_favorited_toggled()
    assert not window.show_1d_tables_action.isEnabled(), "1D filter should be disabled when showing only favorites"

    window.only_show_favorited_tables_action.setChecked(False)
    window._on_only_show_favorited_toggled()
    assert window.show_1d_tables_action.isEnabled(), "1D filter should be enabled when not showing only favorites"

    window.close()

