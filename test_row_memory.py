from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from gyatt_o_tune.ui.main_window import MainWindow
from gyatt_o_tune.core.io import TuneLoader

def _get_app() -> QApplication:
    app = QApplication.instance()
    return app if app is not None else QApplication([])


def test_row_selection_memory() -> None:
    _get_app()
    window = MainWindow()

    test_tune = Path("tuning_data/Final.msq")
    assert test_tune.exists(), "Test tune not found at tuning_data/Final.msq"

    tune_loader = TuneLoader()
    window.tune_data = tune_loader.load(test_tune)
    window.selected_rows_per_table = {}

    assert hasattr(window, "selected_rows_per_table"), "selected_rows_per_table attribute missing"
    assert isinstance(window.selected_rows_per_table, dict), "selected_rows_per_table should be a dict"

    window._update_table_display()
    assert window.table_list.count() > 1, "Need at least two tables for test"

    first_item = window.table_list.item(0)
    window._on_table_selected(first_item, None)
    first_table_name = first_item.data(Qt.ItemDataRole.UserRole)
    assert window.current_table is not None, "Table should be loaded"

    if window.current_table.rows <= 1:
        window.close()
        return

    window._load_selected_table_row(0)
    assert window.selected_table_row_idx == 0, "Row should be selected"

    second_item = window.table_list.item(1)
    window._on_table_selected(second_item, first_item)

    assert first_table_name in window.selected_rows_per_table, "Row selection should be saved for first table"

    window._on_table_selected(first_item, second_item)
    assert window.selected_table_row_idx == 0, "Row selection should be restored for first table"

    window.close()


def _load_first_editable_2d_table(window: MainWindow) -> tuple[str, int, int]:
    window._update_table_display()
    assert window.table_list.count() > 0, "No tables available"

    for idx in range(window.table_list.count()):
        item = window.table_list.item(idx)
        table_name = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(table_name, str):
            continue
        table = window.tune_data.tables.get(table_name) if window.tune_data else None
        if table is None:
            continue
        lower_name = table_name.lower()
        supports_row_editor = ("ve" in lower_name) or ("afr" in lower_name) or ("knock" in lower_name)
        if table.rows > 1 and table.cols > 1 and supports_row_editor:
            window._on_table_selected(item, None)
            window._load_selected_table_row(0, preferred_column=0)
            if len(window.pending_row_values) >= 2:
                return table_name, 0, 0

    raise AssertionError("No row-editable 2D table found for undo test")


def test_ctrl_z_reverts_cells_per_increment() -> None:
    _get_app()
    window = MainWindow()

    test_tune = Path("tuning_data/Final.msq")
    assert test_tune.exists(), "Test tune not found at tuning_data/Final.msq"
    window.tune_data = TuneLoader().load(test_tune)

    _table_name, row_index, _ = _load_first_editable_2d_table(window)
    assert window.current_table is not None

    base_col0 = float(window.pending_row_values[0])
    base_col1 = float(window.pending_row_values[1])

    window._adjust_pending_row_value(0, 0.1)
    window._adjust_pending_row_value(0, 0.1)
    window._adjust_pending_row_value(1, 0.1)

    assert abs(float(window.pending_row_values[0]) - (base_col0 + 0.2)) < 1e-9
    assert abs(float(window.pending_row_values[1]) - (base_col1 + 0.1)) < 1e-9

    window._undo_pending_row_edit()
    assert abs(float(window.pending_row_values[1]) - base_col1) < 1e-9
    assert abs(float(window.pending_row_values[0]) - (base_col0 + 0.2)) < 1e-9

    window._undo_pending_row_edit()
    assert abs(float(window.pending_row_values[0]) - (base_col0 + 0.1)) < 1e-9

    window._undo_pending_row_edit()
    assert abs(float(window.pending_row_values[0]) - base_col0) < 1e-9
    assert window.selected_table_row_idx == row_index

    window.close()


def test_ctrl_z_after_save_restores_previous_cell_value(tmp_path) -> None:
    _get_app()
    window = MainWindow()

    test_tune = Path("tuning_data/Final.msq")
    assert test_tune.exists(), "Test tune not found at tuning_data/Final.msq"
    window.tune_data = TuneLoader().load(test_tune)

    table_name, row_index, col_index = _load_first_editable_2d_table(window)
    assert window.current_table is not None

    baseline = float(window.pending_row_values[col_index])
    window._adjust_pending_row_value(col_index, 0.1)
    edited_value = float(window.pending_row_values[col_index])
    assert abs(edited_value - (baseline + 0.1)) < 1e-9

    save_path = tmp_path / "undo_after_save_test.msq"
    window._save_tune_to_path(save_path)
    assert save_path.exists(), "Expected save output file to exist"
    assert window.current_table is not None
    assert abs(float(window.current_table.values[row_index][col_index]) - baseline) < 1e-9
    assert abs(float(window.pending_row_values[col_index]) - edited_value) < 1e-9

    window._undo_pending_row_edit()

    assert abs(float(window.pending_row_values[col_index]) - baseline) < 1e-9
    assert abs(float(window.current_table.values[row_index][col_index]) - baseline) < 1e-9
    saved_state = window.pending_edits_per_table.get(table_name, {})
    rows_state = saved_state.get("rows", {}) if isinstance(saved_state.get("rows", {}), dict) else {}
    assert row_index not in rows_state

    window.close()


def test_global_undo_navigates_to_edited_table() -> None:
    _get_app()
    window = MainWindow()

    test_tune = Path("tuning_data/Final.msq")
    assert test_tune.exists(), "Test tune not found at tuning_data/Final.msq"
    window.tune_data = TuneLoader().load(test_tune)

    table_name, row_index, col_index = _load_first_editable_2d_table(window)
    assert window.current_table is not None
    original_value = float(window.pending_row_values[col_index])

    window._adjust_pending_row_value(col_index, 0.1)
    assert abs(float(window.pending_row_values[col_index]) - (original_value + 0.1)) < 1e-9

    switched = False
    for idx in range(window.table_list.count()):
        candidate = window.table_list.item(idx)
        candidate_name = candidate.data(Qt.ItemDataRole.UserRole)
        if isinstance(candidate_name, str) and candidate_name != table_name:
            window._on_table_selected(candidate, None)
            switched = True
            break

    assert switched, "Need a second table to validate cross-table undo navigation"

    window._trigger_global_undo()

    assert window.current_table is not None
    assert window.current_table.name == table_name
    assert window.selected_table_row_idx == row_index
    assert abs(float(window.pending_row_values[col_index]) - original_value) < 1e-9

    window.close()

