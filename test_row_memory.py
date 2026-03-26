from pathlib import Path
from PySide6.QtWidgets import QApplication, QCheckBox
from PySide6.QtCore import Qt

from gyatt_o_tune.ui.main_window import MainWindow, TableLogChannelPreferencesDialog
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


def test_ctrl_z_reverts_cell_to_selection_start_and_redo_restores_latest() -> None:
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
    assert abs(float(window.pending_row_values[0]) - base_col0) < 1e-9
    assert window.selected_table_row_idx == row_index

    window._redo_pending_row_edit()
    assert abs(float(window.pending_row_values[0]) - (base_col0 + 0.2)) < 1e-9

    window._redo_pending_row_edit()
    assert abs(float(window.pending_row_values[1]) - (base_col1 + 0.1)) < 1e-9

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


def test_scatterplot_point_detail_preferences_do_not_remove_scatter_series() -> None:
    _get_app()
    window = MainWindow()

    tune_path = Path("tuning_data/SP_NB2_Rev1.msq")
    log_path = Path("tuning_data/SP.msl")
    assert tune_path.exists(), "Test tune not found at tuning_data/SP_NB2_Rev1.msq"
    assert log_path.exists(), "Test log not found at tuning_data/SP.msl"

    window.tune_data = TuneLoader().load(tune_path)
    parse_result = window.log_loader.load_log_with_report(log_path)
    window.log_df = parse_result.dataframe

    window.current_table = window.tune_data.tables["veTable1"]
    window.current_x_axis, window.current_y_axis = window.tune_data.resolve_table_axes(window.current_table)
    window.selected_table_row_idx = 0
    window.pending_row_values = [float(value) for value in window.current_table.values[0]]
    window.table_log_channel_preferences = {
        "veTable1": {
            "VE1": {"show_in_scatterplot": True, "show_when_point_selected": False},
        },
        "__point_cloud__": {
            "VE1": {"RPM": True, "MAP": True, "TPS": True},
        },
    }

    payload = window._build_row_visualization_payload(
        float(window.current_y_axis.values[0]),
        [float(value) for value in window.current_x_axis.values],
        window.pending_row_values,
    )

    scatter_series = [series for series in payload.get("point_sets", []) if series.get("series_id") == "log::VE1"]
    assert scatter_series, "Expected VE1 scatter series to remain present when detail channels are configured"
    detail_channels = scatter_series[0].get("detail_channels", {})
    assert isinstance(detail_channels, dict)
    assert sorted(detail_channels.keys()) == ["MAP", "RPM", "TPS"]
    assert payload.get("title") and "Error:" not in str(payload.get("title"))

    window.close()


def test_preferences_dialog_apply_keeps_current_table_scatter_series(tmp_path) -> None:
    _get_app()
    window = MainWindow()

    tune_path = Path("tuning_data/SP_NB2_Rev1.msq")
    log_path = Path("tuning_data/SP.msl")
    assert tune_path.exists(), "Test tune not found at tuning_data/SP_NB2_Rev1.msq"
    assert log_path.exists(), "Test log not found at tuning_data/SP.msl"

    window.tune_data = TuneLoader().load(tune_path)
    parse_result = window.log_loader.load_log_with_report(log_path)
    window.log_df = parse_result.dataframe

    window.current_table = window.tune_data.tables["veTable1"]
    window.current_x_axis, window.current_y_axis = window.tune_data.resolve_table_axes(window.current_table)
    window.selected_table_row_idx = 0
    window.pending_row_values = [float(value) for value in window.current_table.values[0]]
    window.table_log_channel_preferences = {
        "veTable1": {
            "VE1": {"show_in_scatterplot": True, "show_when_point_selected": False},
        },
        "__point_cloud__": {},
    }

    table_names = sorted(window.tune_data.tables.keys())
    table_display_names = {table_name: table_name for table_name in table_names}
    table_dimensions = {
        table_name: (table.rows == 1 or table.cols == 1)
        for table_name, table in window.tune_data.tables.items()
    }
    log_channel_names = [str(column) for column in window.log_df.columns]

    dialog = TableLogChannelPreferencesDialog(
        table_names=table_names,
        table_dimensions=table_dimensions,
        table_display_names=table_display_names,
        log_channel_names=log_channel_names,
        current_preferences=window.table_log_channel_preferences,
        current_preferences_path=tmp_path / "current_preferences.json",
        initial_table_name="veTable1",
        parent=window,
    )

    source_index = dialog.point_cloud_source_combo.findText("VE1")
    assert source_index >= 0, "Expected VE1 to be available as a Scatterplot Points source"
    dialog.point_cloud_source_combo.setCurrentIndex(source_index)
    for row_index in range(dialog.point_cloud_details_table.rowCount()):
        name_item = dialog.point_cloud_details_table.item(row_index, 0)
        check_widget = dialog.point_cloud_details_table.cellWidget(row_index, 1)
        check_box = check_widget.property("checkbox") if check_widget is not None else None
        if name_item is None or not isinstance(check_box, QCheckBox):
            continue
        check_box.setChecked(name_item.text() in {"RPM", "MAP", "TPS"})

    dialog.accept()
    window.table_log_channel_preferences = dialog.preferences()

    payload = window._build_row_visualization_payload(
        float(window.current_y_axis.values[0]),
        [float(value) for value in window.current_x_axis.values],
        window.pending_row_values,
    )

    scatter_series = [series for series in payload.get("point_sets", []) if series.get("series_id") == "log::VE1"]
    assert scatter_series, "Expected VE1 scatter series to remain present after applying dialog preferences"
    detail_channels = scatter_series[0].get("detail_channels", {})
    assert isinstance(detail_channels, dict)
    assert sorted(detail_channels.keys()) == ["MAP", "RPM", "TPS"]
    assert payload.get("title") and "Error:" not in str(payload.get("title"))

    window.close()


def test_custom_identifier_series_is_computed_and_plotted() -> None:
    _get_app()
    window = MainWindow()

    tune_path = Path("tuning_data/SP_NB2_Rev1.msq")
    log_path = Path("tuning_data/SP.msl")
    assert tune_path.exists(), "Test tune not found at tuning_data/SP_NB2_Rev1.msq"
    assert log_path.exists(), "Test log not found at tuning_data/SP.msl"

    window.tune_data = TuneLoader().load(tune_path)
    parse_result = window.log_loader.load_log_with_report(log_path)
    window.log_df = parse_result.dataframe

    window.current_table = window.tune_data.tables["veTable1"]
    window.current_x_axis, window.current_y_axis = window.tune_data.resolve_table_axes(window.current_table)
    window.selected_table_row_idx = 0
    window.pending_row_values = [float(value) for value in window.current_table.values[0]]
    window.table_log_channel_preferences = {
        "veTable1": {
            "VE1 Custom": {"show_in_scatterplot": True, "show_when_point_selected": True},
        },
        "__point_cloud__": {},
        "__custom_identifiers__": {
            "VE1 Custom": {"expression": "((VE1)/100)*VE1", "units": "%"},
        },
    }

    payload = window._build_row_visualization_payload(
        float(window.current_y_axis.values[0]),
        [float(value) for value in window.current_x_axis.values],
        window.pending_row_values,
    )

    computed_series = [series for series in payload.get("point_sets", []) if series.get("series_id") == "log::VE1 Custom"]
    assert computed_series, "Expected computed custom identifier series to be present"
    assert len(computed_series[0].get("rpm", [])) > 0
    table_series = next((series for series in payload.get("point_sets", []) if series.get("series_id") == "table"), None)
    assert isinstance(table_series, dict)
    extra_channels = table_series.get("extra_channels", {})
    assert isinstance(extra_channels, dict)
    assert "VE1 Custom" in extra_channels

    window.close()


def test_selected_scatter_point_text_only_shows_configured_detail_channels() -> None:
    _get_app()
    window = MainWindow()

    text = window.table_row_panel._format_selected_point_text(
        {
            "series_id": "log::VE1",
            "name": "VE1",
            "rpm": [1500.0],
            "ve": [62.5],
            "map": [48.0],
            "detail_channels": {
                "TPS": [12.0],
                "CLT": [180.0],
            },
        },
        0,
    )

    assert "TPS:" in text
    assert "CLT:" in text
    assert "VE1:" in text
    assert text.find("VE1:") < text.find("CLT:")
    assert text.find("VE1:") < text.find("TPS:")
    assert "RPM:" not in text
    assert "MAP:" not in text
    assert "Configured data:" not in text

    window.close()


def test_selected_table_point_text_only_shows_configured_table_channels() -> None:
    _get_app()
    window = MainWindow()

    text = window.table_row_panel._format_selected_point_text(
        {
            "series_id": "table",
            "name": "Selected Row Data",
            "rpm": [2000.0],
            "ve": [65.0],
            "map": [55.0],
            "extra_channels": {
                "EGO": [101.0],
            },
        },
        0,
    )

    assert "EGO:" in text
    assert "Selected Row Data:" in text
    assert text.find("Selected Row Data:") < text.find("EGO:")
    assert "RPM:" not in text
    assert "MAP:" not in text
    assert "Configured data:" not in text

    window.close()

