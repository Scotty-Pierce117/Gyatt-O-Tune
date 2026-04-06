"""Preferences dialogs and configuration file I/O for Gyatt-O-Tune.

All user-visible preferences are persisted as a single JSON file at
``preferences_path()``.  The top-level structure is::

    {
        "version": 1,
        "tables": {
            # ── special keys ──────────────────────────────────────────────
            "__show_tunerstudio_names__": true,
            "__favorite_tables__": ["veTable1", ...],
            "__table_filter_state__": "favorites" | "1d" | "2d",
            "__custom_identifiers__": {
                "<name>": {"expression": "...", "units": "..."},
                ...
            },
            "__point_cloud__": {
                "<source_channel>": {"<target_channel>": true, ...},
                ...
            },
            "__advanced__": {
                "afr_prediction": {
                    "ve_actual_identifier": "...",
                    "afr_identifier": "..."
                }
            },
            "__table_plugins__": {
                "<table_name>": {"afr_prediction": true, ...},
                ...
            },

            # ── per-table channel preferences ─────────────────────────────
            "<table_name>": {
                "<channel_name>": {
                    "show_in_scatterplot": false,
                    "scatter_color": "#6ED2C8",
                    "scatter_opacity": 70,
                    "map_tolerance": 0
                },
                ...
            },
            ...
        }
    }

When adding a new preference to the UI, make sure to:

1.  Store it in the appropriate ``__<key>__`` or per-table dict when the dialog
    collects it (``preferences()`` method of ``TableLogChannelPreferencesDialog``).
2.  Read it back in ``_apply_*`` / ``_load_from_json_file`` inside the dialog.
3.  Add a normalisation branch in ``normalize_preferences()`` below so that
    round-trips through disk never silently drop or corrupt the value.
4.  Hydrate it in ``MainWindow._load_table_log_channel_preferences()`` (or the
    appropriate QSettings loader) so it is applied at startup and after a
    "Load Preferences…" from the main window.
"""  # noqa: E501

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# ── Canonical preferences file location ─────────────────────────────────────


def preferences_path() -> Path:
    """Return the path to the current preferences JSON file."""
    return Path.home() / ".gyatt_o_tune" / "current_preferences.json"


# ── Normalisation / serialization helpers ───────────────────────────────────


def normalize_preferences(raw: dict[str, Any]) -> dict[str, Any]:
    """Normalise a raw preferences dict so it is safe to store on disk.

    Every new special key added to the file format must have a branch here
    that validates / coerces its value so invalid data on disk can never
    propagate into the running application.
    """
    normalized: dict[str, Any] = {}

    for table_name, channels in raw.items():
        # ── special metadata keys ────────────────────────────────────────
        if table_name == "__show_tunerstudio_names__":
            normalized[table_name] = bool(channels)
            continue

        if table_name == "__custom_identifiers__":
            if isinstance(channels, dict):
                normalized[table_name] = {}
                for identifier_name, cfg in channels.items():
                    if (
                        not isinstance(identifier_name, str)
                        or not identifier_name.strip()
                        or not isinstance(cfg, dict)
                    ):
                        continue
                    expression = str(cfg.get("expression", "")).strip()
                    if not expression:
                        continue
                    normalized[table_name][identifier_name.strip()] = {
                        "expression": expression,
                        "units": str(cfg.get("units", "")).strip(),
                    }
            continue

        if table_name == "__point_cloud__":
            if isinstance(channels, dict):
                normalized[table_name] = {}
                for source_name, targets in channels.items():
                    if not isinstance(targets, dict):
                        continue
                    normalized[table_name][source_name] = {
                        target_name: bool(enabled)
                        for target_name, enabled in targets.items()
                    }
            continue

        if table_name == "__favorite_tables__":
            if isinstance(channels, list):
                normalized[table_name] = [
                    str(name)
                    for name in channels
                    if isinstance(name, str) and str(name).strip()
                ]
            continue

        if table_name == "__advanced__":
            if isinstance(channels, dict):
                normalized[table_name] = channels
            continue

        if table_name == "__table_plugins__":
            if isinstance(channels, dict):
                normalized[table_name] = {
                    str(saved_table_name): {
                        str(plugin_name): bool(enabled)
                        for plugin_name, enabled in plugin_map.items()
                        if isinstance(plugin_name, str)
                    }
                    for saved_table_name, plugin_map in channels.items()
                    if isinstance(saved_table_name, str) and isinstance(plugin_map, dict)
                }
            continue

        if table_name == "__plugins__":
            # Built-in plugin definitions are always reconstructed from constants;
            # skip persisting to avoid stale data.
            continue

        if table_name == "__table_filter_state__":
            if isinstance(channels, str) and channels in ("1d", "2d", "favorites"):
                normalized[table_name] = channels
            continue

        # ── per-table channel preferences ────────────────────────────────
        if not isinstance(channels, dict):
            continue
        normalized[table_name] = {}
        for channel_name, prefs in channels.items():
            if not isinstance(prefs, dict):
                continue
            normalized[table_name][channel_name] = {
                "show_in_scatterplot": bool(prefs.get("show_in_scatterplot", False)),
                "scatter_color": str(prefs.get("scatter_color", "")),
                "scatter_opacity": max(0, min(100, int(prefs.get("scatter_opacity", 70)))),
                "map_tolerance": max(0, min(200, int(prefs.get("map_tolerance", 0)))),
            }

    return normalized


def load_preferences() -> dict[str, Any]:
    """Load, parse, and return normalised preferences from the canonical path.

    Returns an empty dict on any error (missing file, parse error, wrong type).
    """
    path = preferences_path()
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    tables = payload.get("tables", payload)
    if not isinstance(tables, dict):
        return {}
    return normalize_preferences(tables)


def save_preferences(tables: dict[str, Any]) -> None:
    """Normalise *tables* and write it to the canonical preferences path.

    Raises ``OSError`` / ``PermissionError`` propagated from ``Path.write_text``
    so the caller can show a user-facing error message.
    """
    path = preferences_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "tables": normalize_preferences(tables),
    }
    path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")


# ── Dialog classes ───────────────────────────────────────────────────────────


class RowVisualizationPreferencesDialog(QDialog):
    """Preferences dialog for row visualization series visibility by table type."""

    SERIES_LABELS: dict[str, str] = {
        "table": "Selected Row Data",
        "raw": "Raw VE1",
        "corrected": "EGO Corrected VE",
        "afr": "AFR",
        "afr_target": "AFR Target",
        "afr_error": "AFR Error",
        "knock": "Knock In",
    }

    TABLE_TYPE_LABELS: dict[str, str] = {
        "ve": "VE Tables",
        "afr": "AFR Tables",
        "knock": "Knock Threshold Tables",
        "generic": "Other Supported Tables",
    }

    def __init__(
        self,
        table_type_series: dict[str, list[str]],
        current_preferences: dict[str, dict[str, bool]],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Row Data Visualization Preferences")
        self._table_type_series = table_type_series
        self._check_boxes: dict[tuple[str, str], QCheckBox] = {}

        layout = QVBoxLayout(self)

        info_label = QLabel("Choose which series should be shown by default for each table type.")
        layout.addWidget(info_label)

        form_layout = QFormLayout()
        for table_type, series_ids in self._table_type_series.items():
            box = QGroupBox(self.TABLE_TYPE_LABELS.get(table_type, table_type.upper()))
            box_layout = QVBoxLayout(box)
            for series_id in series_ids:
                check = QCheckBox(self.SERIES_LABELS.get(series_id, series_id))
                check.setChecked(bool(current_preferences.get(table_type, {}).get(series_id, True)))
                box_layout.addWidget(check)
                self._check_boxes[(table_type, series_id)] = check
            form_layout.addRow(box)

        layout.addLayout(form_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def preferences(self) -> dict[str, dict[str, bool]]:
        updated: dict[str, dict[str, bool]] = {table_type: {} for table_type in self._table_type_series.keys()}
        for (table_type, series_id), check in self._check_boxes.items():
            updated[table_type][series_id] = check.isChecked()
        return updated


class IdentifierExpressionLineEdit(QLineEdit):
    """QLineEdit for expressions: extends the right-click context menu with an 'Insert Identifier' submenu."""

    def __init__(self, get_identifiers: Callable[[], list[str]], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._get_identifiers = get_identifiers

    def contextMenuEvent(self, event) -> None:  # type: ignore[override]
        menu = self.createStandardContextMenu()
        menu.addSeparator()
        insert_menu = menu.addMenu("Insert Identifier")
        identifiers = sorted(self._get_identifiers())
        if identifiers:
            for name in identifiers:
                action = insert_menu.addAction(name)
                action.triggered.connect(lambda checked=False, n=name: self.insert(f'"{ n }"'))
        else:
            no_action = insert_menu.addAction("(no identifiers available)")
            no_action.setEnabled(False)
        menu.exec(event.globalPos())


class TableLogChannelPreferencesDialog(QDialog):
    """Per-table preferences with tabbed UI for table options and log channel picks."""

    _DEFAULT_SCATTER_PALETTE = [
        "#6ED2C8",
        "#F5C85F",
        "#96BEF5",
        "#F087AA",
        "#B4DC8C",
        "#DCAAF0",
    ]
    _BUILT_IN_PLUGINS: dict[str, dict[str, str]] = {
        "VE Row Best-Fit Suggestion": {
            "type": "ve_row_best_fit_suggestion",
            "button_label": "VE Row Best-Fit Suggestion",
        },
    }

    def __init__(
        self,
        table_names: list[str],
        table_dimensions: dict[str, bool],
        table_display_names: dict[str, str] | None,
        log_channel_names: list[str],
        current_preferences: dict[str, dict[str, dict[str, bool]]],
        current_preferences_path: Path | None = None,
        initial_table_name: str | None = None,
        favorite_tables: set[str] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.resize(980, 580)

        self._all_table_names = sorted(table_names)
        self._table_dimensions = dict(table_dimensions)
        self._table_display_names = dict(table_display_names or {})
        self._base_log_channel_names = list(log_channel_names)
        self._custom_identifiers = self._normalize_custom_identifiers(current_preferences.get("__custom_identifiers__", {}))
        self._log_channel_names = self._combined_identifier_names()
        self._working_preferences = self._normalize_preferences(current_preferences)
        self._point_cloud_preferences = self._normalize_point_cloud_preferences(current_preferences.get("__point_cloud__", {}))
        self._show_tunerstudio_names = bool(current_preferences.get("__show_tunerstudio_names__", True))
        self._current_preferences_path = current_preferences_path
        self._cancel_requested = False
        self._initial_table_name = initial_table_name if isinstance(initial_table_name, str) else None
        self._active_point_cloud_source: str | None = None
        saved_favorites = current_preferences.get("__favorite_tables__", [])
        if isinstance(saved_favorites, list):
            self._favorite_tables = {str(name) for name in saved_favorites if isinstance(name, str) and str(name).strip()}
        else:
            self._favorite_tables = set()
        if not self._favorite_tables:
            self._favorite_tables = set(favorite_tables or set())
        self._show_only_favorited = False
        # Tracks which table's data is currently displayed in channels_table.
        # Used by _save_current_table_state so it always saves to the right key
        # even though _active_table_name() has already advanced to the new selection.
        self._displayed_table_name: str | None = None
        # Sort state for the Tables tab identifier table
        self._channels_sort_col: int = 0
        self._channels_sort_asc: bool = True
        # Sort state for the Scatterplot Points tab table
        self._cloud_sort_col: int = 0
        self._cloud_sort_asc: bool = True

        layout = QVBoxLayout(self)

        header = QLabel(
            "Configure per-table options and which log channels should be used for row visualization."
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        tabs = QTabWidget()

        tune_tab = QWidget()
        tune_tab_layout = QVBoxLayout(tune_tab)

        table_filter_row = QHBoxLayout()
        _toggle_style = (
            "QPushButton {"
            "  border: 1px solid palette(mid);"
            "  border-radius: 4px;"
            "  padding: 4px 0px;"
            "  background: transparent;"
            "  font-weight: bold;"
            "}"
            "QPushButton:hover { background: palette(button); }"
            "QPushButton:checked {"
            "  background-color: #2277CC;"
            "  border-color: #1a5f99;"
            "  color: white;"
            "}"
            "QPushButton:checked:hover { background-color: #3388DD; }"
        )
        _btn_w = 44
        self.favorite_btn = QPushButton("★")
        self.favorite_btn.setCheckable(True)
        self.favorite_btn.setChecked(False)
        self.favorite_btn.setFixedWidth(_btn_w)
        self.favorite_btn.setStyleSheet(_toggle_style)
        self.favorite_btn.setToolTip("Show only favorite tables")
        self.favorite_btn.toggled.connect(self._on_table_filter_mode_toggled)
        table_filter_row.addWidget(self.favorite_btn)
        self.show_1d_btn = QPushButton("1D")
        self.show_1d_btn.setCheckable(True)
        self.show_1d_btn.setChecked(True)
        self.show_1d_btn.setFixedWidth(_btn_w)
        self.show_1d_btn.setStyleSheet(_toggle_style)
        self.show_1d_btn.setToolTip("Show 1D tables")
        self.show_1d_btn.toggled.connect(self._on_table_filter_mode_toggled)
        table_filter_row.addWidget(self.show_1d_btn)
        self.show_2d_btn = QPushButton("2D")
        self.show_2d_btn.setCheckable(True)
        self.show_2d_btn.setChecked(False)
        self.show_2d_btn.setFixedWidth(_btn_w)
        self.show_2d_btn.setStyleSheet(_toggle_style)
        self.show_2d_btn.setToolTip("Show 2D tables")
        self.show_2d_btn.toggled.connect(self._on_table_filter_mode_toggled)
        table_filter_row.addWidget(self.show_2d_btn)
        self._table_filter_group = QButtonGroup(self)
        self._table_filter_group.setExclusive(True)
        self._table_filter_group.addButton(self.favorite_btn)
        self._table_filter_group.addButton(self.show_1d_btn)
        self._table_filter_group.addButton(self.show_2d_btn)
        # Separator between filter buttons and options on the right
        _sep = QLabel("|")
        _sep.setStyleSheet("color: palette(mid); margin: 0 4px;")
        table_filter_row.addWidget(_sep)
        self.tunerstudio_names_checkbox = QCheckBox("TunerStudio Names")
        self.tunerstudio_names_checkbox.setChecked(self._show_tunerstudio_names)
        self.tunerstudio_names_checkbox.setToolTip("Display TunerStudio-defined table names instead of internal key names")
        table_filter_row.addWidget(self.tunerstudio_names_checkbox)
        table_filter_row.addStretch(1)
        self.advanced_table_button = QPushButton("Advanced")
        self.advanced_table_button.setToolTip(
            "Enable or disable advanced features for the selected table."
        )
        self.advanced_table_button.clicked.connect(self._show_table_advanced_menu)
        table_filter_row.addWidget(self.advanced_table_button)
        tune_tab_layout.addLayout(table_filter_row)

        table_picker_row = QHBoxLayout()
        table_picker_row.addWidget(QLabel("Table:"))
        self.table_combo = QComboBox()
        self.table_combo.currentIndexChanged.connect(self._on_table_changed)
        table_picker_row.addWidget(self.table_combo, 1)
        tune_tab_layout.addLayout(table_picker_row)

        self.log_points_label = QLabel("Data log points for selected table")
        tune_tab_layout.addWidget(self.log_points_label)

        self.channels_table = QTableWidget()
        self.channels_table.setColumnCount(5)
        self.channels_table.setHorizontalHeaderLabels(["Identifier", "Display in Scatterplot", "Color", "Opacity", "Tolerance"])
        scatter_header = self.channels_table.horizontalHeaderItem(1)
        if scatter_header is not None:
            scatter_header.setToolTip("shows up as a point cloud in the graph")
        color_header = self.channels_table.horizontalHeaderItem(2)
        if color_header is not None:
            color_header.setToolTip("Scatterplot point color (only active when Display in Scatterplot is checked)")
        opacity_header = self.channels_table.horizontalHeaderItem(3)
        if opacity_header is not None:
            opacity_header.setToolTip("Scatterplot point opacity 0-100% (only active when Display in Scatterplot is checked)")
        map_tol_header = self.channels_table.horizontalHeaderItem(4)
        if map_tol_header is not None:
            map_tol_header.setToolTip("Per-identifier MAP band half-width (kPa). 0 = use the auto-calculated global tolerance.")
        self.channels_table.verticalHeader().setVisible(False)
        self.channels_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        _ch_hdr = self.channels_table.horizontalHeader()
        _ch_hdr.setSortIndicatorShown(True)
        _ch_hdr.setSortIndicator(0, Qt.SortOrder.AscendingOrder)
        _ch_hdr.sectionClicked.connect(self._on_channels_header_clicked)
        tune_tab_layout.addWidget(self.channels_table, 1)

        tabs.addTab(tune_tab, "Tables")

        custom_tab = QWidget()
        custom_tab_layout = QVBoxLayout(custom_tab)
        custom_group = QGroupBox("Macros")
        custom_layout = QVBoxLayout(custom_group)
        custom_inputs_row = QHBoxLayout()
        custom_inputs_row.addWidget(QLabel("Name"))
        self.custom_name_edit = QLineEdit()
        self.custom_name_edit.setPlaceholderText("e.g. VE1_corr")
        custom_inputs_row.addWidget(self.custom_name_edit, 1)
        custom_inputs_row.addWidget(QLabel("Units"))
        self.custom_units_edit = QLineEdit()
        self.custom_units_edit.setPlaceholderText("e.g. %")
        custom_inputs_row.addWidget(self.custom_units_edit, 1)
        custom_layout.addLayout(custom_inputs_row)

        expression_row = QHBoxLayout()
        expression_row.addWidget(QLabel("Expression"))
        self.custom_expression_edit = IdentifierExpressionLineEdit(self._combined_identifier_names)
        self.custom_expression_edit.setPlaceholderText('Wrap identifiers in double quotes, e.g. (("EGO cor1")/100)*"VE1"')
        expression_row.addWidget(self.custom_expression_edit, 1)
        custom_layout.addLayout(expression_row)

        custom_buttons_row = QHBoxLayout()
        self.custom_add_update_button = QPushButton("Add / Update")
        self.custom_add_update_button.clicked.connect(self._on_add_or_update_custom_identifier)
        custom_buttons_row.addWidget(self.custom_add_update_button)
        self.custom_remove_button = QPushButton("Remove Selected")
        self.custom_remove_button.clicked.connect(self._on_remove_custom_identifier)
        custom_buttons_row.addWidget(self.custom_remove_button)
        custom_buttons_row.addStretch(1)
        custom_layout.addLayout(custom_buttons_row)

        self.custom_identifiers_table = QTableWidget()
        self.custom_identifiers_table.setColumnCount(3)
        self.custom_identifiers_table.setHorizontalHeaderLabels(["Identifier", "Units", "Expression"])
        self.custom_identifiers_table.verticalHeader().setVisible(False)
        self.custom_identifiers_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.custom_identifiers_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.custom_identifiers_table.itemSelectionChanged.connect(self._on_custom_identifier_selected)
        custom_layout.addWidget(self.custom_identifiers_table)

        custom_tab_layout.addWidget(custom_group, 1)
        tabs.addTab(custom_tab, "Macros")

        log_tab = QWidget()
        log_tab_layout = QVBoxLayout(log_tab)
        source_row = QHBoxLayout()
        source_row.addWidget(QLabel("Identifier:"))
        self.point_cloud_source_combo = QComboBox()
        self.point_cloud_source_combo.addItems(sorted(self._log_channel_names, key=str.lower))
        self.point_cloud_source_combo.currentTextChanged.connect(self._on_point_cloud_source_changed)
        source_row.addWidget(self.point_cloud_source_combo, 1)
        log_tab_layout.addLayout(source_row)

        self.point_cloud_details_table = QTableWidget()
        self.point_cloud_details_table.setColumnCount(2)
        self.point_cloud_details_table.setHorizontalHeaderLabels(["Identifier", "Display in Summary"])
        id_header = self.point_cloud_details_table.horizontalHeaderItem(0)
        if id_header is not None:
            id_header.setToolTip("Channel name as it appears in the data log")
        detail_header = self.point_cloud_details_table.horizontalHeaderItem(1)
        if detail_header is not None:
            detail_header.setToolTip("Shows value in upper right hand corner of graph when the point is selected")
        self.point_cloud_details_table.verticalHeader().setVisible(False)
        self.point_cloud_details_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        _cl_hdr = self.point_cloud_details_table.horizontalHeader()
        _cl_hdr.setSortIndicatorShown(True)
        _cl_hdr.setSortIndicator(0, Qt.SortOrder.AscendingOrder)
        _cl_hdr.sectionClicked.connect(self._on_cloud_header_clicked)
        log_tab_layout.addWidget(self.point_cloud_details_table, 1)

        tabs.addTab(log_tab, "Scatterplot Points")

        # ── Advanced tab ────────────────────────────────────────────────────
        advanced_tab = QWidget()
        advanced_tab_layout = QVBoxLayout(advanced_tab)

        adv_desc = QLabel(
            "Configure advanced analysis features that use log data together with tune tables."
        )
        adv_desc.setWordWrap(True)
        advanced_tab_layout.addWidget(adv_desc)

        feature_row = QHBoxLayout()
        feature_row.addWidget(QLabel("Feature:"))
        self.advanced_feature_combo = QComboBox()
        self.advanced_feature_combo.addItems(["AFR Prediction"])
        feature_row.addWidget(self.advanced_feature_combo, 1)
        advanced_tab_layout.addLayout(feature_row)

        self.advanced_stack = QStackedWidget()

        # ── AFR Prediction pane ──────────────────────────────────────────────
        afr_pred_widget = QWidget()
        afr_pred_layout = QVBoxLayout(afr_pred_widget)

        afr_pred_group = QGroupBox("AFR Prediction")
        afr_pred_group_layout = QFormLayout(afr_pred_group)

        afr_pred_desc = QLabel(
            "Predicts the expected AFR at each VE table cell using a local linear "
            "regression of logged data. The prediction changes smoothly as you "
            "edit VE values, giving a real-world usable tuning reference."
        )
        afr_pred_desc.setWordWrap(True)
        afr_pred_group_layout.addRow(afr_pred_desc)

        afr_pred_help_btn = QPushButton("?")
        afr_pred_help_btn.setFixedSize(24, 24)
        afr_pred_help_btn.setToolTip("How AFR Prediction works")
        afr_pred_help_btn.clicked.connect(self._show_afr_prediction_help)
        help_row = QHBoxLayout()
        help_row.addStretch(1)
        help_row.addWidget(afr_pred_help_btn)
        afr_pred_group_layout.addRow(help_row)

        self.afr_pred_ve_actual_combo = QComboBox()
        self.afr_pred_ve_actual_combo.addItems(self._log_channel_names)
        afr_pred_group_layout.addRow(QLabel("VE Actual Identifier:"), self.afr_pred_ve_actual_combo)

        self.afr_pred_afr_combo = QComboBox()
        self.afr_pred_afr_combo.addItems(self._log_channel_names)
        afr_pred_group_layout.addRow(QLabel("AFR Identifier:"), self.afr_pred_afr_combo)

        afr_pred_layout.addWidget(afr_pred_group)
        afr_pred_layout.addStretch(1)
        self.advanced_stack.addWidget(afr_pred_widget)

        self.advanced_feature_combo.currentIndexChanged.connect(self.advanced_stack.setCurrentIndex)
        advanced_tab_layout.addWidget(self.advanced_stack)
        advanced_tab_layout.addStretch(1)
        tabs.addTab(advanced_tab, "Advanced")

        # Populate Advanced tab from saved preferences
        self._apply_advanced_preferences(current_preferences.get("__advanced__", {}))

        layout.addWidget(tabs, 1)
        self.tabs = tabs

        button_row = QHBoxLayout()
        save_file_button = QPushButton("Save Preferences As...")
        save_file_button.clicked.connect(self._save_to_json_file)
        button_row.addWidget(save_file_button)

        load_file_button = QPushButton("Load Preferences...")
        load_file_button.clicked.connect(self._load_from_json_file)
        button_row.addWidget(load_file_button)

        button_row.addStretch(1)
        dialog_buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        cancel_button = dialog_buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_button is not None:
            cancel_button.clicked.connect(self._on_cancel_clicked)
        dialog_buttons.accepted.connect(self.accept)
        dialog_buttons.rejected.connect(self.reject)
        button_row.addWidget(dialog_buttons)
        layout.addLayout(button_row)

        self._refresh_table_combo()
        if self._log_channel_names:
            self._active_point_cloud_source = self._log_channel_names[0]
            self._populate_point_cloud_targets(self._active_point_cloud_source)
        self._refresh_custom_identifier_table()

    def select_tune_tables_tab(self, table_name: str | None = None) -> None:
        """Switch to the Tables tab and optionally pre-select a table."""
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == "Tables":
                self.tabs.setCurrentIndex(i)
                break
        if table_name is not None:
            index = self.table_combo.findData(table_name)
            if index >= 0:
                self.table_combo.setCurrentIndex(index)

    def select_scatterplot_tab(self, identifier: str | None = None) -> None:
        """Switch to the Scatterplot Points tab and optionally pre-select an identifier."""
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == "Scatterplot Points":
                self.tabs.setCurrentIndex(i)
                break
        if identifier is not None and self.point_cloud_source_combo.findText(identifier) >= 0:
            self.point_cloud_source_combo.setCurrentText(identifier)

    def _apply_advanced_preferences(self, adv: Any) -> None:
        """Populate Advanced tab widgets from a saved preferences dict."""
        if not isinstance(adv, dict):
            return
        afr_cfg = adv.get("afr_prediction", {})
        if not isinstance(afr_cfg, dict):
            return
        ve_id = str(afr_cfg.get("ve_actual_identifier", ""))
        if ve_id and self.afr_pred_ve_actual_combo.findText(ve_id) >= 0:
            self.afr_pred_ve_actual_combo.setCurrentText(ve_id)
        afr_id = str(afr_cfg.get("afr_identifier", ""))
        if afr_id and self.afr_pred_afr_combo.findText(afr_id) >= 0:
            self.afr_pred_afr_combo.setCurrentText(afr_id)

    def _advanced_preferences(self) -> dict[str, Any]:
        """Return current state of the Advanced tab as a serialisable dict."""
        return {
            "afr_prediction": {
                "ve_actual_identifier": self.afr_pred_ve_actual_combo.currentText(),
                "afr_identifier": self.afr_pred_afr_combo.currentText(),
            },
        }

    def _show_afr_prediction_help(self) -> None:
        """Show a detailed explanation of how AFR Prediction works."""
        text = (
            "<h3>AFR Prediction</h3>"
            "<p>AFR Prediction estimates the Air/Fuel Ratio your engine would "
            "produce at a given VE table value, based on your logged data.</p>"
            "<h4>Method — Local Linear Regression</h4>"
            "<p>For each RPM column in the selected table row the algorithm:</p>"
            "<ol>"
            "<li><b>Collects</b> all log data points in the current MAP band "
            "whose RPM is within &plusmn;150 RPM of the column's RPM value — "
            "regardless of the current VE cell value.</li>"
            "<li><b>Fits a best-fit line</b> (linear regression) through those "
            "points using VE Actual on the X-axis and AFR on the Y-axis.</li>"
            "<li><b>Evaluates</b> the best-fit line at the current table VE "
            "value to produce the predicted AFR.</li>"
            "</ol>"
            "<h4>Why This Approach?</h4>"
            "<p>By fitting a trend line through <i>all</i> the nearby data "
            "instead of only using points that exactly match the current VE "
            "value, the prediction:</p>"
            "<ul>"
            "<li>Changes <b>smoothly and predictably</b> as you adjust VE — "
            "small VE changes produce proportionally small AFR changes.</li>"
            "<li>Uses <b>all available data</b> in the RPM neighbourhood, so "
            "noisy individual readings are averaged out.</li>"
            "<li>Gives a <b>real-world usable</b> reference for choosing a VE "
            "value that achieves your target AFR.</li>"
            "</ul>"
            "<h4>Fallback</h4>"
            "<p>If fewer than 3 data points are available near an RPM column "
            "(not enough for a reliable trend line), the simple average AFR of "
            "the available points is used instead.</p>"
        )
        msg = QMessageBox(self)
        msg.setWindowTitle("AFR Prediction — How It Works")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(text)
        msg.exec()

    def _on_cancel_clicked(self) -> None:
        self._cancel_requested = True

    def _on_scatter_checkbox_changed(self, row: int, checked: bool) -> None:
        """Enable/disable color+opacity controls and refresh filter visibility."""
        enabled = bool(checked)
        color_btn = self.channels_table.cellWidget(row, 2)
        opacity_spin = self.channels_table.cellWidget(row, 3)
        if color_btn is not None:
            color_btn.setEnabled(enabled)
        if opacity_spin is not None:
            opacity_spin.setEnabled(enabled)

    def _on_channels_header_clicked(self, col: int) -> None:
        """Sort the Tables-tab identifier table by the clicked column."""
        if col == 0:
            if self._channels_sort_col == 0:
                self._channels_sort_asc = not self._channels_sort_asc
            else:
                self._channels_sort_col = 0
                self._channels_sort_asc = True
        elif col == 1:
            self._channels_sort_col = 1
        else:
            return
        self._save_current_table_state()
        table_name = self._displayed_table_name
        if table_name:
            self._populate_channels_for_table(table_name)

    def _on_cloud_header_clicked(self, col: int) -> None:
        """Sort the Scatterplot Points table by the clicked column."""
        if col == 0:
            if self._cloud_sort_col == 0:
                self._cloud_sort_asc = not self._cloud_sort_asc
            else:
                self._cloud_sort_col = 0
                self._cloud_sort_asc = True
        elif col == 1:
            self._cloud_sort_col = 1
        else:
            return
        source_name = self.point_cloud_source_combo.currentText()
        self._save_point_cloud_state(source_name)
        if source_name:
            self._populate_point_cloud_targets(source_name)

    def _on_table_filter_mode_toggled(self, checked: bool) -> None:
        """Handle exclusive filter mode changes for the Tables tab."""
        if not checked:
            return
        self._show_only_favorited = self.favorite_btn.isChecked()
        self._refresh_table_combo()

    def _make_centered_checkbox(self, checked: bool, on_changed: Callable[[bool], None] | None = None) -> QWidget:
        """Create a truly centered checkbox cell widget for QTableWidget."""
        box = QCheckBox()
        box.setChecked(bool(checked))
        holder = QWidget()
        layout = QHBoxLayout(holder)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(box)
        holder.setProperty("checkbox", box)
        if on_changed is not None:
            box.toggled.connect(on_changed)
        return holder

    @staticmethod
    def _table_checkbox_checked(table: QTableWidget, row: int, column: int) -> bool:
        widget = table.cellWidget(row, column)
        if widget is None:
            return False
        box = widget.property("checkbox")
        if isinstance(box, QCheckBox):
            return box.isChecked()
        return False

    def _make_color_button(self, hex_color: str) -> QPushButton:
        """Create a color-swatch button that opens a color picker on click."""
        qc = QColor(hex_color)
        lum = 0.299 * qc.red() + 0.587 * qc.green() + 0.114 * qc.blue()
        fg = "#000000" if lum > 128 else "#ffffff"
        btn = QPushButton("")
        btn.setStyleSheet(f"background-color: {hex_color}; color: {fg}; border: 1px solid #666;")
        btn.setProperty("scatter_color", hex_color)

        def _pick_color(checked: bool = False, button: QPushButton = btn) -> None:
            current = QColor(str(button.property("scatter_color") or "#6ED2C8"))
            chosen = QColorDialog.getColor(current, self, "Pick Scatter Color")
            if not chosen.isValid():
                return
            new_hex = chosen.name()
            lum2 = 0.299 * chosen.red() + 0.587 * chosen.green() + 0.114 * chosen.blue()
            fg2 = "#000000" if lum2 > 128 else "#ffffff"
            button.setStyleSheet(f"background-color: {new_hex}; color: {fg2}; border: 1px solid #666;")
            button.setProperty("scatter_color", new_hex)

        btn.clicked.connect(_pick_color)
        return btn

    def reject(self) -> None:  # type: ignore[override]
        # Reject should always discard in-dialog edits unless explicitly accepted.
        super().reject()

    @staticmethod
    def _normalize_preferences(raw: dict[str, Any]) -> dict[str, Any]:
        """Normalise working in-dialog preferences (per-table channel data + plugins).

        This is distinct from the module-level ``normalize_preferences()`` which
        handles the full on-disk format including metadata keys.
        """
        normalized: dict[str, Any] = {}
        normalized_plugins: dict[str, dict[str, Any]] = {
            name: dict(cfg) for name, cfg in TableLogChannelPreferencesDialog._BUILT_IN_PLUGINS.items()
        }
        plugins = raw.get("__plugins__")
        if isinstance(plugins, dict):
            normalized_plugins.update(
                {
                    str(name): dict(cfg)
                    for name, cfg in plugins.items()
                    if isinstance(name, str) and isinstance(cfg, dict)
                }
            )
        normalized["__plugins__"] = normalized_plugins
        table_plugins = raw.get("__table_plugins__")
        if isinstance(table_plugins, dict):
            normalized["__table_plugins__"] = {
                str(table_name): {
                    str(plugin_name): bool(enabled)
                    for plugin_name, enabled in plugin_map.items()
                    if isinstance(plugin_name, str)
                }
                for table_name, plugin_map in table_plugins.items()
                if isinstance(table_name, str) and isinstance(plugin_map, dict)
            }
        adv = raw.get("__advanced__")
        if isinstance(adv, dict):
            normalized["__advanced__"] = adv
        for table_name, channels in raw.items():
            if table_name in {"__point_cloud__", "__custom_identifiers__", "__plugins__", "__table_plugins__", "__advanced__"}:
                continue
            if not isinstance(channels, dict):
                continue
            normalized[table_name] = {}
            for channel_name, prefs in channels.items():
                if not isinstance(prefs, dict):
                    continue
                normalized[table_name][channel_name] = {
                    "show_in_scatterplot": bool(prefs.get("show_in_scatterplot", False)),
                    "scatter_color": str(prefs.get("scatter_color", "")),
                    "scatter_opacity": max(0, min(100, int(prefs.get("scatter_opacity", 70)))),
                    "map_tolerance": max(0, min(200, int(prefs.get("map_tolerance", 0)))),
                }
        return normalized

    @staticmethod
    def _normalize_custom_identifiers(raw: Any) -> dict[str, dict[str, str]]:
        normalized: dict[str, dict[str, str]] = {}
        if not isinstance(raw, dict):
            return normalized
        for name, cfg in raw.items():
            if not isinstance(name, str) or not name.strip() or not isinstance(cfg, dict):
                continue
            expression = str(cfg.get("expression", "")).strip()
            if not expression:
                continue
            normalized[name.strip()] = {
                "expression": expression,
                "units": str(cfg.get("units", "")).strip(),
            }
        return normalized

    def _combined_identifier_names(self) -> list[str]:
        merged: list[str] = []
        seen: set[str] = set()
        for name in self._base_log_channel_names + list(self._custom_identifiers.keys()):
            if name in seen:
                continue
            seen.add(name)
            merged.append(name)
        return merged

    def _extract_identifiers_from_expression(self, expression: str, candidates: list[str]) -> tuple[list[str], str | None]:
        candidate_set = set(candidates)
        found: list[str] = []
        quote_pattern = re.compile(r'"([^"]*)"')
        for match in quote_pattern.finditer(expression):
            name = match.group(1)
            if not name:
                return [], "Quoted identifier cannot be empty."
            if name not in candidate_set:
                return [], f'Unknown identifier: "{name}". All identifiers must be wrapped in double quotes.'
            found.append(name)
        scrubbed = quote_pattern.sub("", expression)
        scrubbed = re.sub(r"[0-9eE+\-*/().\s]", "", scrubbed)
        if scrubbed:
            return [], f"Expression contains unexpected characters outside of quoted identifiers: '{scrubbed}'."
        return found, None

    def _refresh_identifier_lists(self, preserve_source: str | None = None) -> None:
        self._save_current_table_state()
        self._save_point_cloud_state(self._active_point_cloud_source)
        self._log_channel_names = self._combined_identifier_names()
        self._reconcile_identifier_preferences()
        current_source = preserve_source if preserve_source is not None else self.point_cloud_source_combo.currentText()
        self.point_cloud_source_combo.blockSignals(True)
        self.point_cloud_source_combo.clear()
        self.point_cloud_source_combo.addItems(sorted(self._log_channel_names, key=str.lower))
        if current_source and current_source in self._log_channel_names:
            self.point_cloud_source_combo.setCurrentText(current_source)
        elif self._log_channel_names:
            self.point_cloud_source_combo.setCurrentIndex(0)
        self.point_cloud_source_combo.blockSignals(False)
        self._active_point_cloud_source = self.point_cloud_source_combo.currentText() if self._log_channel_names else None
        # Keep Advanced tab identifier combos in sync with the current identifier list
        _adv_ve = self.afr_pred_ve_actual_combo.currentText()
        _adv_afr = self.afr_pred_afr_combo.currentText()
        self.afr_pred_ve_actual_combo.blockSignals(True)
        self.afr_pred_afr_combo.blockSignals(True)
        self.afr_pred_ve_actual_combo.clear()
        self.afr_pred_ve_actual_combo.addItems(self._log_channel_names)
        self.afr_pred_afr_combo.clear()
        self.afr_pred_afr_combo.addItems(self._log_channel_names)
        if _adv_ve in self._log_channel_names:
            self.afr_pred_ve_actual_combo.setCurrentText(_adv_ve)
        if _adv_afr in self._log_channel_names:
            self.afr_pred_afr_combo.setCurrentText(_adv_afr)
        self.afr_pred_ve_actual_combo.blockSignals(False)
        self.afr_pred_afr_combo.blockSignals(False)
        active = self._active_table_name()
        if active:
            self._populate_channels_for_table(active)
        if self._active_point_cloud_source:
            self._populate_point_cloud_targets(self._active_point_cloud_source)

    def _reconcile_identifier_preferences(self) -> None:
        valid_names = set(self._log_channel_names)

        _SPECIAL_KEYS = {
            "__point_cloud__",
            "__custom_identifiers__",
            "__plugins__",
            "__table_plugins__",
            "__advanced__",
        }

        for table_name, table_prefs in list(self._working_preferences.items()):
            if table_name in _SPECIAL_KEYS:
                continue
            if not isinstance(table_prefs, dict):
                continue
            self._working_preferences[table_name] = {
                channel_name: prefs
                for channel_name, prefs in table_prefs.items()
                if isinstance(channel_name, str) and channel_name in valid_names and isinstance(prefs, dict)
            }

        reconciled_point_cloud: dict[str, dict[str, bool]] = {}
        for source_name, targets in list(self._point_cloud_preferences.items()):
            if not isinstance(source_name, str) or source_name not in valid_names or not isinstance(targets, dict):
                continue
            reconciled_point_cloud[source_name] = {
                target_name: bool(enabled)
                for target_name, enabled in targets.items()
                if isinstance(target_name, str) and target_name in valid_names and target_name != source_name
            }
        self._point_cloud_preferences = reconciled_point_cloud

    def _refresh_custom_identifier_table(self) -> None:
        names = sorted(self._custom_identifiers.keys())
        self.custom_identifiers_table.setRowCount(len(names))
        for row, name in enumerate(names):
            cfg = self._custom_identifiers.get(name, {})
            name_item = QTableWidgetItem(name)
            name_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.custom_identifiers_table.setItem(row, 0, name_item)
            units_item = QTableWidgetItem(str(cfg.get("units", "")))
            units_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.custom_identifiers_table.setItem(row, 1, units_item)
            expr_item = QTableWidgetItem(str(cfg.get("expression", "")))
            expr_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.custom_identifiers_table.setItem(row, 2, expr_item)
        self.custom_identifiers_table.resizeColumnsToContents()

    def _on_custom_identifier_selected(self) -> None:
        row = self.custom_identifiers_table.currentRow()
        if row < 0:
            return
        name_item = self.custom_identifiers_table.item(row, 0)
        if name_item is None:
            return
        name = name_item.text()
        cfg = self._custom_identifiers.get(name, {})
        self.custom_name_edit.setText(name)
        self.custom_units_edit.setText(str(cfg.get("units", "")))
        self.custom_expression_edit.setText(str(cfg.get("expression", "")))

    def _on_add_or_update_custom_identifier(self) -> None:
        name = self.custom_name_edit.text().strip()
        units = self.custom_units_edit.text().strip()
        expression = self.custom_expression_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Custom Identifier", "Name is required.")
            return
        if not expression:
            QMessageBox.warning(self, "Custom Identifier", "Expression is required.")
            return
        if name in self._base_log_channel_names:
            QMessageBox.warning(self, "Custom Identifier", "Name conflicts with an existing log identifier.")
            return

        candidates = [n for n in self._combined_identifier_names() if n != name]
        found, error = self._extract_identifiers_from_expression(expression, candidates)
        if error is not None:
            QMessageBox.warning(self, "Custom Identifier", error)
            return
        if len(set(found)) == 0:
            QMessageBox.warning(self, "Custom Identifier", "Expression must reference at least one identifier.")
            return
        if name in found:
            QMessageBox.warning(self, "Custom Identifier", "Expression cannot directly reference itself.")
            return

        self._custom_identifiers[name] = {"expression": expression, "units": units}
        self._refresh_custom_identifier_table()
        self._refresh_identifier_lists(preserve_source=self._active_point_cloud_source)

    def _on_remove_custom_identifier(self) -> None:
        row = self.custom_identifiers_table.currentRow()
        if row < 0:
            return
        name_item = self.custom_identifiers_table.item(row, 0)
        if name_item is None:
            return
        name = name_item.text()
        self._custom_identifiers.pop(name, None)
        for table_name, table_prefs in list(self._working_preferences.items()):
            if isinstance(table_prefs, dict) and name in table_prefs:
                table_prefs.pop(name, None)
                self._working_preferences[table_name] = table_prefs
        for source_name, targets in list(self._point_cloud_preferences.items()):
            if isinstance(targets, dict) and name in targets:
                targets.pop(name, None)
                self._point_cloud_preferences[source_name] = targets
        if name in self._point_cloud_preferences:
            self._point_cloud_preferences.pop(name, None)
        self._refresh_custom_identifier_table()
        self._refresh_identifier_lists(preserve_source=self._active_point_cloud_source)

    @staticmethod
    def _normalize_point_cloud_preferences(raw: dict[str, dict[str, bool]]) -> dict[str, dict[str, bool]]:
        normalized: dict[str, dict[str, bool]] = {}
        if not isinstance(raw, dict):
            return normalized
        for source_name, targets in raw.items():
            if not isinstance(targets, dict):
                continue
            normalized[source_name] = {target_name: bool(enabled) for target_name, enabled in targets.items()}
        return normalized

    def _refresh_table_combo(self) -> None:
        previous = self._active_table_name()
        preferred_initial = self._initial_table_name
        include_1d = self.show_1d_btn.isChecked()
        include_2d = self.show_2d_btn.isChecked()
        show_only_favorites = self.favorite_btn.isChecked()

        filtered: list[str] = []
        for table_name in self._all_table_names:
            if show_only_favorites:
                if table_name not in self._favorite_tables:
                    continue
            else:
                is_1d = bool(self._table_dimensions.get(table_name, False))
                if is_1d and not include_1d:
                    continue
                if (not is_1d) and not include_2d:
                    continue
            filtered.append(table_name)

        filtered.sort(key=lambda t: self._display_name_for_table(t).lower())

        self.table_combo.blockSignals(True)
        self.table_combo.clear()
        for table_name in filtered:
            self.table_combo.addItem(self._display_name_for_table(table_name), table_name)
        if filtered:
            if previous in filtered:
                index = self.table_combo.findData(previous)
                if index >= 0:
                    self.table_combo.setCurrentIndex(index)
                else:
                    self.table_combo.setCurrentIndex(0)
            elif preferred_initial in filtered:
                index = self.table_combo.findData(preferred_initial)
                if index >= 0:
                    self.table_combo.setCurrentIndex(index)
                else:
                    self.table_combo.setCurrentIndex(0)
            else:
                self.table_combo.setCurrentIndex(0)
            self._initial_table_name = None
        self.table_combo.blockSignals(False)

        if filtered:
            self._on_table_changed(self.table_combo.currentIndex())
        else:
            self.channels_table.clearContents()
            self.channels_table.setRowCount(0)
            self.log_points_label.setText("No tables match the current filters.")

    def _active_table_name(self) -> str | None:
        table_name = self.table_combo.currentData()
        return str(table_name) if isinstance(table_name, str) and table_name else None

    def _display_name_for_table(self, table_name: str) -> str:
        return self._table_display_names.get(table_name, table_name)

    def _table_plugin_preferences(self, table_name: str | None = None) -> dict[str, bool]:
        active_table = table_name if table_name is not None else self._active_table_name()
        if not active_table:
            return {}
        table_plugins = self._working_preferences.get("__table_plugins__", {})
        if not isinstance(table_plugins, dict):
            return {}
        plugin_map = table_plugins.get(active_table, {})
        if not isinstance(plugin_map, dict):
            return {}
        return {
            str(plugin_name): bool(enabled)
            for plugin_name, enabled in plugin_map.items()
            if isinstance(plugin_name, str)
        }

    def _set_table_plugin_enabled(self, plugin_name: str, enabled: bool, table_name: str | None = None) -> None:
        active_table = table_name if table_name is not None else self._active_table_name()
        if not active_table:
            return
        table_plugins = self._working_preferences.setdefault("__table_plugins__", {})
        if not isinstance(table_plugins, dict):
            table_plugins = {}
            self._working_preferences["__table_plugins__"] = table_plugins
        plugin_map = table_plugins.setdefault(active_table, {})
        if not isinstance(plugin_map, dict):
            plugin_map = {}
            table_plugins[active_table] = plugin_map
        plugin_map[str(plugin_name)] = bool(enabled)

    def _update_advanced_button_text(self, table_name: str | None = None) -> None:
        plugin_map = self._table_plugin_preferences(table_name)
        enabled_count = sum(1 for enabled in plugin_map.values() if enabled)
        label = "Advanced" if enabled_count == 0 else f"Advanced ({enabled_count})"
        self.advanced_table_button.setText(label)

    def _show_table_advanced_menu(self) -> None:
        table_name = self._active_table_name()
        if not table_name:
            return
        plugin_map = self._table_plugin_preferences(table_name)
        display_name = self._display_name_for_table(table_name)

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Advanced Features — {display_name}")
        dlg.setMinimumWidth(340)
        outer = QVBoxLayout(dlg)
        outer.setSpacing(10)

        title_label = QLabel(f"<b>{display_name}</b>")
        outer.addWidget(title_label)

        group = QGroupBox("Enable / Disable Features")
        group_layout = QVBoxLayout(group)
        group_layout.setSpacing(6)

        # AFR Prediction row
        afr_row = QHBoxLayout()
        afr_cb = QCheckBox("AFR Prediction")
        afr_cb.setChecked(bool(plugin_map.get("afr_prediction", False)))
        afr_cb.setToolTip("Predict AFR from VE and logged data for each table cell")
        afr_row.addWidget(afr_cb, 1)
        afr_help_btn = QPushButton("?")
        afr_help_btn.setFixedWidth(26)
        afr_help_btn.setToolTip("Learn how AFR Prediction works")
        afr_help_btn.clicked.connect(self._show_afr_prediction_help)
        afr_row.addWidget(afr_help_btn)
        group_layout.addLayout(afr_row)

        outer.addWidget(group)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        outer.addWidget(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._set_table_plugin_enabled("afr_prediction", afr_cb.isChecked(), table_name)
        self._update_advanced_button_text(table_name)

    def _save_current_table_state(self) -> None:
        # Use _displayed_table_name so we save the data that is *currently shown*
        # in channels_table, not the table that the combo has already advanced to.
        table_name = self._displayed_table_name
        if table_name is None:
            return

        table_prefs = self._working_preferences.setdefault(table_name, {})
        for row_index in range(self.channels_table.rowCount()):
            channel_item = self.channels_table.item(row_index, 0)
            if channel_item is None:
                continue
            channel_name = channel_item.text()
            show_scatter = self._table_checkbox_checked(self.channels_table, row_index, 1)
            color_btn = self.channels_table.cellWidget(row_index, 2)
            opacity_spin = self.channels_table.cellWidget(row_index, 3)
            map_tol_spin = self.channels_table.cellWidget(row_index, 4)
            scatter_color = str(color_btn.property("scatter_color") or "") if color_btn is not None else ""
            scatter_opacity = int(opacity_spin.value()) if opacity_spin is not None else 70
            map_tolerance = int(map_tol_spin.value()) if map_tol_spin is not None else 0
            table_prefs[channel_name] = {
                "show_in_scatterplot": show_scatter,
                "scatter_color": scatter_color,
                "scatter_opacity": scatter_opacity,
                "map_tolerance": map_tolerance,
            }

    def _populate_channels_for_table(self, table_name: str) -> None:
        table_prefs = self._working_preferences.get(table_name, {})
        # Record what is now displayed so _save_current_table_state targets the right key.
        self._displayed_table_name = table_name
        self.log_points_label.setText(f"Data log points for table: {self._display_name_for_table(table_name)}")
        self._update_advanced_button_text(table_name)

        if self._channels_sort_col == 1:
            channel_order = sorted(
                self._log_channel_names,
                key=lambda n: (not bool(table_prefs.get(n, {}).get("show_in_scatterplot", False)), n.lower()),
            )
        else:
            channel_order = sorted(
                self._log_channel_names,
                key=lambda n: n.lower(),
                reverse=not self._channels_sort_asc,
            )

        self.channels_table.blockSignals(True)
        self.channels_table.clearContents()
        self.channels_table.setRowCount(len(channel_order))
        for row_index, channel_name in enumerate(channel_order):
            pref = table_prefs.get(channel_name, {})
            name_item = QTableWidgetItem(channel_name)
            name_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.channels_table.setItem(row_index, 0, name_item)

            scatter_checked = bool(pref.get("show_in_scatterplot", False))
            scatter_widget = self._make_centered_checkbox(
                scatter_checked,
                on_changed=lambda checked, r=row_index: self._on_scatter_checkbox_changed(r, checked),
            )
            self.channels_table.setCellWidget(row_index, 1, scatter_widget)

            default_hex = self._DEFAULT_SCATTER_PALETTE[row_index % len(self._DEFAULT_SCATTER_PALETTE)]
            hex_color = str(pref.get("scatter_color", "")) or default_hex
            color_btn = self._make_color_button(hex_color)
            color_btn.setEnabled(scatter_checked)
            self.channels_table.setCellWidget(row_index, 2, color_btn)

            opacity_spin = QSpinBox()
            opacity_spin.setRange(0, 100)
            opacity_spin.setSuffix(" %")
            opacity_spin.setValue(max(0, min(100, int(pref.get("scatter_opacity", 70)))))
            opacity_spin.setEnabled(scatter_checked)
            self.channels_table.setCellWidget(row_index, 3, opacity_spin)

            map_tol_spin = QSpinBox()
            map_tol_spin.setRange(0, 200)
            map_tol_spin.setSuffix(" kPa")
            map_tol_spin.setSpecialValueText("auto")
            map_tol_spin.setValue(max(0, min(200, int(pref.get("map_tolerance", 0)))))
            map_tol_spin.setToolTip("MAP band half-width for this identifier (0 = use auto-calculated global tolerance)")
            self.channels_table.setCellWidget(row_index, 4, map_tol_spin)

        self.channels_table.resizeColumnsToContents()
        color_header = self.channels_table.horizontalHeaderItem(2)
        if color_header is not None:
            color_header_width = self.channels_table.fontMetrics().horizontalAdvance(color_header.text()) + 24
            self.channels_table.setColumnWidth(2, color_header_width)
        self.channels_table.setColumnWidth(3, 100)
        self.channels_table.setColumnWidth(4, 120)
        _hdr = self.channels_table.horizontalHeader()
        _hdr.setSortIndicatorShown(True)
        if self._channels_sort_col == 1:
            _hdr.setSortIndicator(1, Qt.SortOrder.AscendingOrder)
        else:
            _hdr.setSortIndicator(0, Qt.SortOrder.AscendingOrder if self._channels_sort_asc else Qt.SortOrder.DescendingOrder)
        self.channels_table.blockSignals(False)

    def _on_table_changed(self, index: int) -> None:
        _ = index
        self._save_current_table_state()
        table_name = self._active_table_name()
        if table_name:
            self._populate_channels_for_table(table_name)

    def preferences(self) -> dict[str, dict[str, dict[str, bool]]]:
        self._save_current_table_state()
        self._save_point_cloud_state()
        self._show_tunerstudio_names = self.tunerstudio_names_checkbox.isChecked()
        merged = self._normalize_preferences(self._working_preferences)
        merged["__point_cloud__"] = self._normalize_point_cloud_preferences(self._point_cloud_preferences)
        merged["__custom_identifiers__"] = self._normalize_custom_identifiers(self._custom_identifiers)
        merged["__show_tunerstudio_names__"] = self._show_tunerstudio_names
        merged["__favorite_tables__"] = sorted(self._favorite_tables)
        table_plugins = self._working_preferences.get("__table_plugins__", {})
        if isinstance(table_plugins, dict):
            merged["__table_plugins__"] = {
                str(table_name): {
                    str(plugin_name): bool(enabled)
                    for plugin_name, enabled in plugin_map.items()
                    if isinstance(plugin_name, str)
                }
                for table_name, plugin_map in table_plugins.items()
                if isinstance(table_name, str) and isinstance(plugin_map, dict)
            }
        merged["__advanced__"] = self._advanced_preferences()
        if self.favorite_btn.isChecked():
            merged["__table_filter_state__"] = "favorites"
        elif self.show_2d_btn.isChecked():
            merged["__table_filter_state__"] = "2d"
        else:
            merged["__table_filter_state__"] = "1d"
        return merged

    def _save_point_cloud_state(self, source_name: str | None = None) -> None:
        active_source = source_name if source_name is not None else self._active_point_cloud_source
        if active_source is None:
            active_source = self.point_cloud_source_combo.currentText()
        source_name = str(active_source)
        if not source_name:
            return
        source_prefs = self._point_cloud_preferences.setdefault(source_name, {})
        for row_index in range(self.point_cloud_details_table.rowCount()):
            name_item = self.point_cloud_details_table.item(row_index, 0)
            if name_item is None:
                continue
            target_name = name_item.text()
            source_prefs[target_name] = self._table_checkbox_checked(self.point_cloud_details_table, row_index, 1)

    def _populate_point_cloud_targets(self, source_name: str) -> None:
        source_prefs = self._point_cloud_preferences.get(source_name, {})
        targets = [name for name in self._log_channel_names if name != source_name]
        if self._cloud_sort_col == 1:
            display_targets = sorted(
                targets,
                key=lambda n: (not bool(source_prefs.get(n, False)), n.lower()),
            )
        else:
            display_targets = sorted(
                targets,
                key=lambda n: n.lower(),
                reverse=not self._cloud_sort_asc,
            )
        self.point_cloud_details_table.blockSignals(True)
        self.point_cloud_details_table.clearContents()
        self.point_cloud_details_table.setRowCount(len(display_targets))
        for row_index, target_name in enumerate(display_targets):
            name_item = QTableWidgetItem(target_name)
            name_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.point_cloud_details_table.setItem(row_index, 0, name_item)
            checked = bool(source_prefs.get(target_name, False))
            check_widget = self._make_centered_checkbox(checked)
            self.point_cloud_details_table.setCellWidget(row_index, 1, check_widget)
        self.point_cloud_details_table.resizeColumnsToContents()
        _hdr = self.point_cloud_details_table.horizontalHeader()
        _hdr.setSortIndicatorShown(True)
        if self._cloud_sort_col == 1:
            _hdr.setSortIndicator(1, Qt.SortOrder.AscendingOrder)
        else:
            _hdr.setSortIndicator(0, Qt.SortOrder.AscendingOrder if self._cloud_sort_asc else Qt.SortOrder.DescendingOrder)
        self.point_cloud_details_table.blockSignals(False)

    def _on_point_cloud_source_changed(self, source_name: str) -> None:
        self._save_point_cloud_state(self._active_point_cloud_source)
        self._active_point_cloud_source = source_name
        if source_name:
            self._populate_point_cloud_targets(source_name)

    def _save_to_json_file(self) -> None:
        self._save_current_table_state()
        self._save_point_cloud_state()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Row Data Preferences",
            "row_data_preferences.json",
            "JSON Files (*.json);;All Files (*.*)",
        )
        if not file_path:
            return

        payload = {
            "version": 1,
            "tables": self.preferences(),
        }
        try:
            Path(file_path).write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Save Preferences Error", f"Could not save preferences:\n{exc}")
            return

        QMessageBox.information(self, "Preferences Saved", f"Saved preferences to:\n{file_path}")

    def _load_from_json_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Row Data Preferences",
            "",
            "JSON Files (*.json);;All Files (*.*)",
        )
        if not file_path:
            return

        try:
            payload = json.loads(Path(file_path).read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Load Preferences Error", f"Could not read preferences file:\n{exc}")
            return

        loaded_tables = payload.get("tables", payload)
        if not isinstance(loaded_tables, dict):
            QMessageBox.critical(self, "Load Preferences Error", "Invalid preferences format.")
            return

        self._working_preferences = self._normalize_preferences(loaded_tables)
        self._point_cloud_preferences = self._normalize_point_cloud_preferences(loaded_tables.get("__point_cloud__", {}))
        self._custom_identifiers = self._normalize_custom_identifiers(loaded_tables.get("__custom_identifiers__", {}))
        self._apply_advanced_preferences(loaded_tables.get("__advanced__", {}))
        loaded_favorites = loaded_tables.get("__favorite_tables__", [])
        if isinstance(loaded_favorites, list):
            self._favorite_tables = {str(name) for name in loaded_favorites if isinstance(name, str) and str(name).strip()}

        # Apply show_tunerstudio_names setting from loaded preferences to UI
        loaded_show_names = bool(loaded_tables.get("__show_tunerstudio_names__", True))
        self._show_tunerstudio_names = loaded_show_names
        self.tunerstudio_names_checkbox.setChecked(loaded_show_names)

        # Apply table filter state from loaded preferences
        _filter_state = loaded_tables.get("__table_filter_state__", "")
        if _filter_state == "2d":
            self.show_2d_btn.setChecked(True)
        elif _filter_state == "1d":
            self.show_1d_btn.setChecked(True)
        elif _filter_state == "favorites":
            self.favorite_btn.setChecked(True)

        self._refresh_custom_identifier_table()
        self._refresh_identifier_lists(preserve_source=self._active_point_cloud_source)
        active = self._active_table_name()
        if active:
            self._populate_channels_for_table(active)
        source_name = self.point_cloud_source_combo.currentText()
        if source_name:
            self._active_point_cloud_source = source_name
            self._populate_point_cloud_targets(source_name)
        QMessageBox.information(self, "Preferences Loaded", f"Loaded preferences from:\n{file_path}")
