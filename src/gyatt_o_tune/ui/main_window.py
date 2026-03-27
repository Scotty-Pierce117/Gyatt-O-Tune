from __future__ import annotations

import json
import re
import ast
import time as _time_mod
from pathlib import Path
from typing import Any, Callable

import pyqtgraph as pg
from PySide6.QtCore import QEvent, Qt, QTimer
from PySide6.QtGui import QAction, QColor, QGuiApplication, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDockWidget,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSlider,
    QSpinBox,
    QStackedWidget,
    QStatusBar,
    QStyle,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import QSettings

from gyatt_o_tune.core.io import AxisVector, LogLoader, TableData, TuneData, TuneLoader


class CopyPasteTableWidget(QTableWidget):
    """Spreadsheet-style copy/paste with TSV payload."""

    def __init__(self) -> None:
        super().__init__()
        self.footer_rows = 0
        self.on_adjust_value: Any = None
        self.on_undo: Any = None
        self.on_redo: Any = None

    def set_footer_rows(self, count: int) -> None:
        self.footer_rows = max(0, count)

    def _data_row_count(self) -> int:
        return max(0, self.rowCount() - self.footer_rows)

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        current_row = self.currentRow()
        current_column = self.currentColumn()
        in_data_region = (
            current_row >= 0
            and current_column >= 0
            and current_row < self._data_row_count()
            and current_column < self.columnCount()
        )

        if event.matches(QKeySequence.StandardKey.Undo) and in_data_region and callable(self.on_undo):
            self.on_undo(current_row, current_column)
            return
        if event.matches(QKeySequence.StandardKey.Redo) and in_data_region and callable(self.on_redo):
            self.on_redo(current_row, current_column)
            return
        if event.key() == Qt.Key.Key_Up and in_data_region and callable(self.on_adjust_value):
            self.on_adjust_value(current_row, current_column, 0.1)
            return
        if event.key() == Qt.Key.Key_Down and in_data_region and callable(self.on_adjust_value):
            self.on_adjust_value(current_row, current_column, -0.1)
            return
        if event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_selection()
            return
        super().keyPressEvent(event)

    def wheelEvent(self, event) -> None:  # type: ignore[override]
        current_row = self.currentRow()
        current_column = self.currentColumn()
        in_data_region = (
            current_row >= 0
            and current_column >= 0
            and current_row < self._data_row_count()
            and current_column < self.columnCount()
        )
        if not in_data_region or not callable(self.on_adjust_value):
            super().wheelEvent(event)
            return

        delta = event.angleDelta().y()
        if delta == 0:
            super().wheelEvent(event)
            return

        steps = max(1, abs(delta) // 120)
        increment = 0.5 if (event.modifiers() & Qt.KeyboardModifier.ControlModifier) else 0.1
        amount = float(steps * increment if delta > 0 else -steps * increment)
        self.on_adjust_value(current_row, current_column, amount)
        event.accept()

    def copy_selection(self) -> None:
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            return
        block = selected_ranges[0]
        data_rows = self._data_row_count()
        rows: list[str] = []
        for row in range(block.topRow(), block.bottomRow() + 1):
            if row >= data_rows:
                continue
            values: list[str] = []
            for col in range(block.leftColumn(), block.rightColumn() + 1):
                item = self.item(row, col)
                values.append(item.text() if item else "")
            if values:
                rows.append("\t".join(values))
        if not rows:
            return
        QGuiApplication.clipboard().setText("\n".join(rows))

    def paste_selection(self) -> None:
        start = self.currentIndex()
        if not start.isValid():
            return
        if start.row() >= self._data_row_count():
            return
        text = QGuiApplication.clipboard().text()
        if not text.strip():
            return
        rows = [line.split("\t") for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if line != ""]
        data_rows = self._data_row_count()
        for row_offset, columns in enumerate(rows):
            for col_offset, value in enumerate(columns):
                row = start.row() + row_offset
                col = start.column() + col_offset
                if row >= data_rows or col >= self.columnCount():
                    continue
                existing = self.item(row, col)
                if existing is None:
                    existing = QTableWidgetItem()
                    self.setItem(row, col, existing)
                existing.setText(value.strip())


class TuneTablesListWidget(QListWidget):
    """Tune table list that accepts dropped tune files for quick loading."""

    ALLOWED_TUNE_SUFFIXES = {".msq", ".ini", ".txt"}

    def __init__(self) -> None:
        super().__init__()
        self.on_tune_file_dropped: Any = None
        self.setAcceptDrops(True)

    def _first_valid_tune_path(self, event: Any) -> Path | None:
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            return None

        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            path = Path(url.toLocalFile())
            if not path.is_file():
                continue
            if path.suffix.lower() in self.ALLOWED_TUNE_SUFFIXES:
                return path
        return None

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if self._first_valid_tune_path(event) is not None:
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if self._first_valid_tune_path(event) is not None:
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:  # type: ignore[override]
        tune_path = self._first_valid_tune_path(event)
        if tune_path is None:
            super().dropEvent(event)
            return

        if callable(self.on_tune_file_dropped):
            self.on_tune_file_dropped(tune_path)
        event.acceptProposedAction()


class RowVisualizationPanel(QGroupBox):
    """Reusable row data visualization with crosshair and point selection."""

    def __init__(self, title: str = "") -> None:
        super().__init__(title)
        self._selected_point: tuple[str, int] | None = None
        self._point_sets: list[dict[str, Any]] = []
        self.on_point_selected: Any = None
        self.on_point_adjust_requested: Any = None
        self.on_visibility_changed: Any = None
        self.on_scatter_preferences_requested: Any = None
        self.on_table_preferences_requested: Any = None
        self.on_log_playback_toggled: Any = None
        self.get_log_playback_state: Any = None
        self.on_plot_all_data_requested: Any = None
        self.legend: Any = None
        self._table_type = "generic"
        self._y_label_text = "Value"
        self._selected_point_text = "<span style='color:#f0f0f0'>Selected point: none</span>"
        self._stats_tooltip_text = ""
        self.dynamic_scatter_items: dict[str, Any] = {}

        layout = QVBoxLayout(self)

        self.graph_title_label = QLabel("No data selected")
        self.graph_title_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.graph_title_label.setToolTip("")
        layout.addWidget(self.graph_title_label)

        self.plot = pg.PlotWidget()
        self.plot.showGrid(x=True, y=True, alpha=0.2)
        self.plot.getPlotItem().setLabel("bottom", "RPM")
        self.plot.getPlotItem().setLabel("left", "Value")

        self.table_curve = pg.PlotCurveItem(pen=pg.mkPen(255, 40, 40, width=5), name="Selected Row Data")
        self.table_scatter = pg.ScatterPlotItem(size=8, name="Selected Row Points")
        self.selected_marker = pg.ScatterPlotItem(
            pen=pg.mkPen(255, 255, 255, 255, width=2),
            brush=pg.mkBrush(255, 225, 120, 255),
            size=11,
            name="Selected Point",
        )
        self.selected_point_overlay = pg.TextItem(
            anchor=(1, 0),
            color=(240, 240, 240),
        )
        self.crosshair_vline = pg.InfiniteLine(angle=90, movable=False, pen=pg.mkPen(200, 200, 200, 140, width=1))
        self.crosshair_hline = pg.InfiniteLine(angle=0, movable=False, pen=pg.mkPen(200, 200, 200, 140, width=1))

        layout.addWidget(self.plot, 1)

        view_box = self.plot.getPlotItem().vb
        if view_box is not None:
            view_box.sigRangeChanged.connect(self._on_plot_range_changed)

        # --- Playback controls ---
        self._playback_timer = QTimer(self)
        self._playback_timer.setInterval(50)  # ~20fps update
        self._playback_timer.timeout.connect(self._playback_tick)
        self._playback_state = "stopped"  # "stopped" | "playing" | "paused"
        self._playback_start_time: float = 0.0
        self._playback_elapsed: float = 0.0  # seconds elapsed since playback start point
        self._playback_all_times: list[float] = []  # sorted unique times across all log series
        self._playback_series_time_data: dict[str, list[float]] = {}  # series_id -> time values
        self._playback_current_wall_start: float = 0.0  # monotonic ref for live ticking
        self._playback_max_time: float = 0.0
        self._has_playback_time_data = False
        self._show_playback_controls = True
        self._playback_speed: float = 1.0
        self._speed_steps: list[float] = [1.0, 2.0, 5.0, 10.0, 20.0]
        self._speed_index: int = 0

        playback_layout = QHBoxLayout()
        playback_layout.setContentsMargins(10, 8, 10, 8)
        playback_layout.setSpacing(8)

        self.playback_controls_bar = QWidget()
        self.playback_controls_bar.setObjectName("rowVizPlaybackBar")

        self.btn_play = QPushButton("")
        self.btn_play.setObjectName("rowVizPlaybackPlayButton")
        self.btn_play.setFixedSize(34, 30)
        self.btn_play.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.btn_play.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_play.setCheckable(True)
        self.btn_play.setToolTip("Play / Resume playback")
        self.btn_play.clicked.connect(self._on_playback_play)
        playback_layout.addWidget(self.btn_play)

        self.btn_pause = QPushButton("")
        self.btn_pause.setObjectName("rowVizPlaybackPauseButton")
        self.btn_pause.setFixedSize(34, 30)
        self.btn_pause.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPause))
        self.btn_pause.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pause.setCheckable(True)
        self.btn_pause.setToolTip("Pause playback")
        self.btn_pause.clicked.connect(self._on_playback_pause)
        playback_layout.addWidget(self.btn_pause)

        self.btn_stop = QPushButton("")
        self.btn_stop.setObjectName("rowVizPlaybackStopButton")
        self.btn_stop.setFixedSize(34, 30)
        self.btn_stop.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaStop))
        self.btn_stop.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_stop.setCheckable(True)
        self.btn_stop.setToolTip("Stop playback and show all data")
        self.btn_stop.clicked.connect(self._on_playback_stop)
        playback_layout.addWidget(self.btn_stop)

        self.btn_speed = QPushButton("1x")
        self.btn_speed.setObjectName("rowVizPlaybackSpeedButton")
        self.btn_speed.setFixedSize(46, 30)
        self.btn_speed.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_speed.setToolTip("Cycle playback speed: 1x \u2192 2x \u2192 5x \u2192 10x \u2192 20x")
        self.btn_speed.clicked.connect(self._on_playback_speed_toggle)
        playback_layout.addWidget(self.btn_speed)

        self.playback_slider = QSlider(Qt.Orientation.Horizontal)
        self.playback_slider.setMinimum(0)
        self.playback_slider.setMaximum(1000)
        self.playback_slider.setValue(0)
        self.playback_slider.setToolTip("Scrub playback position")
        self.playback_slider.sliderPressed.connect(self._on_slider_pressed)
        self.playback_slider.sliderReleased.connect(self._on_slider_released)
        self.playback_slider.valueChanged.connect(self._on_slider_value_changed)
        self._slider_dragging = False
        playback_layout.addWidget(self.playback_slider, 1)

        self.playback_time_label = QLabel("0.0s / 0.0s")
        self.playback_time_label.setFixedWidth(110)
        playback_layout.addWidget(self.playback_time_label)

        self.playback_widget = QWidget()
        self.playback_controls_bar.setLayout(playback_layout)
        self._apply_playback_button_style()
        playback_widget_layout = QVBoxLayout(self.playback_widget)
        playback_widget_layout.setContentsMargins(0, 0, 0, 0)
        playback_widget_layout.setSpacing(0)
        playback_widget_layout.addWidget(self.playback_controls_bar)
        self.playback_widget.setVisible(False)  # hidden until log scatter data is loaded
        self._update_playback_button_states()
        layout.addWidget(self.playback_widget)

        self._mouse_proxy = pg.SignalProxy(
            self.plot.scene().sigMouseMoved,
            rateLimit=30,
            slot=self._on_mouse_moved,
        )
        self.plot.scene().sigMouseClicked.connect(self._on_mouse_clicked)
        self._dynamic_menu_actions: list[Any] = []
        self._context_menu_identifier: str | None = None
        self._setup_viewbox_menu_hook()
        self.plot.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.plot.installEventFilter(self)
        self.plot.viewport().installEventFilter(self)
        self.clear_visualization()

    def _set_graph_title(self, title: str, stats_text: str = "") -> None:
        self.graph_title_label.setText(title)
        self._stats_tooltip_text = stats_text
        self.graph_title_label.setToolTip(stats_text)

    def _set_selected_point_text(self, html: str) -> None:
        self._selected_point_text = html
        self.selected_point_overlay.setHtml(html)
        self._position_selected_point_overlay()

    def _position_selected_point_overlay(self) -> None:
        plot_item = self.plot.getPlotItem()
        if plot_item is None:
            return
        view_box = plot_item.vb
        if view_box is None:
            return
        (x_min, x_max), (y_min, y_max) = view_box.viewRange()
        x_padding = (x_max - x_min) * 0.02
        y_padding = (y_max - y_min) * 0.02
        self.selected_point_overlay.setPos(x_max - x_padding, y_max - y_padding)

    def _on_plot_range_changed(self, *args: Any) -> None:
        self._position_selected_point_overlay()

    def set_playback_controls_enabled(self, enabled: bool) -> None:
        self._show_playback_controls = bool(enabled)
        self._refresh_playback_widget_visibility()

    def _refresh_playback_widget_visibility(self) -> None:
        self.playback_widget.setVisible(self._show_playback_controls)

    def _apply_playback_button_style(self) -> None:
        self.playback_controls_bar.setStyleSheet(
            "QWidget#rowVizPlaybackBar {"
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #212734, stop:1 #2a3142);"
            "border: 1px solid #3c465d;"
            "border-radius: 12px;"
            "}"
            "QWidget#rowVizPlaybackBar QLabel {"
            "color: #d3dbea;"
            "font-size: 11px;"
            "font-weight: 600;"
            "}"
            "QWidget#rowVizPlaybackBar QSlider::groove:horizontal {"
            "height: 6px;"
            "background: #1b212f;"
            "border: 1px solid #394359;"
            "border-radius: 3px;"
            "}"
            "QWidget#rowVizPlaybackBar QSlider::sub-page:horizontal {"
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #5db1ff, stop:1 #7fd0ff);"
            "border-radius: 3px;"
            "}"
            "QWidget#rowVizPlaybackBar QSlider::handle:horizontal {"
            "width: 14px;"
            "margin: -5px 0;"
            "border-radius: 7px;"
            "border: 1px solid #9ccff4;"
            "background: #d4efff;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton {"
            "background: #2f394c;"
            "color: #ebf2ff;"
            "border: 1px solid #4d5a74;"
            "border-radius: 8px;"
            "font-size: 13px;"
            "font-weight: 700;"
            "padding: 0 2px;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton:hover {"
            "background: #3b4960;"
            "border-color: #6a7da0;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton:pressed {"
            "background: #273043;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton:disabled {"
            "background: #2a2f39;"
            "color: #8a95ab;"
            "border-color: #454d5f;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton:checked {"
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2f79ff, stop:1 #3ea8ff);"
            "color: #ffffff;"
            "border-color: #85bcff;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton#rowVizPlaybackStopButton:checked {"
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #d9506f, stop:1 #c64158);"
            "border-color: #f08aa0;"
            "}"
            "QWidget#rowVizPlaybackBar QPushButton#rowVizPlaybackSpeedButton {"
            "font-size: 12px;"
            "}"
        )

    def _update_playback_button_states(self) -> None:
        self.btn_play.setChecked(self._playback_state == "playing")
        self.btn_pause.setChecked(self._playback_state == "paused")
        self.btn_stop.setChecked(self._playback_state == "stopped")

    def eventFilter(self, watched: Any, event: Any) -> bool:  # type: ignore[override]
        if watched in {self.plot, self.plot.viewport()}:
            if event.type() == QEvent.Type.Wheel:
                target_index = self._selected_table_index_for_adjustment()
                if target_index is None:
                    return False

                delta = event.angleDelta().y()
                if delta == 0:
                    delta = event.pixelDelta().y()
                if delta == 0:
                    return False
                steps = max(1, abs(delta) // 120) if abs(delta) >= 120 else 1
                increment = 0.5 if (event.modifiers() & Qt.KeyboardModifier.ControlModifier) else 0.1
                amount = float(steps * increment if delta > 0 else -steps * increment)
                if callable(self.on_point_adjust_requested):
                    # Ensure subsequent refreshes preserve selection and keep wheel-edit active.
                    self.select_table_point(target_index, emit_callback=False)
                    self.on_point_adjust_requested(target_index, amount)
                    event.accept()
                    return True

            if event.type() == QEvent.Type.KeyPress:
                if self._selected_point is None:
                    return False
                series_id, index = self._selected_point
                if series_id != "table":
                    return False

                key = event.key()
                if key != Qt.Key.Key_Tab and key != Qt.Key.Key_Backtab:
                    return False

                table_series = self._get_series("table")
                if table_series is None:
                    return False
                point_count = len(table_series.get("rpm", []))
                if point_count <= 0:
                    return False

                next_index = index - 1 if key == Qt.Key.Key_Backtab else index + 1
                if next_index < 0 or next_index >= point_count:
                    return True

                self.select_table_point(next_index, emit_callback=True)
                return True

        return super().eventFilter(watched, event)

    def _selected_table_index_for_adjustment(self) -> int | None:
        if self._selected_point is None:
            return None

        table_series = self._get_series("table")
        if table_series is None:
            return None

        table_rpm_values = table_series.get("rpm", [])
        if not hasattr(table_rpm_values, "__len__") or len(table_rpm_values) == 0:
            return None

        series_id, index = self._selected_point
        if series_id != "table":
            # Only adjust values when a table point is selected; all other series
            # fall through so the wheel event is passed on for normal graph zooming.
            return None

        if 0 <= index < len(table_rpm_values):
            return index
        return None

    def _emit_visibility_changed(self) -> None:
        if callable(self.on_visibility_changed):
            self.on_visibility_changed(self.current_series_visibility())

    def _refresh_visibility(self) -> None:
        """Update plot item visibility."""
        self.table_curve.show()
        self.table_scatter.show()
        for scatter_item in self.dynamic_scatter_items.values():
            scatter_item.show()
        self._refresh_point_styles()
        self._refresh_legend()

    def _is_series_visible(self, series_id: str) -> bool:
        return True

    def current_series_visibility(self) -> dict[str, bool]:
        return {}

    def configure_series_controls(
        self,
        available_series: list[str],
        preferred_visibility: dict[str, bool] | None = None,
    ) -> None:
        self._refresh_visibility()
        self._emit_visibility_changed()

    def clear_visualization(self, title: str = "No data selected", stats_text: str = "") -> None:
        if self._playback_state != "stopped":
            self._playback_timer.stop()
            self._playback_state = "stopped"
            self._playback_elapsed = 0.0
            self._update_playback_button_states()
        self._selected_point = None
        self._point_sets = []
        self.dynamic_scatter_items = {}
        self._table_type = "generic"
        self._y_label_text = "Value"
        self._remove_legend()
        self.plot.clear()
        self._set_graph_title(title, stats_text)
        self.plot.setLabel('left', self._y_label_text)
        self.plot.setLabel('bottom', 'RPM')
        self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")
        self._add_plot_items()
        self._collect_playback_time_data()
        self._apply_view_all()

    def set_row_data(self, payload: dict[str, Any], auto_view_all: bool = True) -> None:
        # Stop any active playback when new data arrives
        if self._playback_state != "stopped":
            self._playback_timer.stop()
            self._playback_state = "stopped"
            self._playback_elapsed = 0.0
            self._update_playback_button_states()
        self._selected_point = None
        self._point_sets = list(payload.get("point_sets", []))
        self._table_type = str(payload.get("table_type", "generic"))
        self._y_label_text = str(payload.get("y_label", "Value"))
        x_label_text = str(payload.get("x_label", "RPM"))
        available_series = [str(s) for s in payload.get("available_series", [])]
        preferred_visibility = payload.get("series_visibility", {})
        self.configure_series_controls(available_series, preferred_visibility)
        self._remove_legend()
        self.plot.clear()
        self._set_graph_title(str(payload.get("title", "")), str(payload.get("stats", "")))
        self.plot.setLabel('left', self._y_label_text)
        self.plot.setLabel('bottom', x_label_text)
        self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")
        self._add_plot_items()
        self._refresh_point_styles()
        self._refresh_visibility()
        self._collect_playback_time_data()
        if auto_view_all:
            self._apply_view_all()

    def _apply_view_all(self) -> None:
        plot_item = self.plot.getPlotItem()
        if plot_item is None:
            return
        view_box = plot_item.vb
        if view_box is None:
            return
        view_box.autoRange(padding=0.03)

    def _add_plot_items(self) -> None:
        self.plot.addItem(self.table_curve)
        self.plot.addItem(self.table_scatter)
        self.dynamic_scatter_items = {}
        dynamic_palette = [
            (110, 210, 200),
            (245, 200, 95),
            (150, 190, 245),
            (240, 135, 170),
            (180, 220, 140),
            (220, 170, 240),
        ]
        dynamic_series = [series for series in self._point_sets if str(series.get("series_id", "")).startswith("log::")]
        for idx, series in enumerate(dynamic_series):
            series_id = str(series.get("series_id"))
            raw_color = series.get("scatter_color", "")
            if isinstance(raw_color, str) and raw_color.startswith("#"):
                try:
                    qc = QColor(raw_color)
                    color = (qc.red(), qc.green(), qc.blue())
                except Exception:
                    color = dynamic_palette[idx % len(dynamic_palette)]
            else:
                color = dynamic_palette[idx % len(dynamic_palette)]
            opacity_pct = max(0, min(100, int(series.get("scatter_opacity", 70))))
            brush_alpha = round(opacity_pct * 2.0)   # 0–200
            pen_alpha = round(opacity_pct * 2.55)     # 0–255
            scatter_item = pg.ScatterPlotItem(
                pen=pg.mkPen(color[0], color[1], color[2], pen_alpha, width=1),
                brush=pg.mkBrush(color[0], color[1], color[2], brush_alpha),
                size=6,
                name=str(series.get("name", series_id)),
            )
            self.dynamic_scatter_items[series_id] = scatter_item
            self.plot.addItem(scatter_item)
        self.plot.addItem(self.selected_marker)
        self.plot.addItem(self.selected_point_overlay, ignoreBounds=True)
        self.plot.addItem(self.crosshair_vline, ignoreBounds=True)
        self.plot.addItem(self.crosshair_hline, ignoreBounds=True)
        self.selected_marker.hide()
        self.crosshair_vline.hide()
        self.crosshair_hline.hide()
        self._set_selected_point_text(self._selected_point_text)

        self._apply_item_z_order()
        self._remove_legend()
        self.legend = pg.LegendItem(offset=(50, 30))
        self.legend.setParentItem(self.plot.graphicsItem())
        self._refresh_legend()

    def _remove_legend(self) -> None:
        if self.legend is None:
            return
        try:
            scene = self.legend.scene()
            if scene is not None:
                scene.removeItem(self.legend)
            else:
                self.legend.setParentItem(None)
        except Exception:
            pass
        self.legend = None

    def _apply_item_z_order(self) -> None:
        for scatter_item in self.dynamic_scatter_items.values():
            scatter_item.setZValue(12)
        self.table_curve.setZValue(32)
        self.table_scatter.setZValue(33)
        self.selected_marker.setZValue(40)
        self.table_scatter.setZValue(32)
        self.selected_marker.setZValue(40)
        self.selected_point_overlay.setZValue(65)
        self.crosshair_vline.setZValue(60)
        self.crosshair_hline.setZValue(60)

    def _series_has_points(self, series_id: str) -> bool:
        series = self._get_series(series_id)
        if series is None:
            return False
        rpm_values = series.get("rpm", [])
        ve_values = series.get("ve", [])
        return bool(rpm_values) and bool(ve_values)

    def _refresh_legend(self) -> None:
        if self.legend is None:
            return

        self.legend.clear()

        if self._series_has_points("table"):
            self.legend.addItem(self.table_curve, "Selected Row")

        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            if not series_id.startswith("log::"):
                continue
            item = self.dynamic_scatter_items.get(series_id)
            if item is None:
                continue
            if self._series_has_points(series_id):
                self.legend.addItem(item, str(series.get("name", series_id)))

    def _get_series(self, series_id: str) -> dict[str, Any] | None:
        for series in self._point_sets:
            if series.get("series_id") == series_id:
                return series
        return None

    def _refresh_point_styles(self) -> None:
        table_series = self._get_series("table")

        if table_series is not None:
            self.table_curve.setData(table_series["rpm"], table_series["ve"])
            self.table_scatter.setData(x=table_series["rpm"], y=table_series["ve"])
        else:
            self.table_curve.setData([], [])
            self.table_scatter.setData([], [])

        # During active playback, defer log scatter rendering to _apply_playback_at_elapsed
        if not self._is_playback_active():
            for series in self._point_sets:
                series_id = str(series.get("series_id", ""))
                if not series_id.startswith("log::"):
                    continue
                item = self.dynamic_scatter_items.get(series_id)
                if item is None:
                    continue
                if self._is_series_visible(series_id):
                    item.setData(x=series.get("rpm", []), y=series.get("ve", []))
                else:
                    item.setData([], [])

        self._refresh_selected_marker()
        self._refresh_legend()

    def _refresh_selected_marker(self) -> None:
        if self._selected_point is None:
            self.selected_marker.setData([], [])
            self.selected_marker.hide()
            return

        series_id, index = self._selected_point
        if not self._is_series_visible(series_id):
            self.selected_marker.setData([], [])
            self.selected_marker.hide()
            return
        series = self._get_series(series_id)
        if series is None:
            self.selected_marker.setData([], [])
            self.selected_marker.hide()
            return

        rpm_values = series.get("rpm", [])
        ve_values = series.get("ve", [])
        if index < 0 or index >= len(rpm_values) or index >= len(ve_values):
            self.selected_marker.setData([], [])
            self.selected_marker.hide()
            return

        self.selected_marker.setData([float(rpm_values[index])], [float(ve_values[index])])
        self.selected_marker.show()

    def selected_table_point_index(self) -> int | None:
        if self._selected_point is None:
            return None
        series_id, index = self._selected_point
        if series_id != "table":
            return None
        return index

    def select_table_point(self, index: int, emit_callback: bool = False) -> None:
        table_series = self._get_series("table")
        if table_series is None:
            return
        rpm_values = table_series.get("rpm", [])
        if index < 0 or index >= len(rpm_values):
            return

        self._selected_point = ("table", index)
        self._refresh_point_styles()
        self._set_selected_point_text(self._format_selected_point_text(table_series, index))
        if emit_callback and callable(self.on_point_selected):
            self.on_point_selected("table", index, float(rpm_values[index]))

    def clear_selected_point(self) -> None:
        self._selected_point = None
        self._refresh_point_styles()
        self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")

    @staticmethod
    def _format_value(value: float | None, suffix: str = "") -> str:
        if value is None:
            return "n/a"
        return f"{value:.2f}{suffix}"

    def _format_selected_point_text(self, series: dict[str, Any], index: int) -> str:
        name = str(series.get("name", "Point"))
        header = "Table Data Point Summary:" if name == "Selected Row Data" else f"{name} Point Summary:"
        value_at_point: float | None = None
        y_values = series.get("ve")
        if isinstance(y_values, list) and 0 <= index < len(y_values):
            try:
                raw_value = float(y_values[index])
                value_at_point = raw_value if raw_value == raw_value else None
            except Exception:
                value_at_point = None

        detail_rows: list[tuple[str, str]] = [(f"{name}:", self._format_value(value_at_point))]

        extra_channels = series.get("extra_channels", {})
        if isinstance(extra_channels, dict):
            for channel_name in sorted(extra_channels.keys()):
                channel_values = extra_channels.get(channel_name)
                if isinstance(channel_values, list) and 0 <= index < len(channel_values):
                    detail_rows.append((f"{channel_name}:", self._format_value(channel_values[index])))

        count_channels = series.get("count_channels", {})
        if isinstance(count_channels, dict):
            for channel_name in sorted(count_channels.keys()):
                channel_values = count_channels.get(channel_name)
                if isinstance(channel_values, list) and 0 <= index < len(channel_values):
                    cnt = channel_values[index]
                    detail_rows.append((f"{channel_name}:", str(int(cnt)) if cnt is not None else "n/a"))

        detail_channels = series.get("detail_channels", {})
        if isinstance(detail_channels, dict):
            for channel_name in sorted(detail_channels.keys()):
                channel_values = detail_channels.get(channel_name)
                if isinstance(channel_values, list) and 0 <= index < len(channel_values):
                    detail_rows.append((f"{channel_name}:", self._format_value(channel_values[index])))

        rows_html = "".join(
            f"<tr><td style='padding-right:10px'>{label}</td><td>{value}</td></tr>"
            for label, value in detail_rows
        )
        return (
            f"<span style='color:#f0f0f0'>"
            f"<b>{header}</b><br>"
            f"<table cellpadding='0' cellspacing='2'>{rows_html}</table>"
            f"</span>"
        )

    def _on_mouse_moved(self, evt: tuple[Any]) -> None:
        if not evt:
            return
        pos = evt[0]
        if pos is None:
            return

        plot_item = self.plot.getPlotItem()
        vb = plot_item.vb
        if vb is None:
            return

        if not plot_item.sceneBoundingRect().contains(pos):
            self.crosshair_vline.hide()
            self.crosshair_hline.hide()
            return

        mouse_point = vb.mapSceneToView(pos)
        rpm = float(mouse_point.x())
        ve = float(mouse_point.y())
        self.crosshair_vline.setPos(rpm)
        self.crosshair_hline.setPos(ve)
        self.crosshair_vline.show()
        self.crosshair_hline.show()

    def _on_mouse_clicked(self, evt: Any) -> None:
        if evt is None or evt.button() != Qt.MouseButton.LeftButton:
            return

        plot_item = self.plot.getPlotItem()
        vb = plot_item.vb
        if vb is None:
            return

        click_pos = evt.scenePos()

        # Bottom-left corner click (where the two axis lines meet) triggers "Plot All Data".
        vb_rect = vb.sceneBoundingRect()
        plot_rect = plot_item.sceneBoundingRect()
        if (
            click_pos.x() < vb_rect.left()
            and click_pos.y() > vb_rect.bottom()
            and plot_rect.contains(click_pos)
            and callable(self.on_plot_all_data_requested)
        ):
            self.on_plot_all_data_requested()
            return

        if not plot_item.sceneBoundingRect().contains(click_pos):
            return
        self.plot.setFocus()

        click_view = vb.mapSceneToView(click_pos)
        click_rpm = float(click_view.x())
        click_val = float(click_view.y())

        view_rect = vb.viewRect()
        scene_rect = plot_item.sceneBoundingRect()
        x_per_px = max(float(view_rect.width()) / max(float(scene_rect.width()), 1.0), 1e-9)
        y_per_px = max(float(view_rect.height()) / max(float(scene_rect.height()), 1.0), 1e-9)

        nearest: tuple[float, str, int] | None = None
        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            if not self._is_series_visible(series_id):
                continue
            for index, (rpm, ve) in enumerate(zip(series.get("rpm", []), series.get("ve", []))):
                dx_px = (float(rpm) - click_rpm) / x_per_px
                dy_px = (float(ve) - click_val) / y_per_px
                distance = dx_px**2 + dy_px**2
                if nearest is None or distance < nearest[0]:
                    nearest = (distance, series_id, index)

        max_pick_distance_squared = 12.0 ** 2
        if nearest is None or nearest[0] > max_pick_distance_squared:
            self._selected_point = None
            self._refresh_point_styles()
            self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")
            if callable(self.on_point_selected):
                self.on_point_selected(None, None, None)
            return

        _, series_id, index = nearest
        selected_series = self._get_series(series_id)
        if selected_series is None:
            return

        self._selected_point = (series_id, index)
        self._refresh_point_styles()
        self._set_selected_point_text(self._format_selected_point_text(selected_series, index))
        rpm_value = None
        rpm_values = selected_series.get("rpm", [])
        if index < len(rpm_values):
            rpm_value = float(rpm_values[index])
        if callable(self.on_point_selected):
            self.on_point_selected(series_id, index, rpm_value)

    def _setup_viewbox_menu_hook(self) -> None:
        """Patch ViewBox context-menu so our actions are always appended last."""
        plot_item = self.plot.getPlotItem()
        vb = plot_item.vb
        original_raise = vb.raiseContextMenu
        panel = self  # captured reference

        def _append_dynamic_actions() -> None:
            # Remove any actions we added during the previous right-click.
            for action in panel._dynamic_menu_actions:
                vb.menu.removeAction(action)
            panel._dynamic_menu_actions.clear()

            identifier = panel._context_menu_identifier
            sep = vb.menu.addSeparator()
            panel._dynamic_menu_actions = [sep]

            if identifier is not None and callable(panel.on_scatter_preferences_requested):
                act = vb.menu.addAction(f'"{identifier}" Preferences...')
                # Default-argument capture so the lambda closes over the correct value.
                act.triggered.connect(lambda checked=False, ident=identifier: panel.on_scatter_preferences_requested(ident))
                panel._dynamic_menu_actions.append(act)
            if callable(panel.on_table_preferences_requested):
                act2 = vb.menu.addAction("Table Preferences...")
                act2.triggered.connect(lambda checked=False: panel.on_table_preferences_requested())
                panel._dynamic_menu_actions.append(act2)
            if callable(panel.on_log_playback_toggled):
                act_pb = vb.menu.addAction("Log Playback")
                act_pb.setCheckable(True)
                act_pb.setChecked(bool(callable(panel.get_log_playback_state) and panel.get_log_playback_state()))
                act_pb.triggered.connect(lambda checked=False, a=act_pb: panel.on_log_playback_toggled(a.isChecked()))
                panel._dynamic_menu_actions.append(act_pb)
            if callable(panel.on_plot_all_data_requested):
                act3 = vb.menu.addAction("Plot All Data")
                act3.triggered.connect(lambda checked=False: panel.on_plot_all_data_requested())
                panel._dynamic_menu_actions.append(act3)

            # If we only added the separator (no callbacks set), remove it to keep the menu clean.
            if len(panel._dynamic_menu_actions) == 1:
                vb.menu.removeAction(sep)
                panel._dynamic_menu_actions.clear()

        vb.menu.aboutToShow.connect(_append_dynamic_actions)

        def _patched_raise(ev: Any) -> Any:
            panel._context_menu_identifier = panel._find_nearest_log_identifier(ev.scenePos(), vb, plot_item)
            return original_raise(ev)

        vb.raiseContextMenu = _patched_raise

    def _find_nearest_log_identifier(self, scene_pos: Any, vb: Any, plot_item: Any) -> str | None:
        """Return the identifier of the closest visible log:: scatter point, or None if too far."""
        if not plot_item.sceneBoundingRect().contains(scene_pos):
            return None

        click_view = vb.mapSceneToView(scene_pos)
        click_rpm = float(click_view.x())
        click_val = float(click_view.y())

        view_rect = vb.viewRect()
        scene_rect = plot_item.sceneBoundingRect()
        x_per_px = max(float(view_rect.width()) / max(float(scene_rect.width()), 1.0), 1e-9)
        y_per_px = max(float(view_rect.height()) / max(float(scene_rect.height()), 1.0), 1e-9)

        nearest: tuple[float, str] | None = None
        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            if not series_id.startswith("log::"):
                continue
            if not self._is_series_visible(series_id):
                continue
            for rpm, ve in zip(series.get("rpm", []), series.get("ve", [])):
                dx_px = (float(rpm) - click_rpm) / x_per_px
                dy_px = (float(ve) - click_val) / y_per_px
                distance = dx_px ** 2 + dy_px ** 2
                if nearest is None or distance < nearest[0]:
                    nearest = (distance, series_id)

        if nearest is not None and nearest[0] <= 24.0 ** 2:
            return nearest[1][len("log::"):]
        return None

    # ------------------------------------------------------------------
    # Playback feature
    # ------------------------------------------------------------------

    def _collect_playback_time_data(self) -> None:
        """Extract time arrays from log:: series in _point_sets and prepare playback state."""
        self._playback_series_time_data.clear()
        all_times: list[float] = []
        has_time = False
        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            if not series_id.startswith("log::"):
                continue
            time_values = series.get("time")
            if not time_values or not isinstance(time_values, list):
                continue
            valid = [float(t) for t in time_values if t is not None]
            if not valid:
                continue
            has_time = True
            self._playback_series_time_data[series_id] = [float(t) for t in time_values]
            all_times.extend(valid)

        if all_times:
            all_times.sort()
            self._playback_all_times = all_times
            self._playback_max_time = all_times[-1]
            self._playback_start_time = all_times[0]
        else:
            self._playback_all_times = []
            self._playback_max_time = 0.0
            self._playback_start_time = 0.0

        self._has_playback_time_data = has_time
        self._refresh_playback_widget_visibility()
        self.btn_play.setEnabled(has_time)
        self.btn_pause.setEnabled(has_time)
        self.btn_stop.setEnabled(has_time)
        self.btn_speed.setEnabled(has_time)
        self.playback_slider.setEnabled(has_time)
        if has_time:
            duration = self._playback_max_time - self._playback_start_time
            self.playback_time_label.setText(f"0.0s / {duration:.1f}s")
            self.playback_slider.setValue(0)
        else:
            self.playback_time_label.setText("0.0s / 0.0s")
        self._update_playback_button_states()

    def _on_playback_play(self) -> None:
        if not self._playback_all_times:
            return
        if self._playback_state == "stopped":
            # If a scatter point is selected, start from its time
            start_offset = self._selected_point_playback_time()
            self._playback_elapsed = start_offset
            self._playback_current_wall_start = _time_mod.monotonic()
            self._apply_playback_at_elapsed(self._playback_elapsed)
            self._playback_state = "playing"
            self._playback_timer.start()
        elif self._playback_state == "paused":
            self._playback_current_wall_start = _time_mod.monotonic()
            self._playback_state = "playing"
            self._playback_timer.start()
        self._update_playback_button_states()

    def _on_playback_pause(self) -> None:
        if self._playback_state == "playing":
            self._playback_timer.stop()
            wall_delta = _time_mod.monotonic() - self._playback_current_wall_start
            self._playback_elapsed += wall_delta * self._playback_speed
            self._playback_state = "paused"
            self._update_playback_button_states()

    def _on_playback_stop(self) -> None:
        self._playback_timer.stop()
        self._playback_state = "stopped"
        self._playback_elapsed = 0.0
        self._speed_index = 0
        self._playback_speed = self._speed_steps[0]
        self.btn_speed.setText("1x")
        self.playback_slider.blockSignals(True)
        self.playback_slider.setValue(0)
        self.playback_slider.blockSignals(False)
        duration = self._playback_max_time - self._playback_start_time
        self.playback_time_label.setText(f"0.0s / {duration:.1f}s")
        # Restore full scatter data
        self._refresh_point_styles()
        self._update_playback_button_states()

    def _on_playback_speed_toggle(self) -> None:
        # Snapshot the current elapsed time so the speed change is seamless
        if self._playback_state == "playing":
            wall_delta = _time_mod.monotonic() - self._playback_current_wall_start
            self._playback_elapsed += wall_delta * self._playback_speed
            self._playback_current_wall_start = _time_mod.monotonic()
        self._speed_index = (self._speed_index + 1) % len(self._speed_steps)
        self._playback_speed = self._speed_steps[self._speed_index]
        label = f"{int(self._playback_speed)}x" if self._playback_speed == int(self._playback_speed) else f"{self._playback_speed}x"
        self.btn_speed.setText(label)

    def _playback_tick(self) -> None:
        wall_delta = _time_mod.monotonic() - self._playback_current_wall_start
        elapsed = self._playback_elapsed + wall_delta * self._playback_speed
        duration = self._playback_max_time - self._playback_start_time
        if elapsed >= duration:
            elapsed = duration
            self._playback_timer.stop()
            self._playback_elapsed = elapsed
            self._playback_state = "paused"
            self._update_playback_button_states()
        self._apply_playback_at_elapsed(elapsed)

    def _apply_playback_at_elapsed(self, elapsed: float) -> None:
        """Show only scatter points whose time <= start_time + elapsed."""
        cutoff = self._playback_start_time + elapsed
        duration = self._playback_max_time - self._playback_start_time

        # Update slider without re-entrancy
        self.playback_slider.blockSignals(True)
        if duration > 0:
            self.playback_slider.setValue(int(elapsed / duration * 1000))
        else:
            self.playback_slider.setValue(1000)
        self.playback_slider.blockSignals(False)
        self.playback_time_label.setText(f"{elapsed:.1f}s / {duration:.1f}s")

        # Filter each log scatter series to show only points at or before cutoff
        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            if not series_id.startswith("log::"):
                continue
            item = self.dynamic_scatter_items.get(series_id)
            if item is None:
                continue
            if not self._is_series_visible(series_id):
                continue
            time_values = self._playback_series_time_data.get(series_id)
            if not time_values:
                continue
            rpm_all = series.get("rpm", [])
            ve_all = series.get("ve", [])
            rpm_visible = []
            ve_visible = []
            for i, t in enumerate(time_values):
                if t <= cutoff and i < len(rpm_all) and i < len(ve_all):
                    rpm_visible.append(rpm_all[i])
                    ve_visible.append(ve_all[i])
            item.setData(x=rpm_visible, y=ve_visible)

        # Show leading-edge marker on the most recent point
        self._show_playback_leading_edge(cutoff)

    def _show_playback_leading_edge(self, cutoff: float) -> None:
        """Highlight the most-recently-revealed scatter point and show merged summary."""
        best_time = -1.0
        best_rpm: float | None = None
        best_ve: float | None = None

        # Collect the leading-edge index per series for the merged summary
        leading_edges: list[tuple[str, dict[str, Any], int, float]] = []  # (series_id, series, index, time)

        for series_id, time_values in self._playback_series_time_data.items():
            if not self._is_series_visible(series_id):
                continue
            series = self._get_series(series_id)
            if series is None:
                continue
            rpm_all = series.get("rpm", [])
            ve_all = series.get("ve", [])
            local_best_t = -1.0
            local_best_i = -1
            for i, t in enumerate(time_values):
                if t <= cutoff and t > local_best_t and i < len(rpm_all) and i < len(ve_all):
                    local_best_t = t
                    local_best_i = i
            if local_best_i >= 0:
                leading_edges.append((series_id, series, local_best_i, local_best_t))
                if local_best_t > best_time:
                    best_time = local_best_t
                    best_rpm = rpm_all[local_best_i]
                    best_ve = ve_all[local_best_i]

        if best_rpm is not None and best_ve is not None:
            self.selected_marker.setData([float(best_rpm)], [float(best_ve)])
            self.selected_marker.show()
        else:
            self.selected_marker.setData([], [])
            self.selected_marker.hide()

        # Build merged point summary
        if leading_edges:
            self._set_selected_point_text(self._format_playback_summary(leading_edges, cutoff))
        else:
            self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")

    def _format_playback_summary(
        self,
        leading_edges: list[tuple[str, dict[str, Any], int, float]],
        cutoff: float,
    ) -> str:
        """Build a merged point summary for all series at their leading-edge indices."""
        elapsed = cutoff - self._playback_start_time
        detail_rows: list[tuple[str, str]] = [("Time:", f"{elapsed:.2f}s")]
        seen_labels: set[str] = set()

        for series_id, series, index, _t in leading_edges:
            name = str(series.get("name", series_id))
            y_values = series.get("ve")
            if isinstance(y_values, list) and 0 <= index < len(y_values):
                try:
                    val = float(y_values[index])
                    val = val if val == val else None
                except Exception:
                    val = None
            else:
                val = None
            label = f"{name}:"
            if label not in seen_labels:
                detail_rows.append((label, self._format_value(val)))
                seen_labels.add(label)

            rpm_all = series.get("rpm", [])
            if 0 <= index < len(rpm_all):
                rpm_label = "RPM:"
                if rpm_label not in seen_labels:
                    detail_rows.append((rpm_label, self._format_value(float(rpm_all[index]))))
                    seen_labels.add(rpm_label)

            for channel_dict_key in ("extra_channels", "detail_channels"):
                channels = series.get(channel_dict_key, {})
                if not isinstance(channels, dict):
                    continue
                for ch_name in sorted(channels.keys()):
                    ch_label = f"{ch_name}:"
                    if ch_label in seen_labels:
                        continue
                    ch_values = channels.get(ch_name)
                    if isinstance(ch_values, list) and 0 <= index < len(ch_values):
                        detail_rows.append((ch_label, self._format_value(ch_values[index])))
                        seen_labels.add(ch_label)

            count_channels = series.get("count_channels", {})
            if isinstance(count_channels, dict):
                for ch_name in sorted(count_channels.keys()):
                    ch_label = f"{ch_name}:"
                    if ch_label in seen_labels:
                        continue
                    ch_values = count_channels.get(ch_name)
                    if isinstance(ch_values, list) and 0 <= index < len(ch_values):
                        cnt = ch_values[index]
                        detail_rows.append((ch_label, str(int(cnt)) if cnt is not None else "n/a"))
                        seen_labels.add(ch_label)

        rows_html = "".join(
            f"<tr><td style='padding-right:10px'>{label}</td><td>{value}</td></tr>"
            for label, value in detail_rows
        )
        return (
            f"<span style='color:#f0f0f0'>"
            f"<b>Playback Point Summary:</b><br>"
            f"<table cellpadding='0' cellspacing='2'>{rows_html}</table>"
            f"</span>"
        )

    def _selected_point_playback_time(self) -> float:
        """Return the elapsed offset for the currently selected point, or 0."""
        if self._selected_point is None:
            return 0.0
        series_id, index = self._selected_point
        time_values = self._playback_series_time_data.get(series_id)
        if not time_values or index < 0 or index >= len(time_values):
            return 0.0
        return max(0.0, float(time_values[index]) - self._playback_start_time)

    def _on_slider_pressed(self) -> None:
        self._slider_dragging = True
        if self._playback_state == "playing":
            self._playback_timer.stop()

    def _on_slider_released(self) -> None:
        self._slider_dragging = False
        duration = self._playback_max_time - self._playback_start_time
        frac = self.playback_slider.value() / 1000.0
        self._playback_elapsed = frac * duration
        if self._playback_state == "playing":
            self._playback_current_wall_start = _time_mod.monotonic()
            self._playback_timer.start()

    def _on_slider_value_changed(self, value: int) -> None:
        if not self._slider_dragging:
            return
        duration = self._playback_max_time - self._playback_start_time
        elapsed = (value / 1000.0) * duration
        self._apply_playback_at_elapsed(elapsed)

    def _is_playback_active(self) -> bool:
        return self._playback_state in ("playing", "paused")


class RowEditorTableWidget(QTableWidget):
    """Single-row editor with arrow-key navigation and wheel-based value changes."""

    def __init__(self) -> None:
        super().__init__()
        self.on_adjust_value: Any = None
        self.on_undo: Any = None
        self.on_redo: Any = None
        self.on_focus_changed: Any = None
        self.setRowCount(1)
        self.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setSectionsClickable(True)
        self.horizontalHeader().sectionClicked.connect(self._on_header_section_clicked)

    def _on_header_section_clicked(self, section: int) -> None:
        if 0 <= section < self.columnCount():
            self.setCurrentCell(0, section)
            self.setFocus()

    def focusInEvent(self, event) -> None:  # type: ignore[override]
        super().focusInEvent(event)
        if callable(self.on_focus_changed):
            self.on_focus_changed(True)

    def focusOutEvent(self, event) -> None:  # type: ignore[override]
        super().focusOutEvent(event)
        if callable(self.on_focus_changed):
            self.on_focus_changed(False)

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        current_column = self.currentColumn()
        if event.matches(QKeySequence.StandardKey.Undo):
            if callable(self.on_undo):
                self.on_undo(current_column)
            return
        if event.matches(QKeySequence.StandardKey.Redo):
            if callable(self.on_redo):
                self.on_redo(current_column)
            return
        if event.key() == Qt.Key.Key_Left:
            if current_column > 0:
                self.setCurrentCell(0, current_column - 1)
            return
        if event.key() == Qt.Key.Key_Right:
            if current_column < self.columnCount() - 1:
                self.setCurrentCell(0, current_column + 1)
            return
        if event.key() == Qt.Key.Key_Up:
            if callable(self.on_adjust_value) and current_column >= 0:
                self.on_adjust_value(current_column, 0.1)
            return
        if event.key() == Qt.Key.Key_Down:
            if callable(self.on_adjust_value) and current_column >= 0:
                self.on_adjust_value(current_column, -0.1)
            return
        super().keyPressEvent(event)

    def wheelEvent(self, event) -> None:  # type: ignore[override]
        current_column = self.currentColumn()
        if not callable(self.on_adjust_value) or current_column < 0:
            super().wheelEvent(event)
            return
        delta_x = event.angleDelta().x()
        delta = event.angleDelta().y()
        if abs(delta_x) > abs(delta) or (event.modifiers() & Qt.KeyboardModifier.ShiftModifier):
            horizontal_delta = delta_x if delta_x != 0 else delta
            if horizontal_delta == 0:
                super().wheelEvent(event)
                return
            step = 1 if horizontal_delta > 0 else -1
            target_column = max(0, min(self.columnCount() - 1, current_column + step))
            self.setCurrentCell(0, target_column)
            event.accept()
            return
        if delta == 0:
            super().wheelEvent(event)
            return
        steps = max(1, abs(delta) // 120)
        increment = 0.5 if (event.modifiers() & Qt.KeyboardModifier.ControlModifier) else 0.1
        amount = float(steps * increment if delta > 0 else -steps * increment)
        self.on_adjust_value(current_column, amount)
        event.accept()


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
        custom_group = QGroupBox("Custom Identifiers")
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
        tabs.addTab(custom_tab, "Custom Identifiers")

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

        for table_name, table_prefs in list(self._working_preferences.items()):
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


class StartupTuneDialog(QDialog):
    """Startup picker for quickly loading a recent tune."""

    def __init__(
        self,
        recent_tune_files: list[Path],
        window_icon: QIcon,
        default_browse_dir: Path | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Gyatt-O-Tune")
        self.setModal(True)
        self.resize(620, 460)

        self.selected_tune_path: Path | None = None
        self.default_browse_dir = default_browse_dir
        has_recent_tunes = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.setStyleSheet(
            "QDialog { background: #14161c; }"
            "QLabel { color: #e7ebf3; }"
            "QListWidget { background: #1c1f27; color: #f0f3fa; border: 1px solid #2c3240; border-radius: 8px; }"
            "QPushButton { background: #2d5ecf; color: white; border-radius: 6px; padding: 7px 12px; }"
            "QPushButton:hover { background: #3873f0; }"
            "QPushButton:disabled { background: #4a5060; color: #c1c7d8; }"
        )

        icon_label = QLabel()
        if not window_icon.isNull():
            icon_label.setPixmap(window_icon.pixmap(82, 82))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(icon_label)

        title = QLabel("Gyatt-O-Tune")
        title.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        title.setStyleSheet("font-size: 24px; font-weight: 700;")
        layout.addWidget(title)

        subtitle = QLabel("Open a recent tune or browse to select a tune file.")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #c7cedd;")
        layout.addWidget(subtitle)

        recent_title = QLabel("Recent Tunes")
        recent_title.setStyleSheet("font-size: 13px; font-weight: 600;")
        layout.addWidget(recent_title)

        self.recent_list = QListWidget()
        self.recent_list.setAlternatingRowColors(True)
        for file_path in recent_tune_files:
            if not file_path.exists():
                continue
            item = QListWidgetItem(file_path.name)
            item.setData(Qt.ItemDataRole.UserRole, str(file_path))
            item.setToolTip(str(file_path))
            self.recent_list.addItem(item)
            has_recent_tunes = True

        if not has_recent_tunes:
            placeholder = QListWidgetItem("No recent tune files found")
            placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
            self.recent_list.addItem(placeholder)
        else:
            self.recent_list.setCurrentRow(0)

        self.recent_list.itemDoubleClicked.connect(self._open_selected)
        self.recent_list.currentItemChanged.connect(self._update_open_button_state)
        layout.addWidget(self.recent_list, 1)

        button_row = QHBoxLayout()
        button_row.setSpacing(12)
        self.open_button = QPushButton("Open")
        self.open_button.clicked.connect(self._open_selected)
        self.open_button.setEnabled(has_recent_tunes)
        self.open_button.setMinimumHeight(38)
        button_row.addWidget(self.open_button, 1)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self._browse_for_tune)
        browse_button.setMinimumHeight(38)
        button_row.addWidget(browse_button, 1)

        layout.addLayout(button_row)

    def _update_open_button_state(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        path_str = current.data(Qt.ItemDataRole.UserRole) if current is not None else None
        self.open_button.setEnabled(isinstance(path_str, str))

    def _browse_for_tune(self) -> None:
        start_dir = str(self.default_browse_dir) if self.default_browse_dir is not None else ""
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open MegaSquirt Tune File",
            start_dir,
            "Tune Files (*.msq *.ini *.txt);;All Files (*.*)",
        )
        if not selected_path:
            return

        self.selected_tune_path = Path(selected_path)
        self.accept()

    def _open_selected(self, _item: QListWidgetItem | None = None) -> None:
        current_item = self.recent_list.currentItem()
        if current_item is None:
            QMessageBox.information(self, "Select a Tune", "Choose a recent tune file to continue.")
            return

        path_str = current_item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(path_str, str):
            return

        selected = Path(path_str)
        if not selected.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {selected} no longer exists.")
            return

        self.selected_tune_path = selected
        self.accept()


class MainWindow(QMainWindow):
    MAX_RECENT_FILES = 10
    WINDOW_GEOMETRY_KEY = "main_window_geometry"
    WINDOW_DOCK_STATE_KEY = "main_window_dock_state_v2"
    DEFAULT_WINDOW_GEOMETRY_KEY = "default_main_window_geometry_v1"
    DEFAULT_WINDOW_DOCK_STATE_KEY = "default_main_window_dock_state_v1"
    OPEN_TUNE_PLACEHOLDER_KEY = "__open_tune_file__"
    ROW_VIZ_SERIES_BY_TABLE_TYPE: dict[str, list[str]] = {
        "ve": [],
        "afr": [],
        "knock": [],
        "generic": [],
    }

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Gyatt-O-Tune")
        self.resize(1200, 800)

        self.tune_loader = TuneLoader()
        self.log_loader = LogLoader()

        self.tune_data: TuneData | None = None
        self.loaded_tune_path: Path | None = None
        self.save_target_path: Path | None = None
        self.save_tune_as_action: Any | None = None
        self.log_file: Path | None = None
        self.log_df: Any | None = None
        self.current_table: TableData | None = None
        self.current_x_axis: AxisVector | None = None
        self.current_y_axis: AxisVector | None = None
        self.selected_table_row_idx: int | None = None
        self.pending_row_values: list[float] = []
        self.row_default_values: list[float] = []
        self.row_edit_undo_stack: list[list[float]] = []
        self._active_row_edit_column: int | None = None
        self._active_row_edit_snapshot: list[float] | None = None
        self.global_undo_stack: list[dict[str, Any]] = []
        self.global_redo_stack: list[dict[str, Any]] = []
        self._active_global_cell_edit: dict[str, Any] | None = None
        self._history_is_applying = False
        self.undo_action: QAction | None = None
        self.redo_action: QAction | None = None
        self._last_row_viz_dataset_key: tuple[Any, ...] | None = None
        self.loaded_table_snapshots: dict[str, list[list[float]]] = {}
        self.selected_rows_per_table: dict[str, int] = {}  # Track selected row for each table
        self.pending_edits_per_table: dict[str, dict] = {}  # Track unsaved edits by table and row
        self._row_viz_refresh_timer = QTimer(self)
        self._row_viz_refresh_timer.setSingleShot(True)
        self._row_viz_refresh_timer.timeout.connect(self._apply_deferred_row_visualization_update)

        self.recent_tune_files: list[Path] = []
        self.recent_log_files: list[Path] = []
        self.working_folder: Path | None = None

        self.favorite_tables: set[str] = set()
        self.only_show_favorited_tables = False
        self.show_1d_tables = True
        self.show_2d_tables = True
        self.show_log_playback = True
        self.show_tunerstudio_names = True
        self.row_viz_preferences: dict[str, dict[str, bool]] = self._default_row_viz_preferences()
        self.table_log_channel_preferences: dict[str, dict[str, dict[str, bool]]] = {}

        self._load_recent_files()
        self._load_favorites()
        self._load_table_filter_preferences()
        self._load_row_viz_preferences()
        self._load_table_log_channel_preferences()
        self._load_working_folder()

        self._create_menu()
        self._create_layout()
        self.setStatusBar(QStatusBar())
        QTimer.singleShot(0, self._restore_dock_layout)
        self.statusBar().showMessage("Ready")

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        QTimer.singleShot(0, self._auto_size_table_and_row_viz)

    def _create_menu(self) -> None:
        file_menu = self.menuBar().addMenu("&File")

        open_tune_action = file_menu.addAction("Open Tune...")
        open_tune_action.triggered.connect(self._open_tune_file)

        open_log_action = file_menu.addAction("Open Log...")
        open_log_action.triggered.connect(self._open_log_file)

        self.recent_tunes_menu = file_menu.addMenu("Recent Tunes")
        self._update_recent_tunes_menu()

        load_recent_tune_shortcut_action = QAction(self)
        load_recent_tune_shortcut_action.setShortcut(QKeySequence("Ctrl+T"))
        load_recent_tune_shortcut_action.triggered.connect(self._on_load_most_recent_tune_shortcut)
        self.addAction(load_recent_tune_shortcut_action)

        load_matching_log_shortcut_action = QAction(self)
        load_matching_log_shortcut_action.setShortcut(QKeySequence("Ctrl+L"))
        load_matching_log_shortcut_action.triggered.connect(self._on_load_matching_log_shortcut)
        self.addAction(load_matching_log_shortcut_action)

        self.recent_logs_menu = file_menu.addMenu("Recent Logs")
        self._update_recent_logs_menu()

        file_menu.addSeparator()

        self.save_tune_as_action = file_menu.addAction("Save Tune &As...")
        self.save_tune_as_action.setShortcut(QKeySequence("Ctrl+S"))
        self.save_tune_as_action.triggered.connect(self._save_tune_as)
        self.save_tune_as_action.setEnabled(False)

        file_menu.addSeparator()
        preferences_action = file_menu.addAction("&Preferences...")
        preferences_action.triggered.connect(self._open_preferences)

        file_menu.addSeparator()
        exit_action = file_menu.addAction("E&xit")
        exit_action.triggered.connect(self.close)

        edit_menu = self.menuBar().addMenu("&Edit")
        self.undo_action = edit_menu.addAction("&Undo")
        self.undo_action.setShortcut(QKeySequence("Ctrl+Z"))
        self.undo_action.triggered.connect(self._trigger_global_undo)
        self.undo_action.setEnabled(False)

        self.redo_action = edit_menu.addAction("&Redo")
        self.redo_action.setShortcut(QKeySequence("Ctrl+Shift+Z"))
        self.redo_action.triggered.connect(self._trigger_global_redo)
        self.redo_action.setEnabled(False)

        self.view_menu = self.menuBar().addMenu("&View")

        self.window_layout_menu = self.view_menu.addMenu("Window Layout")

        save_layout_action = self.window_layout_menu.addAction("Save")
        save_layout_action.triggered.connect(self._save_default_window_layout)

        load_layout_action = self.window_layout_menu.addAction("Load")
        load_layout_action.triggered.connect(self._load_default_window_layout)

        reset_layout_action = self.window_layout_menu.addAction("Reset")
        reset_layout_action.triggered.connect(self._reset_window_layout)
        self.window_layout_menu.ensurePolished()
        self.window_layout_menu.setMinimumWidth(self.window_layout_menu.sizeHint().width())

        self.windows_menu = self.view_menu.addMenu("Visible Windows")

    def _set_tune_save_actions_enabled(self, enabled: bool) -> None:
        if self.save_tune_as_action is not None:
            self.save_tune_as_action.setEnabled(enabled)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._save_dock_layout()
        try:
            self._save_row_viz_preferences()
            self._save_table_log_channel_preferences()
        except Exception:
            # Do not block app shutdown on preference write failures.
            pass
        super().closeEvent(event)

    def _restore_dock_layout(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        default_geometry = settings.value(self.DEFAULT_WINDOW_GEOMETRY_KEY)
        default_dock_state = settings.value(self.DEFAULT_WINDOW_DOCK_STATE_KEY)
        geometry = settings.value(self.WINDOW_GEOMETRY_KEY)
        dock_state = settings.value(self.WINDOW_DOCK_STATE_KEY)

        restored = False
        if default_geometry is not None and default_dock_state is not None:
            restored = bool(self.restoreGeometry(default_geometry) and self.restoreState(default_dock_state))

        if not restored and geometry is not None and dock_state is not None:
            restored = bool(self.restoreGeometry(geometry) and self.restoreState(dock_state))

        if not restored:
            self._apply_standard_window_layout()

        self.setTabPosition(Qt.DockWidgetArea.AllDockWidgetAreas, QTabWidget.TabPosition.North)

    def _save_dock_layout(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue(self.WINDOW_GEOMETRY_KEY, self.saveGeometry())
        settings.setValue(self.WINDOW_DOCK_STATE_KEY, self.saveState())

    def _apply_standard_window_layout(self) -> None:
        self.resize(1200, 800)
        self.setTabPosition(Qt.DockWidgetArea.AllDockWidgetAreas, QTabWidget.TabPosition.North)

        for dock in [self.tune_tables_dock, self.selected_table_dock, self.row_viz_dock, self.row_editor_dock]:
            dock.setFloating(False)
            dock.show()

        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.tune_tables_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.selected_table_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_viz_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_editor_dock)

        self.tabifyDockWidget(self.selected_table_dock, self.row_viz_dock)
        self.tabifyDockWidget(self.selected_table_dock, self.row_editor_dock)
        self.selected_table_dock.raise_()
        self.resizeDocks([self.tune_tables_dock, self.selected_table_dock], [320, 920], Qt.Orientation.Horizontal)

    def _save_default_window_layout(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue(self.DEFAULT_WINDOW_GEOMETRY_KEY, self.saveGeometry())
        settings.setValue(self.DEFAULT_WINDOW_DOCK_STATE_KEY, self.saveState())
        self.statusBar().showMessage("Saved current layout as default.", 3000)

    def _load_default_window_layout(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        geometry = settings.value(self.DEFAULT_WINDOW_GEOMETRY_KEY)
        dock_state = settings.value(self.DEFAULT_WINDOW_DOCK_STATE_KEY)

        if geometry is None or dock_state is None:
            self._apply_standard_window_layout()
            self.statusBar().showMessage("No saved default layout found. Applied standard layout.", 4000)
            return

        geometry_ok = self.restoreGeometry(geometry)
        dock_ok = self.restoreState(dock_state)
        self.setTabPosition(Qt.DockWidgetArea.AllDockWidgetAreas, QTabWidget.TabPosition.North)
        if not geometry_ok or not dock_ok:
            self._apply_standard_window_layout()
            self.statusBar().showMessage("Saved default layout was invalid. Applied standard layout.", 4000)
            return

        self.statusBar().showMessage("Loaded default window layout.", 3000)

    def _reset_window_layout(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.remove(self.WINDOW_GEOMETRY_KEY)
        settings.remove(self.WINDOW_DOCK_STATE_KEY)

        self._apply_standard_window_layout()
        self.statusBar().showMessage("Window layout reset to standard defaults. Saved default layout was preserved.", 4000)

    def _create_section_dock(self, title: str, object_name: str, content: QWidget) -> QDockWidget:
        dock = QDockWidget(title, self)
        dock.setObjectName(object_name)
        dock.setAllowedAreas(Qt.DockWidgetArea.AllDockWidgetAreas)
        dock.setFeatures(
            QDockWidget.DockWidgetFeature.DockWidgetMovable
            | QDockWidget.DockWidgetFeature.DockWidgetFloatable
            | QDockWidget.DockWidgetFeature.DockWidgetClosable
        )
        dock.setWidget(content)
        return dock

    def _create_outlined_panel(self, content: QWidget) -> QGroupBox:
        panel = QGroupBox("")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.addWidget(content)
        return panel

    def _set_row_editor_button_visibility(
        self,
        show_generate: bool,
        show_write: bool,
        show_revert: bool,
    ) -> None:
        _ = show_generate
        _ = show_write
        self.revert_row_changes_button.setVisible(show_revert)

    def _refresh_history_action_state(self) -> None:
        if self.undo_action is not None:
            self.undo_action.setEnabled(bool(self.global_undo_stack) or self._active_global_cell_edit is not None)
        if self.redo_action is not None:
            self.redo_action.setEnabled(bool(self.global_redo_stack))

    def _clear_global_history(self) -> None:
        self.global_undo_stack.clear()
        self.global_redo_stack.clear()
        self._active_global_cell_edit = None
        self._refresh_history_action_state()

    def _flush_active_global_cell_edit(self) -> None:
        pending = self._active_global_cell_edit
        if pending is None:
            return
        self._active_global_cell_edit = None

        if self._history_is_applying:
            self._refresh_history_action_state()
            return

        old_value = float(pending.get("old", 0.0))
        new_value = float(pending.get("new", 0.0))
        if abs(new_value - old_value) <= 1e-9:
            self._refresh_history_action_state()
            return

        self.global_undo_stack.append(
            {
                "table": str(pending.get("table", "")),
                "row": int(pending.get("row", -1)),
                "col": int(pending.get("col", -1)),
                "old": old_value,
                "new": new_value,
            }
        )
        if len(self.global_undo_stack) > 5000:
            self.global_undo_stack = self.global_undo_stack[-5000:]
        self._refresh_history_action_state()

    def _current_editor_cell_coordinates(self, column_index: int) -> tuple[int, int] | None:
        if self.current_table is None:
            return None
        if column_index < 0:
            return None

        if self.current_table.rows == 1:
            if column_index >= self.current_table.cols:
                return None
            return (0, column_index)

        if self.current_table.cols == 1:
            if column_index >= self.current_table.rows:
                return None
            return (column_index, 0)

        if self.selected_table_row_idx is None:
            return None
        if column_index >= self.current_table.cols:
            return None
        return (self.selected_table_row_idx, column_index)

    def _record_global_cell_edit(
        self,
        table_name: str,
        row_index: int,
        column_index: int,
        old_value: float,
        new_value: float,
    ) -> None:
        if self._history_is_applying:
            return
        if abs(float(new_value) - float(old_value)) <= 1e-9:
            return

        edit_key = (table_name, int(row_index), int(column_index))
        pending = self._active_global_cell_edit
        pending_key = None
        if pending is not None:
            pending_key = (
                str(pending.get("table", "")),
                int(pending.get("row", -1)),
                int(pending.get("col", -1)),
            )

        if pending is None:
            self._active_global_cell_edit = {
                "table": table_name,
                "row": int(row_index),
                "col": int(column_index),
                "old": float(old_value),
                "new": float(new_value),
            }
        elif pending_key == edit_key:
            pending["new"] = float(new_value)
        else:
            self._flush_active_global_cell_edit()
            self._active_global_cell_edit = {
                "table": table_name,
                "row": int(row_index),
                "col": int(column_index),
                "old": float(old_value),
                "new": float(new_value),
            }

        # A new edit invalidates redo history immediately.
        self.global_redo_stack.clear()
        self._refresh_history_action_state()

    def _find_table_list_item(self, table_name: str) -> QListWidgetItem | None:
        for idx in range(self.table_list.count()):
            item = self.table_list.item(idx)
            if item is None:
                continue
            if item.data(Qt.ItemDataRole.UserRole) == table_name:
                return item
        return None

    def _set_staged_cell_value(self, table_name: str, row_index: int, column_index: int, value: float) -> None:
        if self.tune_data is None:
            return
        table = self.tune_data.tables.get(table_name)
        if table is None:
            return
        if row_index < 0 or row_index >= table.rows or column_index < 0 or column_index >= table.cols:
            return

        baseline_cell_value = float(table.values[row_index][column_index])
        table_state = self.pending_edits_per_table.setdefault(
            table_name,
            {"rows": {}, "one_d": None},
        )

        if table.rows == 1:
            baseline_row = [float(v) for v in table.values[0]]
            working_row = table_state.get("one_d")
            if not isinstance(working_row, list) or len(working_row) != table.cols:
                working_row = [float(v) for v in baseline_row]
            working_row[column_index] = float(value)
            table_state["one_d"] = None if not self._values_differ(working_row, baseline_row) else [float(v) for v in working_row]
        elif table.cols == 1:
            baseline_values = [float(r[0]) for r in table.values]
            working_values = table_state.get("one_d")
            if not isinstance(working_values, list) or len(working_values) != table.rows:
                working_values = [float(v) for v in baseline_values]
            working_values[row_index] = float(value)
            table_state["one_d"] = None if not self._values_differ(working_values, baseline_values) else [float(v) for v in working_values]
        else:
            rows_state = table_state.setdefault("rows", {})
            baseline_row = [float(v) for v in table.values[row_index]]
            working_row = rows_state.get(row_index)
            if not isinstance(working_row, list) or len(working_row) != table.cols:
                working_row = [float(v) for v in baseline_row]
            working_row[column_index] = float(value)
            if self._values_differ(working_row, baseline_row):
                rows_state[row_index] = [float(v) for v in working_row]
            else:
                rows_state.pop(row_index, None)

        if abs(float(value) - baseline_cell_value) <= 1e-9:
            # Ensure no stray row override remains for the exact baseline cell.
            self._stage_current_row_edits()

        if not table_state.get("rows") and table_state.get("one_d") is None:
            self.pending_edits_per_table.pop(table_name, None)

    def _navigate_to_table_cell(self, table_name: str, row_index: int, column_index: int) -> bool:
        item = self._find_table_list_item(table_name)
        if item is None:
            return False

        self.table_list.setCurrentItem(item)
        if self.current_table is None:
            return False

        if self.current_table.rows == 1:
            preferred_column = column_index
            self._load_selected_table_row(0, preferred_column=preferred_column)
            return True

        if self.current_table.cols == 1:
            preferred_column = row_index
            self._load_selected_table_row(row_index, preferred_column=preferred_column)
            return True

        self._load_selected_table_row(row_index, preferred_column=column_index)
        return True

    def _apply_history_entry(self, entry: dict[str, Any], use_new_value: bool) -> bool:
        if self.tune_data is None:
            return False

        table_name = str(entry.get("table", ""))
        row_index = int(entry.get("row", -1))
        column_index = int(entry.get("col", -1))
        target_value = float(entry.get("new" if use_new_value else "old", 0.0))
        if not table_name:
            return False

        self._history_is_applying = True
        try:
            self._set_staged_cell_value(table_name, row_index, column_index, target_value)
            navigated = self._navigate_to_table_cell(table_name, row_index, column_index)
            if not navigated:
                return False

            editor_column = column_index
            if self.current_table is not None and self.current_table.cols == 1:
                editor_column = row_index
            if 0 <= editor_column < len(self.pending_row_values):
                self.pending_row_values[editor_column] = float(target_value)
                self._refresh_table_row_editor(preferred_column=editor_column)
                self._update_table_grid_row_visualization(include_row_plot=False)
                self._schedule_table_grid_row_visualization_update()
            return True
        finally:
            self._history_is_applying = False

    def _trigger_global_undo(self) -> None:
        self._flush_active_global_cell_edit()
        if not self.global_undo_stack:
            return
        entry = self.global_undo_stack.pop()
        applied = self._apply_history_entry(entry, use_new_value=False)
        if applied:
            self.global_redo_stack.append(entry)
        self._refresh_history_action_state()

    def _trigger_global_redo(self) -> None:
        self._flush_active_global_cell_edit()
        if not self.global_redo_stack:
            return
        entry = self.global_redo_stack.pop()
        applied = self._apply_history_entry(entry, use_new_value=True)
        if applied:
            self.global_undo_stack.append(entry)
        self._refresh_history_action_state()

    def _apply_row_editor_selection_highlight(self) -> None:
        if not self.pending_row_values:
            return

        minimum, maximum = self._matrix_min_max([self.pending_row_values])
        span = max(maximum - minimum, 1e-9)
        selected_column = self.table_row_editor.currentColumn()

        for column_index, value in enumerate(self.pending_row_values):
            item = self.table_row_editor.item(0, column_index)
            if item is None:
                continue
            if column_index == selected_column:
                item.setBackground(QColor(220, 70, 70, 200))
            else:
                item.setBackground(self._cell_color(value, minimum, span))

    def _sync_editor_selection_to_row_plot(self) -> None:
        if not self.pending_row_values or self.table_row_editor.columnCount() == 0:
            self.table_row_panel.clear_selected_point()
            return

        column_index = self.table_row_editor.currentColumn()
        if 0 <= column_index < self.table_row_editor.columnCount():
            self.table_row_panel.select_table_point(column_index, emit_callback=False)
        else:
            self.table_row_panel.clear_selected_point()

    def _on_row_editor_current_cell_changed(
        self,
        current_row: int,
        current_column: int,
        previous_row: int,
        previous_column: int,
    ) -> None:
        if current_column != previous_column:
            self._flush_active_global_cell_edit()
        _ = (current_row, current_column, previous_row, previous_column)
        self._apply_row_editor_selection_highlight()
        self._sync_editor_selection_to_row_plot()

    def _on_row_editor_focus_changed(self, has_focus: bool) -> None:
        _ = has_focus
        self._apply_row_editor_selection_highlight()
        self._sync_editor_selection_to_row_plot()

    def _on_load_most_recent_tune_shortcut(self) -> None:
        """Ctrl+T: load most recent tune from working folder or recent files."""
        most_recent = self._most_recent_tune_file()
        if most_recent is None:
            self.statusBar().showMessage("No recent tune file found to load.", 5000)
            return

        # Show a confirmation popup when a working folder is set
        if self.working_folder and most_recent.parent == self.working_folder:
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Question)
            msg_box.setWindowTitle("Load Tune File")
            msg_box.setText(f"Load most recent tune?\n\n{most_recent.name}")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)
            if msg_box.exec() != QMessageBox.StandardButton.Yes:
                return

        self._open_recent_tune_file(most_recent)

    def _on_load_matching_log_shortcut(self) -> None:
        """Ctrl+L: load a .msl or .mlg file with the same name as the loaded tune."""
        if self.tune_data is None or self.loaded_tune_path is None:
            QMessageBox.warning(
                self,
                "No Tune Loaded",
                "Please load a tune file first before loading a matching log."
            )
            return

        tune_file = self.loaded_tune_path
        if tune_file.suffix.lower() != ".msq":
            QMessageBox.warning(
                self,
                "Invalid Tune File",
                "The loaded tune must be an .msq file to find a matching log."
            )
            return

        tune_stem = tune_file.stem.lower()
        matching_logs: list[Path] = []
        seen: set[Path] = set()

        # 1) Prefer exact same-name logs next to the tune file.
        for suffix in (".msl", ".mlg"):
            candidate = tune_file.with_suffix(suffix)
            if candidate.exists():
                resolved = candidate.resolve()
                if resolved not in seen:
                    matching_logs.append(resolved)
                    seen.add(resolved)

        # 2) Include same-stem logs from recent log history.
        for recent_log in self.recent_log_files:
            resolved = recent_log.resolve()
            if not resolved.exists():
                continue
            if resolved.suffix.lower() not in {".msl", ".mlg"}:
                continue
            if resolved.stem.lower() != tune_stem:
                continue
            if resolved not in seen:
                matching_logs.append(resolved)
                seen.add(resolved)

        # 3) Include same-stem logs from default tuning data directory.
        data_dir = self._default_data_dir()
        if data_dir.exists():
            for suffix in ("*.msl", "*.mlg"):
                for candidate in data_dir.glob(suffix):
                    resolved = candidate.resolve()
                    if resolved.stem.lower() != tune_stem:
                        continue
                    if resolved not in seen:
                        matching_logs.append(resolved)
                        seen.add(resolved)

        if not matching_logs:
            QMessageBox.warning(
                self,
                "No Matching Log Found",
                f"Could not find a .msl or .mlg file named '{tune_stem}' to load."
            )
            return

        # Prefer .msl, fall back to first available
        preferred_log = next((p for p in matching_logs if p.suffix.lower() == ".msl"), matching_logs[0])
        self._open_recent_log_file(preferred_log)

    def _most_recent_tune_file(self) -> Path | None:
        candidates: list[Path] = []
        seen: set[Path] = set()

        # 1) Prioritize working folder if set
        if self.working_folder and self.working_folder.exists():
            for path in self.working_folder.glob("*.msq"):
                resolved = path.resolve()
                if resolved.exists() and resolved not in seen:
                    candidates.append(resolved)
                    seen.add(resolved)

        # 2) Then include recent tune files
        for path in self.recent_tune_files:
            resolved = path.resolve()
            if resolved.exists() and resolved.suffix.lower() == ".msq" and resolved not in seen:
                candidates.append(resolved)
                seen.add(resolved)

        # 3) Finally include default data directory
        data_dir = self._default_data_dir()
        if data_dir.exists():
            for path in data_dir.glob("*.msq"):
                resolved = path.resolve()
                if resolved.exists() and resolved not in seen:
                    candidates.append(resolved)
                    seen.add(resolved)

        if not candidates:
            return None

        try:
            return max(candidates, key=lambda p: p.stat().st_mtime)
        except OSError:
            return None

    def _load_recent_files(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        recent_tunes = settings.value("recent_tune_files", [])
        recent_logs = settings.value("recent_log_files", [])

        self.recent_tune_files = [Path(p) for p in recent_tunes if Path(p).exists()]
        self.recent_log_files = [Path(p) for p in recent_logs if Path(p).exists()]

    def _save_recent_files(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("recent_tune_files", [str(p) for p in self.recent_tune_files])
        settings.setValue("recent_log_files", [str(p) for p in self.recent_log_files])

    def _load_working_folder(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        working_folder_str = settings.value("working_folder", "")
        if working_folder_str and Path(working_folder_str).exists():
            self.working_folder = Path(working_folder_str)
        else:
            self.working_folder = None

    def _save_working_folder(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        if self.working_folder:
            settings.setValue("working_folder", str(self.working_folder))
        else:
            settings.remove("working_folder")

    def _on_select_working_folder(self) -> None:
        """Open folder selection dialog and save the working folder."""
        selected_folder = QFileDialog.getExistingDirectory(
            self,
            "Select Working Folder",
            str(self.working_folder) if self.working_folder else str(Path.cwd())
        )
        if selected_folder:
            self.working_folder = Path(selected_folder)
            self._save_working_folder()
            self.statusBar().showMessage(f"Working folder set to: {self.working_folder.name}", 5000)

    def _load_favorites(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        favorites = settings.value("favorite_tables", [])
        self.favorite_tables = set(favorites) if favorites else set()

    def _save_favorites(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("favorite_tables", list(self.favorite_tables))

    def _load_table_filter_preferences(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        # Always start with Favorites mode selected on every launch
        self.only_show_favorited_tables = True
        self.show_1d_tables = False
        self.show_2d_tables = False
        self.show_log_playback = bool(settings.value("show_log_playback", True, type=bool))
        self._normalize_table_filter_mode()

    def _normalize_table_filter_mode(self) -> None:
        """Ensure exactly one table filter mode is active."""
        if self.only_show_favorited_tables:
            self.show_1d_tables = False
            self.show_2d_tables = False
            return
        if self.show_1d_tables:
            self.show_2d_tables = False
            return
        if self.show_2d_tables:
            self.show_1d_tables = False
            return
        self.show_1d_tables = True

    def _save_table_filter_preferences(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("only_show_favorited_tables", self.only_show_favorited_tables)
        settings.setValue("show_1d_tables", self.show_1d_tables)
        settings.setValue("show_2d_tables", self.show_2d_tables)
        settings.setValue("show_log_playback", self.show_log_playback)

    def _default_row_viz_preferences(self) -> dict[str, dict[str, bool]]:
        return {
            "ve": {},
            "afr": {},
            "knock": {},
            "generic": {},
        }

    def _load_row_viz_preferences(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        saved = settings.value("row_viz_preferences")
        defaults = self._default_row_viz_preferences()
        if not isinstance(saved, dict):
            self.row_viz_preferences = defaults
            return

        merged: dict[str, dict[str, bool]] = {}
        for table_type, defaults_for_type in defaults.items():
            merged[table_type] = {}
            saved_for_type = saved.get(table_type, {}) if isinstance(saved.get(table_type, {}), dict) else {}
            for series_id, default_value in defaults_for_type.items():
                merged[table_type][series_id] = bool(saved_for_type.get(series_id, default_value))
        self.row_viz_preferences = merged

    def _save_row_viz_preferences(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("row_viz_preferences", self.row_viz_preferences)

    def _table_log_preferences_path(self) -> Path:
        return Path.home() / ".gyatt_o_tune" / "current_preferences.json"

    @staticmethod
    def _normalize_table_log_channel_preferences(
        raw: dict[str, Any],
    ) -> dict[str, Any]:
        normalized: dict[str, Any] = {}
        for table_name, channels in raw.items():
            if table_name == "__custom_identifiers__":
                if isinstance(channels, dict):
                    normalized[table_name] = {}
                    for identifier_name, cfg in channels.items():
                        if not isinstance(identifier_name, str) or not identifier_name.strip() or not isinstance(cfg, dict):
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
                            target_name: bool(enabled) for target_name, enabled in targets.items()
                        }
                continue
            if table_name == "__favorite_tables__":
                if isinstance(channels, list):
                    normalized[table_name] = [
                        str(name) for name in channels if isinstance(name, str) and str(name).strip()
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

    def _load_table_log_channel_preferences(self) -> None:
        path = self._table_log_preferences_path()
        if not path.exists():
            self.table_log_channel_preferences = {}
            return

        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            tables = payload.get("tables", payload)
            if not isinstance(tables, dict):
                self.table_log_channel_preferences = {}
                return
            self.table_log_channel_preferences = self._normalize_table_log_channel_preferences(tables)
            # Load the show_tunerstudio_names setting from preferences
            if "__show_tunerstudio_names__" in tables:
                self.show_tunerstudio_names = bool(tables.get("__show_tunerstudio_names__", True))
            loaded_favorites = tables.get("__favorite_tables__", [])
            if isinstance(loaded_favorites, list) and loaded_favorites:
                self.favorite_tables = {str(name) for name in loaded_favorites if isinstance(name, str) and str(name).strip()}
                self._save_favorites()
        except Exception:
            self.table_log_channel_preferences = {}

    def _save_table_log_channel_preferences(self) -> None:
        path = self._table_log_preferences_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        tables_to_save = self._normalize_table_log_channel_preferences(self.table_log_channel_preferences)
        # Ensure __show_tunerstudio_names__ is included in the saved preferences
        tables_to_save["__show_tunerstudio_names__"] = self.show_tunerstudio_names
        tables_to_save["__favorite_tables__"] = sorted(self.favorite_tables)
        payload = {
            "version": 1,
            "tables": tables_to_save,
        }
        path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")

    def _point_cloud_detail_channels(self, source_channel_name: str) -> list[str]:
        point_cloud = self.table_log_channel_preferences.get("__point_cloud__", {})
        if not isinstance(point_cloud, dict):
            return []
        targets = point_cloud.get(source_channel_name, {})
        if not isinstance(targets, dict):
            return []
        return [name for name, enabled in targets.items() if bool(enabled)]

    def _custom_identifier_definitions(self) -> dict[str, dict[str, str]]:
        raw = self.table_log_channel_preferences.get("__custom_identifiers__", {})
        if not isinstance(raw, dict):
            return {}
        cleaned: dict[str, dict[str, str]] = {}
        for identifier_name, cfg in raw.items():
            if not isinstance(identifier_name, str) or not isinstance(cfg, dict):
                continue
            expression = str(cfg.get("expression", "")).strip()
            if not expression:
                continue
            cleaned[identifier_name] = {
                "expression": expression,
                "units": str(cfg.get("units", "")).strip(),
            }
        return cleaned

    def _all_identifier_names_for_preferences(self) -> list[str]:
        base = [str(col) for col in self.log_df.columns] if self.log_df is not None else []
        merged = list(base)
        for custom_name in self._custom_identifier_definitions().keys():
            if custom_name not in merged:
                merged.append(custom_name)
        return merged

    @staticmethod
    def _tokenize_identifier_expression(expression: str, candidate_names: list[str]) -> tuple[str, dict[str, str], str | None]:
        candidate_set = set(n for n in candidate_names if isinstance(n, str) and n)
        token_map: dict[str, str] = {}
        name_to_token: dict[str, str] = {}
        idx = 0

        def _replace(m: re.Match) -> str:
            nonlocal idx
            name = m.group(1)
            if name not in name_to_token:
                token = f"__id{idx}__"
                idx += 1
                name_to_token[name] = token
                token_map[token] = name
            return name_to_token[name]

        quote_pattern = re.compile(r'"([^"]*)"')
        transformed = quote_pattern.sub(_replace, expression)
        for name in name_to_token:
            if name not in candidate_set:
                return transformed, token_map, f"Unknown identifier: '{name}'"
        scrubbed = transformed
        for token in token_map:
            scrubbed = scrubbed.replace(token, "")
        scrubbed = re.sub(r"[0-9eE+\-*/().\s]", "", scrubbed)
        if scrubbed:
            return transformed, token_map, f"Unknown identifier content: '{scrubbed}'"
        return transformed, token_map, None

    def _resolve_identifier_series(
        self,
        identifier_name: str,
        custom_defs: dict[str, dict[str, str]],
        cache: dict[str, Any],
        stack: set[str],
    ) -> Any:
        if identifier_name in cache:
            return cache[identifier_name]
        if self.log_df is not None and identifier_name in self.log_df.columns:
            numeric = self._to_numeric_series(self.log_df[identifier_name])
            if numeric is not None:
                cache[identifier_name] = numeric
            return numeric
        if identifier_name in stack:
            return None
        cfg = custom_defs.get(identifier_name)
        if not isinstance(cfg, dict):
            return None

        expression = str(cfg.get("expression", "")).strip()
        if not expression:
            return None

        candidates = []
        if self.log_df is not None:
            candidates.extend([str(col) for col in self.log_df.columns])
        candidates.extend(list(custom_defs.keys()))
        transformed, token_map, error = self._tokenize_identifier_expression(expression, candidates)
        if error is not None:
            return None
        if len(set(token_map.values())) == 0:
            return None

        env: dict[str, Any] = {}
        next_stack = set(stack)
        next_stack.add(identifier_name)
        for token, source_name in token_map.items():
            source_series = self._resolve_identifier_series(source_name, custom_defs, cache, next_stack)
            if source_series is None:
                return None
            env[token] = source_series

        try:
            parsed = ast.parse(transformed, mode="eval")
        except Exception:
            return None

        def _eval(node: ast.AST) -> Any:
            if isinstance(node, ast.Expression):
                return _eval(node.body)
            if isinstance(node, ast.BinOp):
                left = _eval(node.left)
                right = _eval(node.right)
                if isinstance(node.op, ast.Add):
                    return left + right
                if isinstance(node.op, ast.Sub):
                    return left - right
                if isinstance(node.op, ast.Mult):
                    return left * right
                if isinstance(node.op, ast.Div):
                    return left / right
                raise ValueError("Unsupported operator")
            if isinstance(node, ast.UnaryOp):
                value = _eval(node.operand)
                if isinstance(node.op, ast.UAdd):
                    return value
                if isinstance(node.op, ast.USub):
                    return -value
                raise ValueError("Unsupported unary operator")
            if isinstance(node, ast.Name):
                if node.id not in env:
                    raise ValueError("Unknown token")
                return env[node.id]
            if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
                return float(node.value)
            raise ValueError("Unsupported expression")

        try:
            import pandas as pd

            evaluated = _eval(parsed)
            if hasattr(evaluated, "shape"):
                numeric_eval = pd.to_numeric(evaluated, errors="coerce")
            else:
                numeric_eval = pd.Series([float(evaluated)] * len(self.log_df), index=self.log_df.index)
            cache[identifier_name] = numeric_eval
            return numeric_eval
        except Exception:
            return None

    def _on_row_viz_visibility_changed(self, visibility: dict[str, bool]) -> None:
        table_type = self._row_table_type()
        defaults_for_type = self._default_row_viz_preferences().get(table_type, {})
        target = self.row_viz_preferences.setdefault(table_type, {})
        for series_id in defaults_for_type.keys():
            if series_id in visibility:
                target[series_id] = bool(visibility[series_id])
        self._save_row_viz_preferences()

    def _row_table_type(self) -> str:
        if self.current_table is None:
            return "generic"
        table_name = self.current_table.name.lower()
        if "knock" in table_name:
            return "knock"
        if "afr" in table_name:
            return "afr"
        if "ve" in table_name:
            return "ve"
        return "generic"

    def _row_viz_available_series(self, table_type: str, payload: dict[str, Any]) -> list[str]:
        allowed = self.ROW_VIZ_SERIES_BY_TABLE_TYPE.get(table_type, self.ROW_VIZ_SERIES_BY_TABLE_TYPE["generic"])
        existing = {
            str(series.get("series_id"))
            for series in payload.get("point_sets", [])
            if isinstance(series, dict) and series.get("series_id") is not None
        }
        available: list[str] = []
        for series_id in allowed:
            if series_id in existing:
                available.append(series_id)
        return available

    def _is_current_table_1d(self) -> bool:
        return self.current_table is not None and (self.current_table.rows == 1 or self.current_table.cols == 1)

    def _active_row_editor_axis_values(self) -> list[float]:
        if self.current_table is None:
            return []

        if self._is_current_table_1d():
            if self.current_table.rows == 1:
                if self.current_x_axis is not None and len(self.current_x_axis.values) == self.current_table.cols:
                    return [float(v) for v in self.current_x_axis.values]
                return [float(i + 1) for i in range(self.current_table.cols)]

            if self.current_y_axis is not None and len(self.current_y_axis.values) == self.current_table.rows:
                return [float(v) for v in self.current_y_axis.values]
            return [float(i + 1) for i in range(self.current_table.rows)]

        if self.current_x_axis is None:
            return []
        return [float(v) for v in self.current_x_axis.values]

    def _one_d_table_plot_data(self, override_values: list[float] | None = None) -> tuple[list[float], list[float], str]:
        if self.current_table is None:
            return [], [], "Index"

        if self.current_table.rows == 1:
            y_values = [float(v) for v in (override_values if override_values is not None else self.current_table.values[0])]
            if self.current_x_axis is not None and len(self.current_x_axis.values) == len(y_values):
                x_values = [float(v) for v in self.current_x_axis.values]
                x_label = self.current_x_axis.name or "X"
            else:
                x_values = [float(i + 1) for i in range(len(y_values))]
                x_label = "Index"
            return x_values, y_values, x_label

        y_values = [float(v) for v in (override_values if override_values is not None else [row[0] for row in self.current_table.values])]
        if self.current_y_axis is not None and len(self.current_y_axis.values) == len(y_values):
            x_values = [float(v) for v in self.current_y_axis.values]
            x_label = self.current_y_axis.name or "Y"
        else:
            x_values = [float(i + 1) for i in range(len(y_values))]
            x_label = "Index"
        return x_values, y_values, x_label

    def _build_1d_table_visualization_payload(self, override_values: list[float] | None = None) -> dict[str, Any]:
        x_values, y_values, x_label = self._one_d_table_plot_data(override_values)
        table_name = self.current_table.name if self.current_table is not None else "Table"
        units = self.current_table.units if self.current_table is not None else None

        point_sets: list[dict[str, Any]] = [
            {
                "series_id": "table",
                "name": "Full 1D Table",
                "rpm": x_values,
                "ve": y_values,
                "map": [0.0] * len(x_values),
                "ve_raw": y_values,
                "ve_scaled": y_values,
                "afr": [None] * len(x_values),
                "afr_target": [None] * len(x_values),
                "afr_error": [None] * len(x_values),
            }
        ]

        title = f"{table_name} - 1D Table Data"
        if y_values:
            stats = f"Full 1D table values\nPoints: {len(y_values)}\nRange: {min(y_values):.3f} - {max(y_values):.3f}"
        else:
            stats = "No table values available"

        y_label = units if units else "Value"

        if self.log_df is not None and not self.log_df.empty:
            rpm_channel = None
            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if rpm_channel is None and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)

            if rpm_channel is not None:
                rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
                if rpm_series is not None:
                    map_mask = rpm_series == rpm_series
                    filtered_rpm_values = [float(v) for v in rpm_series[map_mask]]
                    filtered_map_values = [0.0] * len(filtered_rpm_values)
                    self._apply_table_log_channel_preferences_to_payload(
                        table_name=table_name,
                        point_sets=point_sets,
                        table_rpm_values=[float(v) for v in x_values],
                        map_mask=map_mask,
                        filtered_rpm_values=filtered_rpm_values,
                        filtered_map_values=filtered_map_values,
                    )
                    self._apply_afr_prediction_to_payload(
                        table_name=table_name,
                        point_sets=point_sets,
                        table_rpm_values=[float(v) for v in x_values],
                        map_mask=map_mask,
                        filtered_rpm_values=filtered_rpm_values,
                    )
                    stats += f"\nLog points: {len(filtered_rpm_values)}"

        return {
            "title": title,
            "stats": stats,
            "point_sets": point_sets,
            "y_label": y_label,
            "x_label": x_label,
        }

    @staticmethod
    def _get_tunerstudio_name(table_name: str) -> str:
        """Convert internal table name to TunerStudio display name."""
        name_map = {
            # VE Tables
            "veTable1": "VE Table 1",
            "veTable2": "VE Table 2",
            "veTable3": "VE Table 3",
            "veTable4": "VE Table 4",
            "idleve_table1": "Idle VE Table 1",
            "idleve_table2": "Idle VE Table 2",

            # Ignition Tables
            "advanceTable1": "Ignition Advance Table 1",
            "advanceTable2": "Ignition Advance Table 2",
            "advanceTable3": "Ignition Advance Table 3",
            "advanceTable4": "Ignition Advance Table 4",

            # AFR Tables
            "afrTable1": "AFR Table 1",
            "afrTable2": "AFR Table 2",

            # Knock Tables
            "knock_thresholds": "Knock Threshold Table",

            # Boost Control
            "boost_ctl_load_targets": "Boost Control Load Targets",
            "boost_ctl_load_targets2": "Boost Control Load Targets 2",
            "boost_ctl_pwm_targets": "Boost Control PWM Targets",
            "boost_ctl_pwm_targets2": "Boost Control PWM Targets 2",
            "boost_dome_targets1": "Boost Dome Targets 1",
            "boost_dome_targets2": "Boost Dome Targets 2",

            # Injection
            "inj_deadtime_table1": "Injector Deadtime Table 1",
            "inj_deadtime_table2": "Injector Deadtime Table 2",
            "inj_deadtime_table3": "Injector Deadtime Table 3",
            "inj_deadtime_table4": "Injector Deadtime Table 4",
            "inj_timing": "Injection Timing",
            "inj_timing_sec": "Injection Timing Secondary",

            # Spark Trim Tables
            "spk_trima": "Spark Trim A",
            "spk_trimb": "Spark Trim B",
            "spk_trimc": "Spark Trim C",
            "spk_trimd": "Spark Trim D",
            "spk_trime": "Spark Trim E",
            "spk_trimf": "Spark Trim F",
            "spk_trimg": "Spark Trim G",
            "spk_trimh": "Spark Trim H",
            "spk_trimi": "Spark Trim I",
            "spk_trimj": "Spark Trim J",
            "spk_trimk": "Spark Trim K",
            "spk_triml": "Spark Trim L",
            "spk_trimm": "Spark Trim M",
            "spk_trimn": "Spark Trim N",
            "spk_trimo": "Spark Trim O",
            "spk_trimp": "Spark Trim P",

            # Fuel Trim Tables
            "inj_trima": "Fuel Trim A",
            "inj_trimb": "Fuel Trim B",
            "inj_trimc": "Fuel Trim C",
            "inj_trimd": "Fuel Trim D",
            "inj_trime": "Fuel Trim E",
            "inj_trimf": "Fuel Trim F",
            "inj_trimg": "Fuel Trim G",
            "inj_trimh": "Fuel Trim H",
            "inj_trimi": "Fuel Trim I",
            "inj_trimj": "Fuel Trim J",
            "inj_trimk": "Fuel Trim K",
            "inj_triml": "Fuel Trim L",
            "inj_trimm": "Fuel Trim M",
            "inj_trimn": "Fuel Trim N",
            "inj_trimo": "Fuel Trim O",
            "inj_trimp": "Fuel Trim P",

            # Other common tables
            "RotarySplitTable": "Rotary Split Table",
            "alphaMAPtable": "Alpha-N MAP Table",
            "dwell_table_values": "Dwell Table",
            "ego_auth_table": "EGO Authority Table",
            "ego_auth_table2": "EGO Authority Table 2",
            "ego_delay_table": "EGO Delay Table",
            "etc_targ_pos": "ETC Target Position",
            "fpd_duty": "FPD Duty",
            "ltt_table1": "Launch Timing Table 1",
            "map_predict_lookup_table": "MAP Predict Lookup Table",
            "map_predict_lookup_table2": "MAP Predict Lookup Table 2",
            "narrowband_tgts": "Narrowband Targets",
            "vss_diff_table": "VSS Diff Table",
            "vvt_timing1": "VVT Timing 1",
            "vvt_timing2": "VVT Timing 2",
            "waterinj_duty": "Water Injection Duty",
        }

        return name_map.get(table_name, table_name)  # Return original name if not mapped

    def _add_recent_tune_file(self, file_path: Path) -> None:
        if file_path in self.recent_tune_files:
            self.recent_tune_files.remove(file_path)
        self.recent_tune_files.insert(0, file_path)
        self.recent_tune_files = self.recent_tune_files[:self.MAX_RECENT_FILES]
        self._save_recent_files()
        self._update_recent_tunes_menu()

    def _add_recent_log_file(self, file_path: Path) -> None:
        if file_path in self.recent_log_files:
            self.recent_log_files.remove(file_path)
        self.recent_log_files.insert(0, file_path)
        self.recent_log_files = self.recent_log_files[:self.MAX_RECENT_FILES]
        self._save_recent_files()
        self._update_recent_logs_menu()

    def _update_recent_tunes_menu(self) -> None:
        self.recent_tunes_menu.clear()
        if not self.recent_tune_files:
            self.recent_tunes_menu.setEnabled(False)
            return

        self.recent_tunes_menu.setEnabled(True)
        for i, file_path in enumerate(self.recent_tune_files):
            action = self.recent_tunes_menu.addAction(f"&{i+1} {file_path.name}")
            action.setToolTip(str(file_path))
            action.triggered.connect(lambda checked, path=file_path: self._open_recent_tune_file(path))

    def _update_recent_logs_menu(self) -> None:
        self.recent_logs_menu.clear()
        if not self.recent_log_files:
            self.recent_logs_menu.setEnabled(False)
            return

        self.recent_logs_menu.setEnabled(True)
        for i, file_path in enumerate(self.recent_log_files):
            action = self.recent_logs_menu.addAction(f"&{i+1} {file_path.name}")
            action.setToolTip(str(file_path))
            action.triggered.connect(lambda checked, path=file_path: self._open_recent_log_file(path))

    def _open_recent_tune_file(self, file_path: Path) -> bool:
        if not file_path.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {file_path} no longer exists.")
            if file_path in self.recent_tune_files:
                self.recent_tune_files.remove(file_path)
            self._save_recent_files()
            self._update_recent_tunes_menu()
            return False

        try:
            self.tune_data = self.tune_loader.load(file_path)
            self.selected_rows_per_table = {}  # Clear saved row selections for new tune
            self.pending_edits_per_table = {}  # Clear saved pending edits for new tune
            self.loaded_table_snapshots = self._snapshot_table_values(self.tune_data)
            self._clear_global_history()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Tune Load Error", f"Could not load tune file:\n{exc}")
            return False

        self.loaded_tune_path = file_path
        self.save_target_path = None
        self._set_tune_save_actions_enabled(True)
        self.table_list.clear()
        if self.tune_data.tables:
            self._update_table_display()
            self._select_default_table_on_load()
        else:
            self.table_list.addItem("No tables found")
            self.table_grid.clear()
            self.table_grid.setRowCount(0)
            self.table_grid.setColumnCount(0)
            self.table_meta.setText("No tables found in this tune file")
            self.table_meta.setToolTip("")
            self.axis_meta.setText("Axes: X (index), Y (index)")

        self.statusBar().showMessage(f"Loaded tune: {file_path.name}")
        self._refresh_workspace_text()
        self._add_recent_tune_file(file_path)
        self._maybe_prompt_load_matching_log(file_path)
        return True

    def _open_recent_log_file(self, file_path: Path) -> None:
        if not file_path.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {file_path} no longer exists.")
            if file_path in self.recent_log_files:
                self.recent_log_files.remove(file_path)
            self._save_recent_files()
            self._update_recent_logs_menu()
            return

        self.statusBar().showMessage(f"Loading log: {file_path.name}...")
        self.setCursor(Qt.CursorShape.WaitCursor)
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            self.statusBar().showMessage(f"Failed to load log: {file_path.name}", 5000)
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return
        finally:
            self.setCursor(Qt.CursorShape.ArrowCursor)

        self.log_file = file_path
        self.statusBar().showMessage(
            f"Loaded log: {file_path.name} ({len(log_df):,} rows, parser={parse_result.parser_used}, enc={parse_result.encoding})"
        )
        self.log_df = log_df
        self._add_recent_log_file(file_path)

    def _is_1d_table(self, table_name: str) -> bool:
        """Check if a table is 1D (either rows==1 or cols==1)."""
        if not self.tune_data or table_name not in self.tune_data.tables:
            return False
        table = self.tune_data.tables[table_name]
        return table.rows == 1 or table.cols == 1

    def _update_table_display(self) -> None:
        """Update the table list display based on current settings."""
        if not self.tune_data:
            return

        current_item = self.table_list.currentItem()
        current_table_name = current_item.data(Qt.ItemDataRole.UserRole) if current_item else None

        self.table_list.clear()
        all_items = []
        for table_name in self.tune_data.tables.keys():
            # Apply filters
            if self.only_show_favorited_tables:
                if table_name not in self.favorite_tables:
                    continue
            else:
                is_1d = self._is_1d_table(table_name)
                if is_1d and not self.show_1d_tables:
                    continue
                if (not is_1d) and not self.show_2d_tables:
                    continue

            display_name = self._get_tunerstudio_name(table_name) if self.show_tunerstudio_names else table_name
            if table_name in self.favorite_tables and not self.only_show_favorited_tables:
                display_name = f"★ {display_name}"
            all_items.append((table_name, display_name))

        all_items.sort(key=lambda x: x[1].lstrip("★ ").lower())

        if all_items:
            selected_idx = 0
            for table_name, display_name in all_items:
                item = QListWidgetItem(display_name)
                item.setData(Qt.ItemDataRole.UserRole, table_name)
                self.table_list.addItem(item)
            if current_table_name is not None:
                for idx, (table_name, _) in enumerate(all_items):
                    if table_name == current_table_name:
                        selected_idx = idx
                        break
                else:
                    for idx, (table_name, _) in enumerate(all_items):
                        if table_name.lower() == "vetable1":
                            selected_idx = idx
                            break
            else:
                for idx, (table_name, _) in enumerate(all_items):
                    if table_name.lower() == "vetable1":
                        selected_idx = idx
                        break

            self.table_list.setCurrentRow(selected_idx)
        else:
            if self.only_show_favorited_tables:
                self.table_list.addItem("No favorite tables")
            elif self.show_1d_tables:
                self.table_list.addItem("No 1D tables found")
            elif self.show_2d_tables:
                self.table_list.addItem("No 2D tables found")
            else:
                self.table_list.addItem("No tables found")

    def _select_default_table_on_load(self) -> None:
        if self.table_list.count() <= 0:
            return

        target_idx = 0
        for idx in range(self.table_list.count()):
            item = self.table_list.item(idx)
            if not item:
                continue
            table_name = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(table_name, str) and table_name.lower() == "vetable1":
                target_idx = idx
                break

        self.table_list.setCurrentRow(target_idx)

    def _sync_table_filter_buttons(self) -> None:
        """Keep table filter button states in sync."""
        self._normalize_table_filter_mode()
        self.favorite_tables_button.blockSignals(True)
        self.favorite_tables_button.setChecked(self.only_show_favorited_tables)
        self.favorite_tables_button.blockSignals(False)

        self.show_1d_tables_button.blockSignals(True)
        self.show_1d_tables_button.setChecked(self.show_1d_tables)
        self.show_1d_tables_button.blockSignals(False)

        self.show_2d_tables_button.blockSignals(True)
        self.show_2d_tables_button.setChecked(self.show_2d_tables)
        self.show_2d_tables_button.blockSignals(False)

    def _on_only_show_favorited_toggled(self, checked: bool) -> None:
        """Handle 'Favorite Tables Only' toggle."""
        if not checked:
            return
        self.only_show_favorited_tables = True
        self.show_1d_tables = False
        self.show_2d_tables = False
        self._save_table_filter_preferences()
        self._update_table_display()

    def _on_show_2d_tables_toggled(self, checked: bool) -> None:
        """Handle 'Show 2D Tables' toggle."""
        if not checked:
            return
        self.only_show_favorited_tables = False
        self.show_1d_tables = False
        self.show_2d_tables = True
        self._save_table_filter_preferences()
        self._update_table_display()

    def _on_show_1d_tables_toggled(self, checked: bool) -> None:
        """Handle 'Show 1D Tables' toggle."""
        if not checked:
            return
        self.only_show_favorited_tables = False
        self.show_1d_tables = True
        self.show_2d_tables = False
        self._save_table_filter_preferences()
        self._update_table_display()

    def _on_show_log_playback_toggled(self, enabled: bool | None = None) -> None:
        """Handle 'Log Playback' toggle."""
        if enabled is None:
            enabled = not self.show_log_playback
        self.show_log_playback = bool(enabled)
        if hasattr(self, "table_row_panel") and self.table_row_panel is not None:
            self.table_row_panel.set_playback_controls_enabled(self.show_log_playback)
        self._save_table_filter_preferences()

    def _open_preferences(self, initial_table_name: str | None = None, initial_scatter_identifier: str | None = None) -> None:
        if self.tune_data is None or not self.tune_data.tables:
            QMessageBox.information(self, "Preferences", "Load a tune file first to configure per-table preferences.")
            return

        if self.log_df is None or self.log_df.empty:
            QMessageBox.information(self, "Preferences", "Load a log file first to configure logged channel preferences.")
            return

        # Always start the dialog from the canonical current preferences file.
        self._load_table_log_channel_preferences()

        table_names = sorted(self.tune_data.tables.keys())
        table_display_names = {
            table_name: (self._get_tunerstudio_name(table_name) if self.show_tunerstudio_names else table_name)
            for table_name in table_names
        }
        table_dimensions = {
            table_name: (table.rows == 1 or table.cols == 1)
            for table_name, table in self.tune_data.tables.items()
        }
        log_channel_names = self._all_identifier_names_for_preferences()
        if initial_table_name is None:
            initial_table_name = self.current_table.name if self.current_table is not None else "veTable1"
        # Add current show_tunerstudio_names setting to preferences dict for dialog
        prefs_with_tunerstudio = dict(self.table_log_channel_preferences)
        prefs_with_tunerstudio["__show_tunerstudio_names__"] = self.show_tunerstudio_names
        prefs_with_tunerstudio["__favorite_tables__"] = sorted(self.favorite_tables)
        dialog = TableLogChannelPreferencesDialog(
            favorite_tables=self.favorite_tables,
            table_names=table_names,
            table_dimensions=table_dimensions,
            table_display_names=table_display_names,
            log_channel_names=log_channel_names,
            current_preferences=prefs_with_tunerstudio,
            current_preferences_path=self._table_log_preferences_path(),
            initial_table_name=initial_table_name,
            parent=self,
        )
        # Sync the dialog's Tables tab filter buttons to the current main-window state.
        if self.only_show_favorited_tables:
            dialog.favorite_btn.setChecked(True)
        elif self.show_2d_tables:
            dialog.show_2d_btn.setChecked(True)
        else:
            dialog.show_1d_btn.setChecked(True)

        if initial_scatter_identifier is not None:
            dialog.select_scatterplot_tab(initial_scatter_identifier)
        else:
            dialog.select_tune_tables_tab(initial_table_name)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.statusBar().showMessage("Cancelled row data preferences changes", 3000)
            # Reload preferences from disk to discard any in-memory changes
            self._load_table_log_channel_preferences()
            return

        # Apply preferences whenever the dialog is closed, except explicit Cancel.
        self.table_log_channel_preferences = dialog.preferences()
        self.show_tunerstudio_names = dialog._show_tunerstudio_names
        loaded_favorites = self.table_log_channel_preferences.get("__favorite_tables__", [])
        if isinstance(loaded_favorites, list):
            self.favorite_tables = {str(name) for name in loaded_favorites if isinstance(name, str) and str(name).strip()}
            self._save_favorites()
        try:
            self._save_table_log_channel_preferences()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Preferences", f"Preferences updated in memory, but could not save to disk:\n{exc}")

        self.statusBar().showMessage("Saved row data preferences", 3000)
        self._update_table_display()
        self._update_table_grid_row_visualization()

    def _on_table_list_item_double_clicked(self, item: QListWidgetItem) -> None:
        """Add to favorites when a table item is double-clicked (does not remove if already favorited)."""
        table_name = item.data(Qt.ItemDataRole.UserRole)
        if table_name and table_name not in self.favorite_tables:
            self._toggle_favorite(table_name)

    def _show_table_context_menu(self, position) -> None:
        """Show context menu for table list items."""
        list_widget = self.table_list
        item = list_widget.itemAt(position)
        if not item:
            return

        table_name = item.data(Qt.ItemDataRole.UserRole)
        if not table_name:
            return

        menu = QMenu(self)

        if table_name in self.favorite_tables:
            unfavorite_action = menu.addAction("Remove from Favorites")
            unfavorite_action.triggered.connect(lambda: self._toggle_favorite(table_name))
        else:
            favorite_action = menu.addAction("Add to Favorites")
            favorite_action.triggered.connect(lambda: self._toggle_favorite(table_name))

        menu.addSeparator()
        prefs_action = menu.addAction("Table Preferences...")
        prefs_action.triggered.connect(lambda: self._open_preferences(initial_table_name=table_name))

        menu.exec(list_widget.mapToGlobal(position))

    def _toggle_favorite(self, table_name: str) -> None:
        """Toggle favorite status for a table."""
        if table_name in self.favorite_tables:
            self.favorite_tables.remove(table_name)
        else:
            self.favorite_tables.add(table_name)
        self._save_favorites()
        self.table_log_channel_preferences["__favorite_tables__"] = sorted(self.favorite_tables)
        try:
            self._save_table_log_channel_preferences()
        except Exception:
            pass
        self._update_table_display()

    def _create_layout(self) -> None:
        self.setDockNestingEnabled(True)
        self.setDockOptions(
            QMainWindow.DockOption.AllowNestedDocks
            | QMainWindow.DockOption.AllowTabbedDocks
            | QMainWindow.DockOption.GroupedDragging
            | QMainWindow.DockOption.AnimatedDocks
        )
        self.setTabPosition(Qt.DockWidgetArea.AllDockWidgetAreas, QTabWidget.TabPosition.North)

        # Left panel - Tables
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        table_filters_row = QHBoxLayout()
        table_filters_row.setContentsMargins(4, 4, 4, 4)
        table_filters_row.setSpacing(8)

        filter_btn_w = 48

        self.favorite_tables_button = QPushButton()
        self.favorite_tables_button.setCheckable(True)
        self.favorite_tables_button.setFixedWidth(filter_btn_w)
        self.favorite_tables_button.setToolTip("Favorite Tables Only")
        star_pixmap = getattr(QStyle.StandardPixmap, "SP_StarButton", None)
        if star_pixmap is not None:
            self.favorite_tables_button.setIcon(self.style().standardIcon(star_pixmap))
        else:
            self.favorite_tables_button.setText("★")
        self.favorite_tables_button.toggled.connect(self._on_only_show_favorited_toggled)

        self.show_1d_tables_button = QPushButton("1D")
        self.show_1d_tables_button.setCheckable(True)
        self.show_1d_tables_button.setFixedWidth(filter_btn_w)
        self.show_1d_tables_button.setToolTip("Show 1D tables")
        self.show_1d_tables_button.toggled.connect(self._on_show_1d_tables_toggled)

        self.show_2d_tables_button = QPushButton("2D")
        self.show_2d_tables_button.setCheckable(True)
        self.show_2d_tables_button.setFixedWidth(filter_btn_w)
        self.show_2d_tables_button.setToolTip("Show 2D tables")
        self.show_2d_tables_button.toggled.connect(self._on_show_2d_tables_toggled)

        self._main_table_filter_group = QButtonGroup(self)
        self._main_table_filter_group.setExclusive(True)
        self._main_table_filter_group.addButton(self.favorite_tables_button)
        self._main_table_filter_group.addButton(self.show_1d_tables_button)
        self._main_table_filter_group.addButton(self.show_2d_tables_button)

        table_filters_row.addStretch(1)
        table_filters_row.addWidget(self.favorite_tables_button)
        table_filters_row.addStretch(1)
        table_filters_row.addWidget(self.show_1d_tables_button)
        table_filters_row.addStretch(1)
        table_filters_row.addWidget(self.show_2d_tables_button)
        table_filters_row.addStretch(1)

        left_layout.addLayout(table_filters_row)

        self.table_list = TuneTablesListWidget()
        self.table_list.on_tune_file_dropped = self._on_tune_file_dropped
        self.table_list.setStyleSheet(
            """
            QListWidget {
                border: none;
                background: transparent;
                padding: 4px;
                outline: 0;
            }
            QListWidget::item {
                border-radius: 10px;
                border: 1px solid palette(mid);
                background-color: palette(base);
                color: palette(text);
                margin: 2px 6px;
                padding: 5px 10px;
            }
            QListWidget::item:hover {
                background-color: palette(alternate-base);
                border: 1px solid palette(midlight);
            }
            QListWidget::item:selected {
                background-color: #2d5ecf;
                color: #ffffff;
                border: 1px solid #4b82ff;
            }
            """
        )
        self._add_open_tune_file_placeholder()
        self.table_list.currentItemChanged.connect(self._on_table_selected)
        self.table_list.itemClicked.connect(self._on_table_list_item_clicked)
        self.table_list.itemDoubleClicked.connect(self._on_table_list_item_double_clicked)
        self.table_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_list.customContextMenuRequested.connect(self._show_table_context_menu)
        left_layout.addWidget(self.table_list)
        self._sync_table_filter_buttons()

        # Selected table panel
        selected_table_panel = QWidget()
        selected_table_layout = QVBoxLayout(selected_table_panel)
        selected_table_layout.setContentsMargins(0, 0, 0, 0)

        # Header area: labels on the left, action button on the right spanning both rows.
        header_layout = QGridLayout()

        self.table_meta = QLabel("No table selected")
        header_layout.addWidget(self.table_meta, 0, 0)

        # Second row with axes labels
        self.axis_meta = QLabel("Axes: X (index), Y (index)")
        header_layout.addWidget(self.axis_meta, 1, 0)

        self.plot_all_data_button = QPushButton("Plot All Data")
        self.plot_all_data_button.setToolTip("Show all unfiltered scatterplot data for the selected table")
        self.plot_all_data_button.clicked.connect(self._on_plot_all_data_requested)
        self.plot_all_data_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 6px;")
        self.plot_all_data_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        header_layout.addWidget(self.plot_all_data_button, 0, 1, 2, 1)
        header_layout.setColumnStretch(0, 1)
        selected_table_layout.addLayout(header_layout)

        self.table_grid = CopyPasteTableWidget()
        self.table_grid.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table_grid.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_grid.itemSelectionChanged.connect(self._on_table_grid_row_selection_changed)
        self.table_grid.cellClicked.connect(self._on_table_grid_cell_clicked)
        self.table_grid.currentCellChanged.connect(self._on_table_grid_current_cell_changed)
        self.table_grid.on_adjust_value = self._on_selected_table_cell_adjust_requested
        self.table_grid.on_undo = self._on_selected_table_cell_undo_requested
        self.table_grid.on_redo = self._on_selected_table_cell_redo_requested
        self.table_grid.horizontalHeader().setVisible(False)
        self.table_grid.verticalHeader().setVisible(True)  # leftmost axis labels

        table_grid_container = QWidget()
        self.table_grid_container = table_grid_container
        table_grid_layout = QVBoxLayout(table_grid_container)
        table_grid_layout.setContentsMargins(0, 0, 0, 0)
        table_grid_layout.addWidget(self.table_grid)
        self.modified_table_meta = QLabel("")
        self.modified_table_meta.setVisible(False)
        table_grid_layout.addWidget(self.modified_table_meta)

        self.modified_table_grid = CopyPasteTableWidget()
        self.modified_table_grid.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.modified_table_grid.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.modified_table_grid.horizontalHeader().setVisible(False)
        self.modified_table_grid.verticalHeader().setVisible(True)
        self.modified_table_grid.setVisible(False)
        table_grid_layout.addWidget(self.modified_table_grid)
        selected_table_layout.addWidget(table_grid_container, 1)

        self.table_row_panel = RowVisualizationPanel()
        self.table_row_panel.on_point_selected = self._on_table_row_plot_point_selected
        self.table_row_panel.on_scatter_preferences_requested = self._on_scatter_preferences_requested
        self.table_row_panel.on_table_preferences_requested = self._on_table_preferences_requested
        self.table_row_panel.on_point_adjust_requested = self._on_table_row_plot_point_adjust_requested
        self.table_row_panel.on_visibility_changed = self._on_row_viz_visibility_changed
        self.table_row_panel.on_log_playback_toggled = self._on_show_log_playback_toggled
        self.table_row_panel.get_log_playback_state = lambda: self.show_log_playback
        self.table_row_panel.on_plot_all_data_requested = self._on_plot_all_data_requested
        self.table_row_panel.set_playback_controls_enabled(self.show_log_playback)

        self.table_row_editor_group = QGroupBox("")
        row_editor_layout = QVBoxLayout(self.table_row_editor_group)

        row_editor_controls = QHBoxLayout()
        self.table_row_label = QLabel("Selected row: none")
        row_editor_controls.addWidget(self.table_row_label)
        self.revert_row_changes_button = QPushButton("Revert Row")
        self.revert_row_changes_button.clicked.connect(self._revert_pending_row_values)
        row_editor_controls.addWidget(self.revert_row_changes_button)
        row_editor_controls.addStretch(1)
        row_editor_layout.addLayout(row_editor_controls)

        self.table_row_status = QLabel("Select a VE table and click a row to view and edit it.")
        self.table_row_status.setWordWrap(True)
        row_editor_layout.addWidget(self.table_row_status)

        self.table_row_editor = RowEditorTableWidget()
        self.table_row_editor.on_adjust_value = self._adjust_pending_row_value
        self.table_row_editor.on_undo = self._undo_pending_row_edit
        self.table_row_editor.on_redo = self._redo_pending_row_edit
        self.table_row_editor.on_focus_changed = self._on_row_editor_focus_changed
        self.table_row_editor.currentCellChanged.connect(self._on_row_editor_current_cell_changed)
        self.table_row_editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_row_editor.customContextMenuRequested.connect(self._on_row_editor_context_menu)
        row_editor_layout.addWidget(self.table_row_editor)

        self.table_row_editor_help = QLabel(
            "Click a row in the table grid. Use left/right to move, up/down or mouse wheel to change by 0.1. Ctrl+Z reverts the selected cell; Revert Row resets to the loaded row values."
        )
        row_editor_layout.addWidget(self.table_row_editor_help)

        row_editor_container = QWidget()
        row_editor_container_layout = QVBoxLayout(row_editor_container)
        row_editor_container_layout.setContentsMargins(0, 0, 0, 0)
        row_editor_container_layout.addWidget(self.table_row_editor_group)
        self.table_row_editor.setEnabled(False)
        self._set_row_editor_button_visibility(
            show_generate=False,
            show_write=False,
            show_revert=False,
        )

        # Wrap row visualization panel to avoid it inheriting dock margins directly.
        row_viz_container = QWidget()
        row_viz_layout = QVBoxLayout(row_viz_container)
        row_viz_layout.setContentsMargins(0, 0, 0, 0)
        row_viz_layout.addWidget(self.table_row_panel)

        self.tune_tables_dock = self._create_section_dock(
            "Tables",
            "tuneTablesDock",
            self._create_outlined_panel(left),
        )
        self.selected_table_dock = self._create_section_dock(
            "Selected Table",
            "selectedTableDock",
            self._create_outlined_panel(selected_table_panel),
        )
        self.row_viz_dock = self._create_section_dock("Scatterplot", "rowVizDock", row_viz_container)
        self.row_editor_dock = self._create_section_dock(
            "Row Editor",
            "rowEditorDock",
            self._create_outlined_panel(row_editor_container),
        )

        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.tune_tables_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.selected_table_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_viz_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_editor_dock)

        self._apply_standard_window_layout()

        windows_tables_action = self.tune_tables_dock.toggleViewAction()
        windows_tables_action.setText("Tables")
        self.windows_menu.addAction(windows_tables_action)

        windows_selected_tables_action = self.selected_table_dock.toggleViewAction()
        windows_selected_tables_action.setText("Selected Tables")
        self.windows_menu.addAction(windows_selected_tables_action)

        windows_scatterplot_action = self.row_viz_dock.toggleViewAction()
        windows_scatterplot_action.setText("Scatterplot")
        self.windows_menu.addAction(windows_scatterplot_action)

        windows_row_editor_action = self.row_editor_dock.toggleViewAction()
        windows_row_editor_action.setText("Row Editor")
        self.windows_menu.addAction(windows_row_editor_action)

    def _find_closest_index(self, values: list[float], target: float) -> int | None:
        """Find the index of the closest value in a list."""
        if not values:
            return None

        closest_idx = 0
        min_diff = abs(values[0] - target)

        for i, val in enumerate(values[1:], 1):
            diff = abs(val - target)
            if diff < min_diff:
                min_diff = diff
                closest_idx = i

        return closest_idx

    def _build_row_visualization_payload(
        self,
        map_value: float,
        rpm_values: list[float],
        table_row_values: list[float],
    ) -> dict[str, Any]:
        rpm_point_values = [float(v) for v in rpm_values]
        table_point_values = [float(v) for v in table_row_values]
        table_name = self.current_table.name if self.current_table is not None else "Table"
        y_label = self.current_table.units if self.current_table is not None and self.current_table.units else "Value"
        point_sets: list[dict[str, Any]] = [
            {
                "series_id": "table",
                "name": "Selected Row Data",
                "rpm": rpm_point_values,
                "ve": table_point_values,
                "map": [float(map_value)] * len(rpm_point_values),
                "ve_raw": table_point_values,
                "ve_scaled": table_point_values,
                "afr": [None] * len(rpm_point_values),
                "afr_predicted": [None] * len(rpm_point_values),
                "afr_target": [None] * len(rpm_point_values),
                "afr_error": [None] * len(rpm_point_values),
            }
        ]

        title = f"MAP {map_value} kPa - Selected Row"
        stats = f"Selected row values for MAP {map_value} kPa\n"
        stats += f"RPM range: {min(rpm_values):.0f} - {max(rpm_values):.0f}\n"
        stats += f"Value range: {min(table_row_values):.3f} - {max(table_row_values):.3f}"

        if self.log_df is None or self.log_df.empty:
            return {
                "title": title,
                "stats": stats + "\n\nNo log data loaded.",
                "point_sets": point_sets,
                "y_label": y_label,
            }

        try:
            rpm_channel = None
            map_channel = None

            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if not map_channel and ('map' in col_lower):
                    map_channel = str(col)

            if not rpm_channel or not map_channel:
                return {
                    "title": f"MAP {map_value} kPa - Required channels not found",
                    "stats": stats + "\n\nRequired channels not found in log data.\nNeed: RPM and MAP",
                    "point_sets": point_sets,
                    "y_label": y_label,
                }

            rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
            map_series = self._to_numeric_series(self.log_df[map_channel])

            if rpm_series is None or map_series is None:
                return {
                    "title": f"MAP {map_value} kPa - Could not convert data",
                    "stats": stats + "\n\nCould not convert log data to numeric values.",
                    "point_sets": point_sets,
                    "y_label": y_label,
                }

            map_tolerance = 10.0
            if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
                sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
                map_tolerance = max(10.0, min_spacing * 0.75)
            map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
            filtered_rpm = rpm_series[map_mask]
            filtered_map = map_series[map_mask]

            if len(filtered_rpm) == 0 and map_tolerance < 20.0:
                map_tolerance = 20.0
                map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                filtered_rpm = rpm_series[map_mask]
                filtered_map = map_series[map_mask]

            if len(filtered_rpm) == 0:
                return {
                    "title": f"MAP {map_value} kPa - No data near MAP value",
                    "stats": stats + f"\n\nNo log data found near MAP {map_value} kPa (±{map_tolerance} kPa).",
                    "point_sets": point_sets,
                    "y_label": y_label,
                }

            filtered_rpm_values = [float(v) for v in filtered_rpm]
            filtered_map_values = [float(v) for v in filtered_map]
            self._apply_table_log_channel_preferences_to_payload(
                table_name=table_name,
                point_sets=point_sets,
                table_rpm_values=rpm_point_values,
                map_mask=map_mask,
                filtered_rpm_values=filtered_rpm_values,
                filtered_map_values=filtered_map_values,
                map_value=map_value,
                map_series=map_series,
                rpm_series=rpm_series,
            )
            self._apply_afr_prediction_to_payload(
                table_name=table_name,
                point_sets=point_sets,
                table_rpm_values=rpm_point_values,
                map_mask=map_mask,
                filtered_rpm_values=filtered_rpm_values,
            )
            stats += f"\n\nLogged Data ({len(filtered_rpm)} points):\n"
            stats += f"RPM range: {filtered_rpm.min():.0f} - {filtered_rpm.max():.0f}"

            return {
                "title": f"MAP {map_value} kPa - {len(filtered_rpm)} points",
                "stats": stats,
                "point_sets": point_sets,
                "y_label": y_label,
            }

        except Exception as exc:
            return {
                "title": f"MAP {map_value} kPa - Error: {exc}",
                "stats": stats + f"\n\nError plotting logged data: {exc}",
                "point_sets": point_sets,
                "y_label": y_label,
            }

    def _build_all_data_visualization_payload(self) -> dict[str, Any] | None:
        """Build a row-viz payload containing ALL log scatter data with no MAP band filter.

        The "table" (selected-row) series is intentionally excluded so only the unfiltered
        log point clouds are shown.
        """
        if self.current_table is None or self.current_x_axis is None:
            return None

        table_name = self._current_table_name()
        if table_name is None:
            return None

        y_label = self.current_table.units if self.current_table.units else "Value"
        title = f"All Log Data — {table_name}"

        if self.log_df is None or self.log_df.empty:
            return {
                "title": title,
                "stats": "No log data loaded.",
                "available_series": [],
                "series_visibility": {},
                "point_sets": [],
                "y_label": y_label,
            }

        # Locate the RPM channel.
        rpm_channel: str | None = None
        for col in self.log_df.columns:
            col_lower = str(col).lower()
            if not rpm_channel and ("rpm" in col_lower or "engine speed" in col_lower):
                rpm_channel = str(col)
                break

        if rpm_channel is None:
            return {
                "title": title,
                "stats": "RPM channel not found in log data.",
                "available_series": [],
                "series_visibility": {},
                "point_sets": [],
                "y_label": y_label,
            }

        rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
        if rpm_series is None:
            return {
                "title": title,
                "stats": "Could not convert RPM data to numeric values.",
                "available_series": [],
                "series_visibility": {},
                "point_sets": [],
                "y_label": y_label,
            }

        table_prefs = self.table_log_channel_preferences.get(table_name, {})
        if not isinstance(table_prefs, dict) or not table_prefs:
            return {
                "title": title,
                "stats": "No scatter channels configured for this table.\n\nConfigure channels via Table Preferences.",
                "available_series": [],
                "series_visibility": {},
                "point_sets": [],
                "y_label": y_label,
            }

        custom_defs = self._custom_identifier_definitions()
        series_cache: dict[str, Any] = {}

        # Resolve the time column for playback support.
        time_column: str | None = None
        for col in self.log_df.columns:
            if str(col).strip().lower() == "time":
                time_column = str(col)
                break

        time_series_all: Any = None
        if time_column is not None:
            try:
                time_series_all = self._to_numeric_series(self.log_df[time_column])
            except Exception:
                time_series_all = None

        point_sets: list[dict[str, Any]] = []
        total_points = 0

        for channel_name, prefs in table_prefs.items():
            if not isinstance(prefs, dict):
                continue
            if not bool(prefs.get("show_in_scatterplot", False)):
                continue

            numeric_series = self._resolve_identifier_series(
                channel_name,
                custom_defs=custom_defs,
                cache=series_cache,
                stack=set(),
            )
            if numeric_series is None:
                continue

            rpm_vals: list[float] = []
            channel_vals: list[float] = []
            scatter_time_values: list[float | None] = []

            n = min(len(rpm_series), len(numeric_series))
            for i in range(n):
                try:
                    rpm_val = float(rpm_series.iloc[i])
                    ch_val = float(numeric_series.iloc[i])
                except Exception:
                    continue
                if rpm_val != rpm_val or ch_val != ch_val:  # NaN check
                    continue
                rpm_vals.append(rpm_val)
                channel_vals.append(ch_val)
                if time_series_all is not None:
                    try:
                        tv = float(time_series_all.iloc[i])
                        scatter_time_values.append(tv if tv == tv else None)
                    except Exception:
                        scatter_time_values.append(None)

            if not rpm_vals:
                continue

            total_points += len(rpm_vals)
            point_sets.append(
                {
                    "series_id": f"log::{channel_name}",
                    "name": channel_name,
                    "rpm": rpm_vals,
                    "ve": channel_vals,
                    "detail_channels": {},
                    "scatter_color": str(prefs.get("scatter_color", "")),
                    "scatter_opacity": max(0, min(100, int(prefs.get("scatter_opacity", 70)))),
                    "time": scatter_time_values,
                }
            )

        stats = f"All log data for {table_name}\nTotal points: {total_points}"
        return {
            "title": title,
            "stats": stats,
            "available_series": [],
            "series_visibility": {},
            "point_sets": point_sets,
            "y_label": y_label,
        }

    def _on_plot_all_data_requested(self) -> None:
        """Load all unfiltered log scatter data for the current table into the row viz panel."""
        payload = self._build_all_data_visualization_payload()
        if payload is None:
            return
        # Use a distinct dataset key so that a subsequent row selection always re-views.
        self._last_row_viz_dataset_key = ("all_data", self._current_table_name())
        self.table_row_panel.set_row_data(payload, auto_view_all=True)

    @staticmethod
    def _simple_linear_regression(
        x_vals: list[float], y_vals: list[float]
    ) -> tuple[float, float] | None:
        """Return (slope, intercept) for y = slope*x + intercept via least-squares.

        Returns None when fewer than 3 points or the x-values have no variance.
        """
        n = len(x_vals)
        if n < 3 or n != len(y_vals):
            return None
        sx = sum(x_vals)
        sy = sum(y_vals)
        sxx = sum(xi * xi for xi in x_vals)
        sxy = sum(xi * yi for xi, yi in zip(x_vals, y_vals))
        denom = n * sxx - sx * sx
        if abs(denom) < 1e-12:
            return None
        slope = (n * sxy - sx * sy) / denom
        intercept = (sy - slope * sx) / n
        return slope, intercept

    def _apply_afr_prediction_to_payload(
        self,
        table_name: str,
        point_sets: list[dict[str, Any]],
        table_rpm_values: list[float],
        map_mask: Any,
        filtered_rpm_values: list[float],
    ) -> None:
        """Compute predicted AFR for each table RPM column and store in extra_channels.

        For each column, this gathers ALL log data points inside the MAP band
        whose RPM is within tolerance, then fits a local linear regression of
        AFR vs VE Actual.  The regression line is evaluated at the current table
        VE value to produce a smooth predicted AFR that changes predictably as
        the user edits the VE cell.  When fewer than 3 points are available for
        regression the simple mean AFR of the neighbourhood is used instead.
        """
        if self.log_df is None or self.log_df.empty or not point_sets:
            return

        # Check per-table enable flag
        table_plugins = self.table_log_channel_preferences.get("__table_plugins__", {})
        if not isinstance(table_plugins, dict):
            return
        table_plugin_map = table_plugins.get(table_name, {})
        if not isinstance(table_plugin_map, dict) or not bool(table_plugin_map.get("afr_prediction", False)):
            return

        # Get identifier names from Advanced preferences
        adv = self.table_log_channel_preferences.get("__advanced__", {})
        if not isinstance(adv, dict):
            return
        afr_cfg = adv.get("afr_prediction", {})
        if not isinstance(afr_cfg, dict):
            return
        ve_actual_id = str(afr_cfg.get("ve_actual_identifier", "")).strip()
        afr_id = str(afr_cfg.get("afr_identifier", "")).strip()
        if not ve_actual_id or not afr_id:
            return

        # Resolve both series from the log
        custom_defs = self._custom_identifier_definitions()
        cache: dict[str, Any] = {}
        ve_actual_series = self._resolve_identifier_series(ve_actual_id, custom_defs, cache, set())
        afr_series = self._resolve_identifier_series(afr_id, custom_defs, cache, set())
        if ve_actual_series is None or afr_series is None:
            return

        # Filter to the rows matching the current MAP band
        try:
            ve_actual_filtered = list(ve_actual_series[map_mask])
            afr_filtered = list(afr_series[map_mask])
        except Exception:
            return

        n_log = min(len(filtered_rpm_values), len(ve_actual_filtered), len(afr_filtered))
        if n_log == 0:
            return

        rpm_tol = 150.0

        table_series = point_sets[0]
        table_ve_values = table_series.get("ve", [])

        predicted_afr: list[float | None] = []
        predicted_afr_counts: list[int | None] = []
        for i, rpm_val in enumerate(table_rpm_values):
            ve_val = float(table_ve_values[i]) if i < len(table_ve_values) else None

            # Gather ALL log points near this RPM (no VE filter)
            local_ve: list[float] = []
            local_afr: list[float] = []
            for j in range(n_log):
                if abs(float(filtered_rpm_values[j]) - rpm_val) > rpm_tol:
                    continue
                try:
                    log_ve = float(ve_actual_filtered[j])
                    afr_val = float(afr_filtered[j])
                    # Exclude NaN
                    if log_ve != log_ve or afr_val != afr_val:
                        continue
                    local_ve.append(log_ve)
                    local_afr.append(afr_val)
                except Exception:
                    continue

            if not local_afr:
                predicted_afr.append(None)
                predicted_afr_counts.append(None)
                continue

            if ve_val is None:
                # No table VE to evaluate against; use simple mean
                predicted_afr.append(sum(local_afr) / len(local_afr))
                predicted_afr_counts.append(len(local_afr))
                continue

            # Try local linear regression of AFR vs VE Actual
            fit = self._simple_linear_regression(local_ve, local_afr)
            if fit is not None:
                slope, intercept = fit
                predicted_afr.append(slope * ve_val + intercept)
            else:
                # Not enough data for regression; fall back to mean
                predicted_afr.append(sum(local_afr) / len(local_afr))
            predicted_afr_counts.append(len(local_afr))

        if any(v is not None for v in predicted_afr):
            extra_channels = table_series.setdefault("extra_channels", {})
            extra_channels["Predicted AFR"] = predicted_afr
            count_channels = table_series.setdefault("count_channels", {})
            count_channels["Prediction Points"] = predicted_afr_counts

    def _apply_table_log_channel_preferences_to_payload(
        self,
        table_name: str,
        point_sets: list[dict[str, Any]],
        table_rpm_values: list[float],
        map_mask: Any,
        filtered_rpm_values: list[float],
        filtered_map_values: list[float],
        map_value: float | None = None,
        map_series: Any = None,
        rpm_series: Any = None,
    ) -> None:
        if self.log_df is None or self.log_df.empty or not point_sets:
            return

        table_prefs = self.table_log_channel_preferences.get(table_name, {})
        if not isinstance(table_prefs, dict) or not table_prefs:
            return

        table_series = point_sets[0]
        extra_channels = table_series.get("extra_channels", {})
        if not isinstance(extra_channels, dict):
            extra_channels = {}
        custom_defs = self._custom_identifier_definitions()
        series_cache: dict[str, Any] = {}

        # Resolve log time column for playback support
        time_column: str | None = None
        for col in self.log_df.columns:
            if str(col).strip().lower() == "time":
                time_column = str(col)
                break

        for channel_name, prefs in table_prefs.items():
            if not isinstance(prefs, dict):
                continue

            show_scatter = bool(prefs.get("show_in_scatterplot", False))
            if not show_scatter:
                continue

            numeric_series = self._resolve_identifier_series(
                channel_name,
                custom_defs=custom_defs,
                cache=series_cache,
                stack=set(),
            )
            if numeric_series is None:
                continue

            # Compute per-channel MAP mask if a tolerance override is set
            tol_override = int(prefs.get("map_tolerance", 0))
            if tol_override > 0 and map_value is not None and map_series is not None and rpm_series is not None:
                ch_map_mask = (map_series >= map_value - tol_override) & (map_series <= map_value + tol_override)
                ch_filtered_rpm_values = [float(v) for v in rpm_series[ch_map_mask]]
                ch_filtered_map_values = [float(v) for v in map_series[ch_map_mask]]
            else:
                ch_map_mask = map_mask
                ch_filtered_rpm_values = filtered_rpm_values
                ch_filtered_map_values = filtered_map_values

            filtered_series = numeric_series[ch_map_mask]
            filtered_channel_values: list[float | None] = []
            for raw_value in filtered_series:
                value = float(raw_value)
                filtered_channel_values.append(value if value == value else None)

            valid_points = [
                (float(ch_filtered_rpm_values[idx]), float(ch_filtered_map_values[idx]), float(val))
                for idx, val in enumerate(filtered_channel_values)
                if idx < len(ch_filtered_rpm_values) and idx < len(ch_filtered_map_values) and val is not None
            ]

            if show_scatter and valid_points:
                source_indices = [
                    idx
                    for idx, val in enumerate(filtered_channel_values)
                    if idx < len(ch_filtered_rpm_values) and idx < len(ch_filtered_map_values) and val is not None
                ]
                detail_channels: dict[str, list[float | None]] = {}
                for detail_name in self._point_cloud_detail_channels(channel_name):
                    detail_series = self._resolve_identifier_series(
                        detail_name,
                        custom_defs=custom_defs,
                        cache=series_cache,
                        stack=set(),
                    )
                    if detail_series is None:
                        continue
                    detail_filtered = detail_series[ch_map_mask]
                    detail_filtered_values: list[float | None] = []
                    for raw_value in detail_filtered:
                        detail_value = float(raw_value)
                        detail_filtered_values.append(detail_value if detail_value == detail_value else None)
                    detail_values: list[float | None] = []
                    for source_idx in source_indices:
                        if source_idx >= len(detail_filtered_values):
                            detail_values.append(None)
                            continue
                        detail_values.append(detail_filtered_values[source_idx])
                    detail_channels[detail_name] = detail_values

                # Extract time values for playback
                scatter_time_values: list[float | None] = []
                if time_column is not None:
                    try:
                        time_series = self._to_numeric_series(self.log_df[time_column])
                        if time_series is not None:
                            time_filtered = list(time_series[ch_map_mask])
                            for source_idx in source_indices:
                                if source_idx < len(time_filtered):
                                    tv = float(time_filtered[source_idx])
                                    scatter_time_values.append(tv if tv == tv else None)
                                else:
                                    scatter_time_values.append(None)
                    except Exception:
                        scatter_time_values = []

                point_sets.append(
                    {
                        "series_id": f"log::{channel_name}",
                        "name": channel_name,
                        "rpm": [p[0] for p in valid_points],
                        "ve": [p[2] for p in valid_points],
                        "map": [p[1] for p in valid_points],
                        "ve_raw": [p[2] for p in valid_points],
                        "ve_scaled": [p[2] for p in valid_points],
                        "detail_channels": detail_channels,
                        "scatter_color": str(prefs.get("scatter_color", "")),
                        "scatter_opacity": max(0, min(100, int(prefs.get("scatter_opacity", 70)))),
                        "time": scatter_time_values,
                    }
                )

            if valid_points:
                valid_rpm_values = [p[0] for p in valid_points]
                valid_channel_values = [p[2] for p in valid_points]
                extra_channels[channel_name] = [
                    valid_channel_values[
                        min(
                            range(len(valid_rpm_values)),
                            key=lambda i: abs(float(valid_rpm_values[i]) - float(rpm_value)),
                        )
                    ]
                    for rpm_value in table_rpm_values
                ]

        if extra_channels:
            table_series["extra_channels"] = extra_channels

    def _update_table_grid_row_controls(self) -> None:
        available, message = self._table_row_editing_available()
        if not available:
            self._clear_table_row_editor(message)
            return

        if self._is_current_table_1d():
            self._load_selected_table_row(0)
            return

        if self.selected_table_row_idx is None:
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        if self.current_y_axis is None or self.selected_table_row_idx >= len(self.current_y_axis.values):
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        display_row = self._source_row_to_display_row(self.selected_table_row_idx)
        if display_row is None:
            self._clear_table_row_editor(message)
            return

        self.table_grid.blockSignals(True)
        self.table_grid.selectRow(display_row)
        self.table_grid.blockSignals(False)
        self._load_selected_table_row(self.selected_table_row_idx)

    def _interpolate_average_value(self, x_points: list[float], y_points: list[float], target_x: float) -> float:
        if len(x_points) == 1:
            return float(y_points[0])
        if target_x <= x_points[0]:
            return float(y_points[0])
        if target_x >= x_points[-1]:
            return float(y_points[-1])

        for idx in range(1, len(x_points)):
            left_x = float(x_points[idx - 1])
            right_x = float(x_points[idx])
            if target_x <= right_x:
                left_y = float(y_points[idx - 1])
                right_y = float(y_points[idx])
                span = right_x - left_x
                if abs(span) < 1e-9:
                    return right_y
                fraction = (target_x - left_x) / span
                return left_y + (right_y - left_y) * fraction

        return float(y_points[-1])

    def _on_row_editor_context_menu(self, pos) -> None:
        _ = pos
        return

    def _update_table_grid_row_visualization(self, include_row_plot: bool = True) -> None:
        self._stage_current_row_edits()
        self._refresh_current_table_diff_previews()

        available, message = self._table_row_editing_available()
        if not available:
            self._clear_table_row_editor(message)
            return

        if not include_row_plot:
            return

        if self.current_table is not None and self._is_current_table_1d():
            table_type = self._row_table_type()
            values_for_plot = [float(v) for v in self.pending_row_values] if self.pending_row_values else None
            payload = self._build_1d_table_visualization_payload(values_for_plot)
            dataset_key = ("1d", self.current_table.name)
            auto_view_all = dataset_key != self._last_row_viz_dataset_key
            payload["table_type"] = table_type
            payload["cursor_x_label"] = str(payload.get("x_label", "X"))
            payload["cursor_y_label"] = str(payload.get("y_label", "Value"))
            payload["available_series"] = self._row_viz_available_series(table_type, payload)
            payload["series_visibility"] = self.row_viz_preferences.get(table_type, {})

            selected_table_point_index = self.table_row_panel.selected_table_point_index()
            self.table_row_panel.set_row_data(payload, auto_view_all=auto_view_all)
            if not auto_view_all and selected_table_point_index is not None:
                self.table_row_panel.select_table_point(selected_table_point_index)

            self._last_row_viz_dataset_key = dataset_key
            self.table_row_label.setText("Selected row: full 1D table")
            self._set_row_editor_button_visibility(
                show_generate=False,
                show_write=False,
                show_revert=bool(self.pending_row_values),
            )
            self.table_row_status.setText(
                f"Editing full 1D table values for {self.current_table.name}. "
                "Use scroll wheel, arrow keys, and plot point selection to stage changes."
            )
            return

        if self.selected_table_row_idx is None or not self.pending_row_values:
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        if self.current_y_axis is None or self.current_x_axis is None or self.current_table is None:
            self._clear_table_row_editor("Select a table row in the table grid to view and edit it.")
            return

        dataset_key = ("2d", self.current_table.name, int(self.selected_table_row_idx))
        auto_view_all = dataset_key != self._last_row_viz_dataset_key

        table_name = self.current_table.name.lower()
        table_type = self._row_table_type()
        map_value = float(self.current_y_axis.values[self.selected_table_row_idx])
        cursor_x_label = self._axis_title("X", self.current_x_axis)
        cursor_y_label = self._axis_title("Y", self.current_y_axis)
        payload = self._build_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
        
        payload["available_series"] = self._row_viz_available_series(table_type, payload)
        payload["series_visibility"] = self.row_viz_preferences.get(table_type, {})
        payload["table_type"] = table_type
        payload["cursor_x_label"] = cursor_x_label
        payload["cursor_y_label"] = cursor_y_label

        selected_table_point_index = self.table_row_panel.selected_table_point_index()
        self.table_row_panel.set_row_data(payload, auto_view_all=auto_view_all)
        if not auto_view_all and selected_table_point_index is not None:
            self.table_row_panel.select_table_point(selected_table_point_index)
        self._last_row_viz_dataset_key = dataset_key
        y_label = self._axis_title("Y", self.current_y_axis)
        status_message = f"Editing {y_label} = {map_value:g} for {self.current_table.name}"
        self.table_row_status.setText(status_message)

    def _schedule_table_grid_row_visualization_update(self, delay_ms: int = 60) -> None:
        # Coalesce rapid editor updates so large log datasets do not stall the UI.
        self._row_viz_refresh_timer.start(max(0, int(delay_ms)))

    def _apply_deferred_row_visualization_update(self) -> None:
        self._update_table_grid_row_visualization(include_row_plot=True)

    def _refresh_current_table_diff_previews(self) -> None:
        if self.current_table is None:
            return

        base_matrix, _, _ = self._table_display_state(self.current_table, self.current_x_axis, self.current_y_axis)
        modified_matrix = self._table_matrix_with_pending_edits(self.current_table, self.current_x_axis, self.current_y_axis)
        diff_cells = self._diff_cells(base_matrix, modified_matrix)
        self._refresh_main_table_diff_highlight(base_matrix, diff_cells)
        self._refresh_modified_table_preview(
            self.current_table,
            self.current_x_axis,
            self.current_y_axis,
            modified_matrix,
            diff_cells,
        )

    def _refresh_main_table_diff_highlight(
        self,
        base_matrix: list[list[float]],
        diff_cells: set[tuple[int, int]],
    ) -> None:
        if not base_matrix:
            return

        row_count = len(base_matrix)
        col_count = len(base_matrix[0]) if base_matrix else 0
        if self.table_grid.rowCount() < row_count or self.table_grid.columnCount() < col_count:
            return

        display_matrix = list(reversed(base_matrix))
        minimum, maximum = self._matrix_min_max(display_matrix)
        span = max(maximum - minimum, 1e-9)

        for display_row, row_values in enumerate(display_matrix):
            source_row = row_count - 1 - display_row
            for col_index, value in enumerate(row_values):
                item = self.table_grid.item(display_row, col_index)
                if item is None:
                    continue
                if (source_row, col_index) in diff_cells:
                    item.setBackground(QColor(210, 65, 65, 210))
                else:
                    item.setBackground(self._cell_color(value, minimum, span))

    def _table_row_editing_available(self) -> tuple[bool, str]:
        if not self.current_table or not self.current_x_axis or not self.current_y_axis:
            return False, "Select a table and click a row to view and edit it."

        if self._is_current_table_1d():
            return True, ""


        return True, ""

    def _clear_table_row_editor(self, status_message: str) -> None:
        self._flush_active_global_cell_edit()
        self.selected_table_row_idx = None
        self.pending_row_values = []
        self.row_default_values = []
        self.row_edit_undo_stack = []
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None
        self.modified_table_meta.setVisible(False)
        self.modified_table_grid.setVisible(False)
        self.modified_table_grid.clear()
        self.modified_table_grid.setRowCount(0)
        self.modified_table_grid.setColumnCount(0)
        self.table_row_label.setText("Selected row: none")
        self.table_row_editor.clear()
        self.table_row_editor.setRowCount(1)
        self.table_row_editor.setColumnCount(0)
        self.table_row_editor.setEnabled(False)
        self._set_row_editor_button_visibility(
            show_generate=False,
            show_write=False,
            show_revert=False,
        )
        self._last_row_viz_dataset_key = None
        self.table_row_panel.clear_visualization()
        self.table_row_status.setText(status_message)

    def _display_row_to_source_row(self, display_row: int) -> int | None:
        if self.current_y_axis is None:
            return None
        row_count = len(self.current_y_axis.values)
        if display_row < 0 or display_row >= row_count:
            return None
        return row_count - 1 - display_row

    def _source_row_to_display_row(self, source_row: int) -> int | None:
        if self.current_y_axis is None:
            return None
        row_count = len(self.current_y_axis.values)
        if source_row < 0 or source_row >= row_count:
            return None
        return row_count - 1 - source_row

    def _on_table_grid_row_selection_changed(self) -> None:
        available, message = self._table_row_editing_available()
        if not available:
            self._clear_table_row_editor(message)
            return

        if self._is_current_table_1d():
            self._update_table_grid_row_visualization()
            return

        current_row = self.table_grid.currentRow()
        if current_row < 0:
            selected_items = self.table_grid.selectedItems()
            if selected_items:
                current_row = selected_items[0].row()

        if current_row < 0:
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        source_row = self._display_row_to_source_row(current_row)
        if source_row is None:
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        self._load_selected_table_row(source_row, preferred_column=None, select_editor_cell=False)

    def _on_table_grid_cell_clicked(self, row: int, column: int) -> None:
        source_row = self._display_row_to_source_row(row)
        if source_row is None:
            return
        editor_column = self._table_cell_to_editor_column(source_row, column)
        self._load_selected_table_row(source_row, preferred_column=editor_column, select_editor_cell=False)

    def _on_table_grid_current_cell_changed(
        self,
        current_row: int,
        current_column: int,
        previous_row: int,
        previous_column: int,
    ) -> None:
        _ = (current_column, previous_row, previous_column)
        if self._is_current_table_1d():
            self._update_table_grid_row_visualization()
            return
        source_row = self._display_row_to_source_row(current_row)
        if source_row is None:
            return
        self._load_selected_table_row(source_row, preferred_column=None, select_editor_cell=False)

    def _table_cell_to_editor_column(self, source_row: int, source_column: int) -> int | None:
        if self.current_table is None:
            return None

        if self.current_table.rows == 1:
            return source_column if 0 <= source_column < self.current_table.cols else None
        if self.current_table.cols == 1:
            return source_row if 0 <= source_row < self.current_table.rows else None
        return source_column if 0 <= source_column < self.current_table.cols else None

    def _load_selected_table_row(
        self,
        source_row: int,
        preferred_column: int | None = None,
        select_editor_cell: bool = True,
    ) -> None:
        if self.current_table is None or self.current_x_axis is None or self.current_y_axis is None:
            return

        self._flush_active_global_cell_edit()

        table_name = self._current_table_name()
        if table_name is None:
            return

        # Preserve edits for the previously active row before switching rows.
        self._stage_current_row_edits()

        if self._is_current_table_1d():
            self.selected_table_row_idx = 0
            baseline_values = self._row_values_for_table(self.current_table)
            staged_values = self._staged_row_values(table_name, self.current_table, 0)
            self.pending_row_values = staged_values if staged_values is not None else baseline_values
            self.row_default_values = [float(value) for value in baseline_values]
            self.row_edit_undo_stack = []
            self._active_row_edit_column = None
            self._active_row_edit_snapshot = None
            self.table_row_label.setText("Selected row: full 1D table")
            self._refresh_table_row_editor(preferred_column=preferred_column, select_cell=select_editor_cell)
            self._update_table_grid_row_visualization()
            if not select_editor_cell and preferred_column is not None:
                self.table_row_panel.select_table_point(preferred_column, emit_callback=False)
            return

        self.selected_table_row_idx = source_row
        baseline_values = self._row_values_for_table(self.current_table, source_row)
        staged_values = self._staged_row_values(table_name, self.current_table, source_row)
        self.pending_row_values = staged_values if staged_values is not None else baseline_values
        self.row_default_values = [float(value) for value in baseline_values]
        self.row_edit_undo_stack = []
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None
        map_value = float(self.current_y_axis.values[source_row])
        self.table_row_label.setText(f"Selected row: MAP {map_value:g} kPa")
        self._refresh_table_row_editor(preferred_column=preferred_column, select_cell=select_editor_cell)
        self._update_table_grid_row_visualization()
        if not select_editor_cell and preferred_column is not None:
            self.table_row_panel.select_table_point(preferred_column, emit_callback=False)

    def _refresh_table_row_editor(self, preferred_column: int | None = None, select_cell: bool = True) -> None:
        self.table_row_editor.blockSignals(True)
        self.table_row_editor.clear()
        self.table_row_editor.setRowCount(1)
        self.table_row_editor.setColumnCount(len(self.pending_row_values))
        minimum, maximum = self._matrix_min_max([self.pending_row_values])
        span = max(maximum - minimum, 1e-9)
        axis_values = self._active_row_editor_axis_values()
        if axis_values and len(axis_values) == len(self.pending_row_values):
            self.table_row_editor.setHorizontalHeaderLabels([f"{float(value):g}" for value in axis_values])
        for column_index, value in enumerate(self.pending_row_values):
            item = QTableWidgetItem(f"{value:g}")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(self._cell_color(value, minimum, span))
            self.table_row_editor.setItem(0, column_index, item)
        if self.pending_row_values and select_cell:
            if preferred_column is not None and 0 <= preferred_column < len(self.pending_row_values):
                target_column = preferred_column
            else:
                current_column = self.table_row_editor.currentColumn()
                target_column = current_column if 0 <= current_column < len(self.pending_row_values) else 0
            self.table_row_editor.setCurrentCell(0, target_column)
        else:
            self.table_row_editor.clearSelection()
        self.table_row_editor.resizeColumnsToContents()
        self.table_row_editor.resizeRowsToContents()
        self.table_row_editor.setEnabled(True)
        self._set_row_editor_button_visibility(
            show_generate=False,
            show_write=False,
            show_revert=bool(self.pending_row_values),
        )
        self.table_row_editor.blockSignals(False)
        self._apply_row_editor_selection_highlight()
        self._sync_editor_selection_to_row_plot()

    def _commit_pending_row_undo_state(self) -> None:
        if self._active_row_edit_snapshot is not None:
            self.row_edit_undo_stack.append([float(value) for value in self._active_row_edit_snapshot])
        else:
            self.row_edit_undo_stack.append([float(value) for value in self.pending_row_values])
        if len(self.row_edit_undo_stack) > 100:
            self.row_edit_undo_stack = self.row_edit_undo_stack[-100:]
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None

    def _adjust_pending_row_value(self, column_index: int, amount: float) -> None:
        if column_index < 0 or column_index >= len(self.pending_row_values):
            return
        previous_value = float(self.pending_row_values[column_index])
        new_value = float(previous_value + amount)
        self.pending_row_values[column_index] = new_value

        table_name = self._current_table_name()
        cell_coords = self._current_editor_cell_coordinates(column_index)
        if table_name is not None and cell_coords is not None:
            row_index, source_col = cell_coords
            self._record_global_cell_edit(table_name, row_index, source_col, previous_value, new_value)

        item = self.table_row_editor.item(0, column_index)
        if item is None:
            item = QTableWidgetItem()
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_row_editor.setItem(0, column_index, item)
        item.setText(f"{self.pending_row_values[column_index]:g}")
        self._update_table_grid_row_visualization(include_row_plot=False)
        self._schedule_table_grid_row_visualization_update()

    def _undo_pending_row_edit(self, column_index: int | None = None) -> None:
        _ = column_index
        self._trigger_global_undo()

    def _redo_pending_row_edit(self, column_index: int | None = None) -> None:
        _ = column_index
        self._trigger_global_redo()

    def _on_scatter_preferences_requested(self, identifier: str | None) -> None:
        self._open_preferences(initial_scatter_identifier=identifier)

    def _on_table_preferences_requested(self) -> None:
        self._open_preferences()

    def _on_table_row_plot_point_selected(
        self,
        series_id: str | None,
        index: int | None,
        rpm_value: float | None,
    ) -> None:
        if self.table_row_editor.columnCount() == 0:
            return
        if series_id is None:
            return
        if series_id != "table":
            # Allow non-table points (raw/corrected/AFR/etc.) to stay selected in
            # the plot without forcing editor focus back to the Selected Row line.
            return

        target_column: int | None = None
        if series_id == "table" and index is not None and 0 <= index < self.table_row_editor.columnCount():
            target_column = index
        elif rpm_value is not None:
            axis_values = self._active_row_editor_axis_values()
            if not axis_values:
                return
            target_column = min(
                range(len(axis_values)),
                key=lambda i: abs(float(axis_values[i]) - float(rpm_value)),
            )

        if target_column is None or target_column < 0 or target_column >= self.table_row_editor.columnCount():
            return

        self.table_row_editor.setCurrentCell(0, target_column)
        self.table_row_editor.setFocus()

    def _on_table_row_plot_point_adjust_requested(self, index: int, amount: float) -> None:
        if index < 0 or index >= len(self.pending_row_values):
            return
        self.table_row_editor.setCurrentCell(0, index)
        self._adjust_pending_row_value(index, amount)

    def _on_selected_table_cell_adjust_requested(self, display_row: int, column: int, amount: float) -> None:
        available, _ = self._table_row_editing_available()
        if not available:
            return

        if self.current_table is None or self.current_x_axis is None:
            return

        if column < 0 or column >= self.current_x_axis.length:
            return

        if self._is_current_table_1d():
            self._load_selected_table_row(0, preferred_column=column)
            self._adjust_pending_row_value(column, amount)
            return

        source_row = self._display_row_to_source_row(display_row)
        if source_row is None:
            return

        self._load_selected_table_row(source_row, preferred_column=column)
        self._adjust_pending_row_value(column, amount)

    def _on_selected_table_cell_undo_requested(self, display_row: int, column: int) -> None:
        available, _ = self._table_row_editing_available()
        if not available:
            return

        if self.current_table is None or self.current_x_axis is None:
            return

        if column < 0 or column >= self.current_x_axis.length:
            return

        if self._is_current_table_1d():
            self._load_selected_table_row(0, preferred_column=column)
        else:
            source_row = self._display_row_to_source_row(display_row)
            if source_row is None:
                return
            self._load_selected_table_row(source_row, preferred_column=column)
        self._trigger_global_undo()

    def _on_selected_table_cell_redo_requested(self, display_row: int, column: int) -> None:
        available, _ = self._table_row_editing_available()
        if not available:
            return

        if self.current_table is None or self.current_x_axis is None:
            return

        if column < 0 or column >= self.current_x_axis.length:
            return

        if self._is_current_table_1d():
            self._load_selected_table_row(0, preferred_column=column)
        else:
            source_row = self._display_row_to_source_row(display_row)
            if source_row is None:
                return
            self._load_selected_table_row(source_row, preferred_column=column)
        self._trigger_global_redo()

    def _revert_pending_row_values(self) -> None:
        if not self.row_default_values:
            return
        self._commit_pending_row_undo_state()
        self.pending_row_values = [float(value) for value in self.row_default_values]
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None

        table_name = self._current_table_name()
        if table_name is not None:
            state = self.pending_edits_per_table.get(table_name)
            if state is not None:
                if self.current_table is not None and self._is_current_table_1d():
                    state["one_d"] = None
                elif self.selected_table_row_idx is not None:
                    state.get("rows", {}).pop(self.selected_table_row_idx, None)
                if not state.get("rows") and state.get("one_d") is None:
                    self.pending_edits_per_table.pop(table_name, None)

        self._refresh_table_row_editor()
        self._update_table_grid_row_visualization()

    def _apply_pending_row_changes(self) -> None:
        if (
            self.current_table is None
            or self.current_x_axis is None
            or self.current_y_axis is None
            or not self.pending_row_values
        ):
            return

        if self._is_current_table_1d():
            if self.current_table.rows == 1:
                if len(self.pending_row_values) != self.current_table.cols:
                    return
                self.current_table.values[0] = [float(value) for value in self.pending_row_values]
            else:
                if len(self.pending_row_values) != self.current_table.rows:
                    return
                for row_index, value in enumerate(self.pending_row_values):
                    self.current_table.values[row_index][0] = float(value)

            self.row_default_values = [float(value) for value in self.pending_row_values]
            self.row_edit_undo_stack = []
            self._render_table(self.current_table, self.current_x_axis, self.current_y_axis)
            self._update_table_grid_row_visualization()
            self.statusBar().showMessage(f"Updated 1D table values in {self.current_table.name}.")
            return

        if self.selected_table_row_idx is None:
            return

        self.current_table.values[self.selected_table_row_idx] = [float(value) for value in self.pending_row_values]
        self.row_default_values = [float(value) for value in self.pending_row_values]
        self.row_edit_undo_stack = []
        display_row = self._source_row_to_display_row(self.selected_table_row_idx)
        self._render_table(self.current_table, self.current_x_axis, self.current_y_axis)
        if display_row is not None:
            self.table_grid.blockSignals(True)
            self.table_grid.selectRow(display_row)
            self.table_grid.blockSignals(False)
        self._load_selected_table_row(self.selected_table_row_idx)
        map_value = float(self.current_y_axis.values[self.selected_table_row_idx])
        self.statusBar().showMessage(f"Updated MAP {map_value:g} kPa row in {self.current_table.name}.")

    def _open_tune_file(self) -> None:
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open MegaSquirt Tune File",
            str(self._default_data_dir()),
            "Tune Files (*.msq *.ini *.txt);;All Files (*.*)",
        )
        if not selected_path:
            return

        file_path = Path(selected_path)
        try:
            self.tune_data = self.tune_loader.load(file_path)
            self.selected_rows_per_table = {}  # Clear saved row selections for new tune
            self.pending_edits_per_table = {}  # Clear saved pending edits for new tune
            self.loaded_table_snapshots = self._snapshot_table_values(self.tune_data)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Tune Load Error", f"Could not load tune file:\n{exc}")
            return

        self.loaded_tune_path = file_path
        self.save_target_path = None
        self._set_tune_save_actions_enabled(True)
        self.table_list.clear()
        if self.tune_data.tables:
            self._update_table_display()
            self._select_default_table_on_load()
        else:
            self.table_list.addItem("No tables found")
            self.table_grid.clear()
            self.table_grid.setRowCount(0)
            self.table_grid.setColumnCount(0)
            self.table_meta.setText("No tables found in this tune file")
            self.table_meta.setToolTip("")
            self.axis_meta.setText("Axes: X (index), Y (index)")

        self.statusBar().showMessage(f"Loaded tune: {file_path.name}")
        self._refresh_workspace_text()
        self._add_recent_tune_file(file_path)
        self._maybe_prompt_load_matching_log(file_path)

    def _on_tune_file_dropped(self, file_path: Path) -> None:
        self._open_recent_tune_file(file_path)

    def _add_open_tune_file_placeholder(self) -> None:
        item = QListWidgetItem("Open Tune File...")
        item.setData(Qt.ItemDataRole.UserRole, self.OPEN_TUNE_PLACEHOLDER_KEY)
        item.setToolTip("Click to browse for a tune file, or drag and drop a tune file here.")
        self.table_list.addItem(item)
        self.table_list.clearSelection()
        self.table_list.setCurrentRow(-1)

    def _on_table_list_item_clicked(self, item: QListWidgetItem) -> None:
        if item.data(Qt.ItemDataRole.UserRole) == self.OPEN_TUNE_PLACEHOLDER_KEY:
            self._open_tune_file()

    def _maybe_prompt_load_matching_log(self, tune_file: Path) -> None:
        if tune_file.suffix.lower() != ".msq":
            return

        tune_stem = tune_file.stem.lower()
        matching_logs: list[Path] = []
        seen: set[Path] = set()

        # 1) Prefer exact same-name logs next to the tune file.
        for suffix in (".msl", ".mlg"):
            candidate = tune_file.with_suffix(suffix)
            if candidate.exists():
                resolved = candidate.resolve()
                if resolved not in seen:
                    matching_logs.append(resolved)
                    seen.add(resolved)

        # 2) Include same-stem logs from recent log history.
        for recent_log in self.recent_log_files:
            resolved = recent_log.resolve()
            if not resolved.exists():
                continue
            if resolved.suffix.lower() not in {".msl", ".mlg"}:
                continue
            if resolved.stem.lower() != tune_stem:
                continue
            if resolved not in seen:
                matching_logs.append(resolved)
                seen.add(resolved)

        # 3) Include same-stem logs from default tuning data directory.
        data_dir = self._default_data_dir()
        if data_dir.exists():
            for suffix in ("*.msl", "*.mlg"):
                for candidate in data_dir.glob(suffix):
                    resolved = candidate.resolve()
                    if resolved.stem.lower() != tune_stem:
                        continue
                    if resolved not in seen:
                        matching_logs.append(resolved)
                        seen.add(resolved)

        if not matching_logs:
            prompt = QMessageBox(self)
            prompt.setIcon(QMessageBox.Icon.Question)
            prompt.setWindowTitle("No Matching Log Found")
            prompt.setText(
                f"No matching .msl/.mlg log was found for {tune_file.name}.\n\n"
                "Would you like to browse and load a log file now?"
            )
            prompt.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            prompt.setDefaultButton(QMessageBox.StandardButton.Yes)
            prompt.setEscapeButton(QMessageBox.StandardButton.No)
            if prompt.exec() == QMessageBox.StandardButton.Yes:
                self._open_log_file()
            return

        preferred_log = next((p for p in matching_logs if p.suffix.lower() == ".msl"), matching_logs[0])

        details = "\n".join(f"- {p.name}" for p in matching_logs)
        message = (
            f"Found matching log file(s) for {tune_file.name}:\n\n"
            f"{details}\n\n"
            f"Load {preferred_log.name} now?"
        )

        prompt = QMessageBox(self)
        prompt.setIcon(QMessageBox.Icon.Question)
        prompt.setWindowTitle("Load Matching Log File")
        prompt.setText(message)
        prompt.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        prompt.setDefaultButton(QMessageBox.StandardButton.Yes)
        prompt.setEscapeButton(QMessageBox.StandardButton.No)

        if prompt.exec() == QMessageBox.StandardButton.Yes:
            self._open_recent_log_file(preferred_log)

    def _next_revision_path(self, path: Path) -> Path:
        """Return a path with _RevX suffix incremented (or _Rev1 appended if absent)."""
        stem = path.stem
        match = re.search(r"^(.*?)_Rev(\d+)$", stem, re.IGNORECASE)
        if match:
            base = match.group(1)
            rev = int(match.group(2)) + 1
        else:
            base = stem
            rev = 1
        return path.with_name(f"{base}_Rev{rev}{path.suffix}")

    def _save_tune_as(self) -> None:
        if self.tune_data is None or self.loaded_tune_path is None:
            return

        suggested = self._next_revision_path(self.loaded_tune_path)

        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Tune As",
            str(suggested),
            "Tune Files (*.msq);;All Files (*.*)",
        )
        if not dest_path:
            return

        dest = Path(dest_path)
        self._save_tune_to_path(dest)

    def _save_tune_to_path(self, dest: Path) -> None:
        if self.tune_data is None:
            QMessageBox.warning(self, "No Tune Loaded", "Load a tune file before saving.")
            return

        pre_save_snapshot = self._snapshot_table_values(self.tune_data)
        preserved_table_name = self._current_table_name()
        preserved_row_index = self.selected_table_row_idx
        preserved_undo_stack = [[float(value) for value in state] for state in self.row_edit_undo_stack]
        preserved_active_snapshot = (
            [float(value) for value in self._active_row_edit_snapshot]
            if self._active_row_edit_snapshot is not None
            else None
        )
        preserved_active_column = self._active_row_edit_column

        self._apply_all_pending_row_edits_to_tune_data()
        try:
            self.tune_loader.save(self.tune_data, dest)
        except Exception as exc:  # noqa: BLE001
            self._restore_table_values(self.tune_data, pre_save_snapshot)
            QMessageBox.critical(self, "Save Error", f"Could not save tune file:\n{exc}")
            return

        # Persist staged edits to disk, but keep the in-memory workspace exactly as it was pre-save.
        self._restore_table_values(self.tune_data, pre_save_snapshot)

        self.statusBar().showMessage(f"Saved tune: {dest.name}")
        self.loaded_tune_path = dest
        self.save_target_path = dest

        if self._current_table_name() == preserved_table_name and self.selected_table_row_idx == preserved_row_index:
            self.row_edit_undo_stack = preserved_undo_stack
            self._active_row_edit_snapshot = preserved_active_snapshot
            self._active_row_edit_column = preserved_active_column
        else:
            self.row_edit_undo_stack = []
            self._active_row_edit_snapshot = None
            self._active_row_edit_column = None

        self._add_recent_tune_file(dest)

    def _open_log_file(self) -> None:
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open MegaSquirt Log File",
            str(self._default_data_dir()),
            "Log Files (*.csv *.msl *.mlg *.txt);;All Files (*.*)",
        )
        if not selected_path:
            return

        file_path = Path(selected_path)
        self.statusBar().showMessage(f"Loading log: {file_path.name}...")
        self.setCursor(Qt.CursorShape.WaitCursor)
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            self.statusBar().showMessage(f"Failed to load log: {file_path.name}", 5000)
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return
        finally:
            self.setCursor(Qt.CursorShape.ArrowCursor)

        self.log_file = file_path
        self.statusBar().showMessage(
            f"Loaded log: {file_path.name} ({len(log_df):,} rows, parser={parse_result.parser_used}, enc={parse_result.encoding})"
        )
        self.log_df = log_df
        self._add_recent_log_file(file_path)

    def _refresh_workspace_text(self) -> None:
        """Workspace text refresh - disabled since workspace is hidden by default."""
        pass

    def _current_table_name(self) -> str | None:
        if self.current_table is None or self.tune_data is None:
            return None
        for table_name, table in self.tune_data.tables.items():
            if table is self.current_table:
                return table_name
        return None

    @staticmethod
    def _row_values_for_table(table: TableData, row_index: int | None = None) -> list[float]:
        if table.rows == 1:
            return [float(value) for value in table.values[0]]
        if table.cols == 1:
            return [float(row[0]) for row in table.values]
        if row_index is None or row_index < 0 or row_index >= table.rows:
            return []
        return [float(value) for value in table.values[row_index]]

    @staticmethod
    def _values_differ(values_a: list[float], values_b: list[float], tolerance: float = 1e-9) -> bool:
        if len(values_a) != len(values_b):
            return True
        return any(abs(float(a) - float(b)) > tolerance for a, b in zip(values_a, values_b))

    def _stage_current_row_edits(self) -> None:
        table_name = self._current_table_name()
        if table_name is None or self.current_table is None or not self.pending_row_values:
            return

        table_state = self.pending_edits_per_table.setdefault(
            table_name,
            {"rows": {}, "one_d": None},
        )
        rows_state = table_state.setdefault("rows", {})

        if self._is_current_table_1d():
            baseline_values = self._row_values_for_table(self.current_table)
            staged_values = [float(value) for value in self.pending_row_values]
            if self._values_differ(staged_values, baseline_values):
                table_state["one_d"] = staged_values
            else:
                table_state["one_d"] = None
        else:
            if self.selected_table_row_idx is None:
                return
            baseline_values = self._row_values_for_table(self.current_table, self.selected_table_row_idx)
            staged_values = [float(value) for value in self.pending_row_values]
            if self._values_differ(staged_values, baseline_values):
                rows_state[self.selected_table_row_idx] = staged_values
            else:
                rows_state.pop(self.selected_table_row_idx, None)

        if not table_state.get("rows") and table_state.get("one_d") is None:
            self.pending_edits_per_table.pop(table_name, None)

    def _staged_row_values(self, table_name: str, table: TableData, row_index: int | None) -> list[float] | None:
        state = self.pending_edits_per_table.get(table_name)
        if not state:
            return None

        if table.rows == 1 or table.cols == 1:
            one_d_values = state.get("one_d")
            if isinstance(one_d_values, list):
                return [float(value) for value in one_d_values]
            return None

        if row_index is None:
            return None
        rows_state = state.get("rows", {})
        row_values = rows_state.get(row_index)
        if isinstance(row_values, list):
            return [float(value) for value in row_values]
        return None

    def _on_table_selected(self, current_item: QListWidgetItem | None, previous_item: QListWidgetItem | None) -> None:
        # Save current row edits/selection before switching away.
        current_table_name = self._current_table_name()
        self._stage_current_row_edits()
        if current_table_name is not None and self.selected_table_row_idx is not None:
            self.selected_rows_per_table[current_table_name] = self.selected_table_row_idx

        if not current_item or not self.tune_data:
            self.current_table = None
            self.current_x_axis = None
            self.current_y_axis = None
            self.axis_meta.setText("Axes: X (index), Y (index)")
            self._update_table_grid_row_controls()
            return

        table_name = current_item.data(Qt.ItemDataRole.UserRole)
        if not table_name or table_name not in self.tune_data.tables:
            return

        table = self.tune_data.tables[table_name]
        self.selected_table_row_idx = None
        self.pending_row_values = []
        self.current_table = table
        self.current_x_axis, self.current_y_axis = self.tune_data.resolve_table_axes(table)

        if self.current_x_axis is None:
            self.current_x_axis = AxisVector(
                name="index_x",
                source_tag="synthetic",
                length=table.cols,
                orientation="row",
                units=None,
                digits=None,
                values=[float(i + 1) for i in range(table.cols)],
            )
        if self.current_y_axis is None:
            self.current_y_axis = AxisVector(
                name="index_y",
                source_tag="synthetic",
                length=table.rows,
                orientation="column",
                units=None,
                digits=None,
                values=[float(i + 1) for i in range(table.rows)],
            )
        self._render_table(table, self.current_x_axis, self.current_y_axis)
        self._update_table_grid_row_controls()

        # Restore the saved row selection for this table
        if table_name in self.selected_rows_per_table:
            saved_row_idx = self.selected_rows_per_table[table_name]
            # Validate that the saved row index is still valid for this table
            if 0 <= saved_row_idx < table.rows:
                self._load_selected_table_row(saved_row_idx)

        if table_name in self.pending_edits_per_table:
            self._refresh_current_table_view()

    def _refresh_current_table_view(self) -> None:
        if not self.current_table:
            return
        self._render_table(self.current_table, self.current_x_axis, self.current_y_axis)
        self._update_table_grid_row_controls()

    def _render_table(self, table: TableData, x_axis: AxisVector | None, y_axis: AxisVector | None) -> None:
        matrix, display_x_axis, display_y_axis = self._table_display_state(table, x_axis, y_axis)
        row_count = len(matrix)
        col_count = len(matrix[0]) if matrix else 0
        y_labels = list(reversed(self._header_labels(display_y_axis, row_count)))
        if (
            table.cols == 1
            and table.rows > 1
            and row_count == 1
        ):
            y_labels = ["Values"]
        x_labels = self._header_labels(display_x_axis, col_count)
        x_axis_title = self._axis_title("X", display_x_axis)
        y_axis_title = self._axis_title("Y", display_y_axis)

        modified_matrix = self._table_matrix_with_pending_edits(table, x_axis, y_axis)
        diff_cells = self._diff_cells(matrix, modified_matrix)

        self._populate_table_grid(self.table_grid, matrix, x_labels, y_labels, diff_cells=diff_cells)
        self._refresh_modified_table_preview(table, x_axis, y_axis, modified_matrix, diff_cells)
        QTimer.singleShot(0, self._auto_size_table_and_row_viz)

        units = table.units or "-"
        display_name = self._get_tunerstudio_name(table.name) if self.show_tunerstudio_names else table.name
        table_details = (
            f"{table.name} | {row_count}x{col_count} | units: {units} | "
            f"x(bottom): {x_axis_title} | y(left): {y_axis_title}"
        )
        self.table_meta.setText(display_name)
        self.table_meta.setToolTip(table_details)
        self.axis_meta.setText(f"Axes: X = {x_axis_title} | Y = {y_axis_title}")

    def _populate_table_grid(
        self,
        grid: QTableWidget,
        matrix: list[list[float]],
        x_labels: list[str],
        y_labels: list[str],
        diff_cells: set[tuple[int, int]] | None = None,
    ) -> None:
        row_count = len(matrix)
        col_count = len(matrix[0]) if matrix else 0
        display_matrix = list(reversed(matrix))
        diff_cells = diff_cells or set()

        grid.clear()
        grid.setRowCount(row_count + 1)
        grid.setColumnCount(col_count)
        if isinstance(grid, CopyPasteTableWidget):
            grid.set_footer_rows(1)
        grid.setVerticalHeaderLabels(y_labels + [""])

        minimum, maximum = self._matrix_min_max(display_matrix)
        span = max(maximum - minimum, 1e-9)
        for display_row, row in enumerate(display_matrix):
            source_row = row_count - 1 - display_row
            for col, value in enumerate(row):
                item = QTableWidgetItem(f"{value:g}")
                if (source_row, col) in diff_cells:
                    item.setBackground(QColor(210, 65, 65, 210))
                else:
                    item.setBackground(self._cell_color(value, minimum, span))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                grid.setItem(display_row, col, item)

        axis_row = row_count
        for col, label in enumerate(x_labels):
            axis_item = QTableWidgetItem(label)
            axis_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            axis_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            axis_item.setBackground(QColor(60, 60, 60))
            grid.setItem(axis_row, col, axis_item)

        grid.resizeColumnsToContents()
        grid.resizeRowsToContents()

    def _refresh_modified_table_preview(
        self,
        table: TableData,
        x_axis: AxisVector | None,
        y_axis: AxisVector | None,
        modified_matrix: list[list[float]] | None = None,
        diff_cells: set[tuple[int, int]] | None = None,
    ) -> None:
        base_matrix, display_x_axis, display_y_axis = self._table_display_state(table, x_axis, y_axis)
        row_count = len(base_matrix)
        col_count = len(base_matrix[0]) if base_matrix else 0
        y_labels = list(reversed(self._header_labels(display_y_axis, row_count)))
        if (
            table.cols == 1
            and table.rows > 1
            and row_count == 1
        ):
            y_labels = ["Values"]
        x_labels = self._header_labels(display_x_axis, col_count)

        if modified_matrix is None:
            modified_matrix = self._table_matrix_with_pending_edits(table, x_axis, y_axis)
        if diff_cells is None:
            diff_cells = self._diff_cells(base_matrix, modified_matrix)

        if not diff_cells:
            self.modified_table_meta.setVisible(False)
            self.modified_table_grid.setVisible(False)
            self.modified_table_grid.clear()
            self.modified_table_grid.setRowCount(0)
            self.modified_table_grid.setColumnCount(0)
            return

        display_name = self._get_tunerstudio_name(table.name) if self.show_tunerstudio_names else table.name
        self.modified_table_meta.setText(f"{display_name} (modified)")
        self.modified_table_meta.setVisible(True)
        self.modified_table_grid.setVisible(True)
        self._populate_table_grid(self.modified_table_grid, modified_matrix, x_labels, y_labels, diff_cells=diff_cells)

    def _table_matrix_with_pending_edits(
        self,
        table: TableData,
        x_axis: AxisVector | None,
        y_axis: AxisVector | None,
    ) -> list[list[float]]:
        source_values = [[float(value) for value in row] for row in table.values]
        table_name = None
        if self.tune_data is not None:
            for name, candidate in self.tune_data.tables.items():
                if candidate is table:
                    table_name = name
                    break

        if table_name is not None:
            state = self.pending_edits_per_table.get(table_name, {})
            one_d_values = state.get("one_d")
            row_overrides = state.get("rows", {}) if isinstance(state.get("rows", {}), dict) else {}

            if table.rows == 1 and isinstance(one_d_values, list) and len(one_d_values) == table.cols:
                source_values[0] = [float(value) for value in one_d_values]
            elif table.cols == 1 and isinstance(one_d_values, list) and len(one_d_values) == table.rows:
                for row_index, value in enumerate(one_d_values):
                    source_values[row_index][0] = float(value)
            elif table.rows > 1 and table.cols > 1:
                for row_index, row_values in row_overrides.items():
                    if not isinstance(row_index, int) or not isinstance(row_values, list):
                        continue
                    if 0 <= row_index < table.rows and len(row_values) == table.cols:
                        source_values[row_index] = [float(value) for value in row_values]

        matrix, _, _ = self._table_display_state(table, x_axis, y_axis, source_values)
        return matrix

    @staticmethod
    def _diff_cells(base_matrix: list[list[float]], modified_matrix: list[list[float]]) -> set[tuple[int, int]]:
        if not base_matrix or not modified_matrix:
            return set()

        row_count = min(len(base_matrix), len(modified_matrix))
        col_count = min(len(base_matrix[0]), len(modified_matrix[0]))
        diff: set[tuple[int, int]] = set()
        for row in range(row_count):
            for col in range(col_count):
                if abs(float(base_matrix[row][col]) - float(modified_matrix[row][col])) > 1e-9:
                    diff.add((row, col))
        return diff

    @staticmethod
    def _snapshot_table_values(tune_data: TuneData) -> dict[str, list[list[float]]]:
        snapshots: dict[str, list[list[float]]] = {}
        for table_name, table in tune_data.tables.items():
            snapshots[table_name] = [[float(value) for value in row] for row in table.values]
        return snapshots

    @staticmethod
    def _restore_table_values(tune_data: TuneData, snapshots: dict[str, list[list[float]]]) -> None:
        for table_name, rows in snapshots.items():
            table = tune_data.tables.get(table_name)
            if table is None:
                continue
            if len(rows) != table.rows:
                continue
            for row_index, row_values in enumerate(rows):
                if len(row_values) != table.cols:
                    continue
                table.values[row_index] = [float(value) for value in row_values]

    def _apply_all_pending_row_edits_to_tune_data(self) -> None:
        if self.tune_data is None:
            return

        self._stage_current_row_edits()

        def apply_row_edit(table_name: str, row_index: int | None, pending_values: list[float]) -> None:
            table = self.tune_data.tables.get(table_name)
            if table is None or not pending_values:
                return

            if table.rows == 1:
                if len(pending_values) == table.cols:
                    table.values[0] = [float(value) for value in pending_values]
                return

            if table.cols == 1:
                if len(pending_values) == table.rows:
                    for r, value in enumerate(pending_values):
                        table.values[r][0] = float(value)
                return

            if row_index is None or row_index < 0 or row_index >= table.rows:
                return
            if len(pending_values) != table.cols:
                return
            table.values[row_index] = [float(value) for value in pending_values]

        for table_name, saved in self.pending_edits_per_table.items():
            table = self.tune_data.tables.get(table_name)
            if table is None:
                continue

            one_d_values = saved.get("one_d")
            if isinstance(one_d_values, list):
                apply_row_edit(table_name, 0, [float(value) for value in one_d_values])

            row_overrides = saved.get("rows", {}) if isinstance(saved.get("rows", {}), dict) else {}
            for row_index, row_values in row_overrides.items():
                if not isinstance(row_index, int) or not isinstance(row_values, list):
                    continue
                apply_row_edit(table_name, row_index, [float(value) for value in row_values])

    def _table_content_width(self) -> int:
        frame = self.table_grid.frameWidth() * 2
        header_width = self.table_grid.verticalHeader().width() if self.table_grid.verticalHeader().isVisible() else 0
        column_width = sum(self.table_grid.columnWidth(col) for col in range(self.table_grid.columnCount()))

        scrollbar_width = 0
        if self.table_grid.verticalScrollBar().isVisible():
            scrollbar_width = self.table_grid.verticalScrollBar().sizeHint().width()

        # Include a small buffer for grid lines and container margins.
        return frame + header_width + column_width + scrollbar_width + 16

    def _auto_size_table_and_row_viz(self) -> None:
        if not hasattr(self, "table_sections_splitter") or not hasattr(self, "table_details_splitter"):
            return
        if self.table_sections_splitter.count() < 2:
            return

        available_width = self.table_sections_splitter.size().width()
        if available_width <= 0:
            return

        desired_table_width = self._table_content_width()
        min_table_width = 260
        min_viz_width = 420

        max_table_width = max(min_table_width, available_width - min_viz_width)
        table_width = max(min_table_width, min(desired_table_width, max_table_width))
        viz_width = max(min_viz_width, available_width - table_width)
        if table_width + viz_width != available_width:
            table_width = max(min_table_width, available_width - viz_width)

        self.table_sections_splitter.setSizes([table_width, viz_width])

        details_height = self.table_details_splitter.size().height()
        if details_height > 0:
            editor_needed = self.table_row_editor_group.sizeHint().height() + self.table_row_status.sizeHint().height() + 8
            min_editor_height = max(110, editor_needed)
            max_editor_height = max(min_editor_height, int(details_height * 0.35))
            editor_height = min(min_editor_height, max_editor_height)
            row_viz_height = max(260, details_height - editor_height)
            if row_viz_height + editor_height > details_height:
                row_viz_height = max(220, details_height - editor_height)
            self.table_details_splitter.setSizes([row_viz_height, editor_height])

    def _table_display_state(
        self,
        table: TableData,
        x_axis: AxisVector | None,
        y_axis: AxisVector | None,
        source_values: list[list[float]] | None = None,
    ) -> tuple[list[list[float]], AxisVector | None, AxisVector | None]:
        if source_values is None:
            matrix = [row[:] for row in table.values]
        else:
            matrix = [row[:] for row in source_values]
        display_x_axis = x_axis
        display_y_axis = y_axis

        # Default 1D column tables (Nx1) to horizontal presentation (1xN).
        if table.cols == 1 and table.rows > 1:
            matrix = [[float(row[0]) for row in matrix]]
            display_x_axis = y_axis
            display_y_axis = AxisVector(
                name="index_y",
                source_tag="synthetic",
                length=1,
                orientation="column",
                units=None,
                digits=None,
                values=[1.0],
            )

        return matrix, display_x_axis, display_y_axis

    @staticmethod
    def _header_labels(axis: AxisVector | None, size: int) -> list[str]:
        labels: list[str] = []
        for i in range(size):
            if axis and i < len(axis.values):
                labels.append(f"{axis.values[i]:g}")
            else:
                labels.append(str(i + 1))
        return labels

    @staticmethod
    def _axis_title(default_name: str, axis: AxisVector | None) -> str:
        if not axis:
            return f"{default_name} (index)"
        friendly_name = MainWindow._friendly_axis_name(axis.name)
        units = f" ({axis.units})" if axis.units else ""
        return f"{friendly_name}{units}"

    @staticmethod
    def _friendly_axis_name(axis_name: str) -> str:
        normalized = axis_name.strip()
        lowered = normalized.lower()

        direct_map = {
            "rpm": "RPM",
            "map": "MAP",
            "maf": "MAF",
            "tps": "TPS",
            "afr": "AFR",
            "lambda": "Lambda",
            "load": "Load",
            "knock": "Knock",
            "index_x": "X",
            "index_y": "Y",
        }
        if lowered in direct_map:
            return direct_map[lowered]

        compact = re.sub(r"[^a-z0-9]+", "", lowered)

        if "rpm" in compact:
            return "RPM"
        if "map" in compact:
            return "MAP"
        if "maf" in compact:
            return "MAF"
        if "tps" in compact:
            return "TPS"
        if "afr" in compact:
            return "AFR"
        if "lambda" in compact:
            return "Lambda"
        if "knock" in compact or "knk" in compact:
            return "Knock"

        pretty = re.sub(r"[_\-]+", " ", normalized).strip()
        return pretty.title() if pretty else axis_name

    @staticmethod
    def _matrix_min_max(matrix: list[list[float]]) -> tuple[float, float]:
        if not matrix or not matrix[0]:
            return 0.0, 0.0
        minimum = min(min(row) for row in matrix)
        maximum = max(max(row) for row in matrix)
        return minimum, maximum

    @staticmethod
    def _cell_color(value: float, minimum: float, span: float) -> QColor:
        ratio = (value - minimum) / span
        ratio = max(0.0, min(1.0, ratio))
        if ratio < 0.5:
            local = ratio / 0.5
            red = int(30 + (30 * local))
            green = int(90 + (120 * local))
            blue = int(220 - (120 * local))
        else:
            local = (ratio - 0.5) / 0.5
            red = int(60 + (180 * local))
            green = int(210 - (120 * local))
            blue = int(100 - (70 * local))
        return QColor(red, green, blue, 140)

    @staticmethod
    def _to_numeric_series(series):
        import pandas as pd

        direct = pd.to_numeric(series, errors="coerce")
        if direct.notna().mean() >= 0.2:  # Lowered from 0.5 to 0.2 to accept columns with mixed data
            return direct

        cleaned = (
            series.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(r"[^0-9eE+\-\.]", "", regex=True)
        )
        coerced = pd.to_numeric(cleaned, errors="coerce")
        if coerced.notna().mean() >= 0.2:  # Lowered from 0.5 to 0.2 to accept columns with mixed data
            return coerced
        return None

    @staticmethod
    def _weighted_average(values: list[float], weights: list[float]) -> float | None:
        if not values or not weights or len(values) != len(weights):
            return None
        total_weight = sum(weights)
        if abs(total_weight) < 1e-12:
            return None
        return sum(v * w for v, w in zip(values, weights)) / total_weight

    @staticmethod
    def _predict_afr_from_neighbors(
        target_rpm: float,
        target_map: float,
        target_ve: float,
        rpm_values: list[float],
        map_values: list[float],
        ve_corrected_values: list[float],
        afr_values: list[float | None],
        neighbor_count: int = 24,
    ) -> float | None:
        if abs(target_ve) < 1e-9:
            return None

        candidates: list[tuple[float, float]] = []
        for rpm, map_val, ve_corr, afr in zip(rpm_values, map_values, ve_corrected_values, afr_values):
            if afr is None:
                continue
            if abs(float(ve_corr)) < 1e-9:
                continue

            rpm_dist = abs(float(rpm) - target_rpm) / 500.0
            map_dist = abs(float(map_val) - target_map) / 10.0
            distance = (rpm_dist * rpm_dist) + (map_dist * map_dist)
            afr_predicted = float(afr) * (float(ve_corr) / target_ve)
            candidates.append((distance, afr_predicted))

        if not candidates:
            return None

        candidates.sort(key=lambda item: item[0])
        selected = candidates[: max(1, neighbor_count)]
        predicted_values = [value for _, value in selected]
        weights = [1.0 / (distance + 0.05) for distance, _ in selected]
        return MainWindow._weighted_average(predicted_values, weights)

    @staticmethod
    @staticmethod
    def _default_data_dir() -> Path:
        cwd_data = Path.cwd() / "tuning_data"
        return cwd_data if cwd_data.exists() else Path.cwd()
        return cwd_data if cwd_data.exists() else Path.cwd()
