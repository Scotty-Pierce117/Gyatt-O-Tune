from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pyqtgraph as pg
from PySide6.QtCore import QEvent, Qt, QTimer
from PySide6.QtGui import QAction, QColor, QGuiApplication, QKeySequence
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDockWidget,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStatusBar,
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
        self.legend: Any = None
        self._table_type = "generic"
        self._y_label_text = "Value"
        self._selected_point_text = "<span style='color:#f0f0f0'>Selected point: none</span>"
        self._stats_tooltip_text = ""
        
        # Visibility toggles for different data series
        self._show_table = True
        self._show_raw = True
        self._show_corrected = True
        self._show_average = True
        self._show_afr = True
        self._show_afr_target = True
        self._show_afr_error = True
        self._show_knock = True

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
        self.raw_scatter = pg.ScatterPlotItem(pen=pg.mkPen(95, 140, 205, 170, width=1), brush=pg.mkBrush(95, 140, 205, 115), size=6, name="Raw VE1")
        self.corrected_scatter = pg.ScatterPlotItem(pen=pg.mkPen(95, 175, 115, 170, width=1), brush=pg.mkBrush(95, 175, 115, 115), size=6, name="EGO Corrected VE")
        self.average_curve = pg.PlotCurveItem(pen=pg.mkPen(235, 170, 85, width=3), name="Predicted VE")
        self.average_scatter = pg.ScatterPlotItem(pen=pg.mkPen(235, 170, 85, 220, width=2), brush=pg.mkBrush(235, 170, 85, 200), size=8, name="Predicted VE Points")
        self.afr_scatter = pg.ScatterPlotItem(pen=pg.mkPen(80, 155, 200, 170, width=1), brush=pg.mkBrush(80, 155, 200, 115), size=6, name="AFR")
        self.afr_target_scatter = pg.ScatterPlotItem(pen=pg.mkPen(165, 120, 205, 170, width=1), brush=pg.mkBrush(165, 120, 205, 115), size=6, name="AFR Target")
        self.afr_error_scatter = pg.ScatterPlotItem(pen=pg.mkPen(220, 170, 100, 170, width=1), brush=pg.mkBrush(220, 170, 100, 115), size=6, name="AFR Error")
        self.knock_scatter = pg.ScatterPlotItem(pen=pg.mkPen(205, 95, 135, 170, width=1), brush=pg.mkBrush(205, 95, 135, 115), size=6, name="Knock In")
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
        
        # Add toggle controls for each data series
        toggle_layout = QHBoxLayout()
        toggle_layout.setContentsMargins(0, 4, 0, 0)
        
        self.check_table = QCheckBox("Selected Row")
        self.check_table.setChecked(True)
        self.check_table.toggled.connect(self._on_toggle_table)
        toggle_layout.addWidget(self.check_table)
        
        self.check_raw = QCheckBox("Raw VE1")
        self.check_raw.setChecked(True)
        self.check_raw.toggled.connect(self._on_toggle_raw)
        toggle_layout.addWidget(self.check_raw)
        
        self.check_corrected = QCheckBox("EGO Corrected")
        self.check_corrected.setChecked(True)
        self.check_corrected.toggled.connect(self._on_toggle_corrected)
        toggle_layout.addWidget(self.check_corrected)
        
        self.check_average = QCheckBox("Predicted VE")
        self.check_average.setChecked(True)
        self.check_average.toggled.connect(self._on_toggle_average)
        toggle_layout.addWidget(self.check_average)

        self.check_afr = QCheckBox("AFR")
        self.check_afr.setChecked(True)
        self.check_afr.toggled.connect(self._on_toggle_afr)
        toggle_layout.addWidget(self.check_afr)

        self.check_afr_target = QCheckBox("AFR Target")
        self.check_afr_target.setChecked(True)
        self.check_afr_target.toggled.connect(self._on_toggle_afr_target)
        toggle_layout.addWidget(self.check_afr_target)

        self.check_afr_error = QCheckBox("AFR Error")
        self.check_afr_error.setChecked(True)
        self.check_afr_error.toggled.connect(self._on_toggle_afr_error)
        toggle_layout.addWidget(self.check_afr_error)

        self.check_knock = QCheckBox("Knock In")
        self.check_knock.setChecked(True)
        self.check_knock.toggled.connect(self._on_toggle_knock)
        toggle_layout.addWidget(self.check_knock)

        self._check_by_series: dict[str, QCheckBox] = {
            "table": self.check_table,
            "raw": self.check_raw,
            "corrected": self.check_corrected,
            "average": self.check_average,
            "afr": self.check_afr,
            "afr_target": self.check_afr_target,
            "afr_error": self.check_afr_error,
            "knock": self.check_knock,
        }
        
        toggle_layout.addStretch()
        layout.addLayout(toggle_layout)

        self._mouse_proxy = pg.SignalProxy(
            self.plot.scene().sigMouseMoved,
            rateLimit=30,
            slot=self._on_mouse_moved,
        )
        self.plot.scene().sigMouseClicked.connect(self._on_mouse_clicked)
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

    def _on_toggle_table(self, checked: bool) -> None:
        self._show_table = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_raw(self, checked: bool) -> None:
        self._show_raw = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_corrected(self, checked: bool) -> None:
        self._show_corrected = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_average(self, checked: bool) -> None:
        self._show_average = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_afr(self, checked: bool) -> None:
        self._show_afr = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_afr_target(self, checked: bool) -> None:
        self._show_afr_target = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_afr_error(self, checked: bool) -> None:
        self._show_afr_error = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _on_toggle_knock(self, checked: bool) -> None:
        self._show_knock = checked
        self._refresh_visibility()
        self._emit_visibility_changed()

    def _emit_visibility_changed(self) -> None:
        if callable(self.on_visibility_changed):
            self.on_visibility_changed(self.current_series_visibility())

    def _refresh_visibility(self) -> None:
        """Update plot item visibility based on checkboxes."""
        self.table_curve.show() if self._show_table else self.table_curve.hide()
        self.table_scatter.show() if self._show_table else self.table_scatter.hide()
        self.raw_scatter.show() if self._show_raw else self.raw_scatter.hide()
        self.corrected_scatter.show() if self._show_corrected else self.corrected_scatter.hide()
        self.average_curve.show() if self._show_average else self.average_curve.hide()
        self.average_scatter.show() if self._show_average else self.average_scatter.hide()
        self.afr_scatter.show() if self._show_afr else self.afr_scatter.hide()
        self.afr_target_scatter.show() if self._show_afr_target else self.afr_target_scatter.hide()
        self.afr_error_scatter.show() if self._show_afr_error else self.afr_error_scatter.hide()
        self.knock_scatter.show() if self._show_knock else self.knock_scatter.hide()
        self._refresh_legend()

    def current_series_visibility(self) -> dict[str, bool]:
        return {
            "table": self._show_table,
            "raw": self._show_raw,
            "corrected": self._show_corrected,
            "average": self._show_average,
            "afr": self._show_afr,
            "afr_target": self._show_afr_target,
            "afr_error": self._show_afr_error,
            "knock": self._show_knock,
        }

    def configure_series_controls(
        self,
        available_series: list[str],
        preferred_visibility: dict[str, bool] | None = None,
    ) -> None:
        preferred_visibility = preferred_visibility or {}
        available = set(available_series)
        for series_id, check in self._check_by_series.items():
            is_available = series_id in available
            check.setVisible(is_available)
            if not is_available:
                continue
            check.blockSignals(True)
            check.setChecked(bool(preferred_visibility.get(series_id, check.isChecked())))
            check.blockSignals(False)

        self._show_table = self.check_table.isChecked()
        self._show_raw = self.check_raw.isChecked()
        self._show_corrected = self.check_corrected.isChecked()
        self._show_average = self.check_average.isChecked()
        self._show_afr = self.check_afr.isChecked()
        self._show_afr_target = self.check_afr_target.isChecked()
        self._show_afr_error = self.check_afr_error.isChecked()
        self._show_knock = self.check_knock.isChecked()
        self._refresh_visibility()
        self._emit_visibility_changed()

    def clear_visualization(self, title: str = "No data selected", stats_text: str = "") -> None:
        self._selected_point = None
        self._point_sets = []
        self._table_type = "generic"
        self._y_label_text = "Value"
        self._remove_legend()
        self.plot.clear()
        self._set_graph_title(title, stats_text)
        self.plot.setLabel('left', self._y_label_text)
        self.plot.setLabel('bottom', 'RPM')
        self._set_selected_point_text("<span style='color:#f0f0f0'>Selected point: none</span>")
        self._add_plot_items()
        self._apply_view_all()

    def set_row_data(self, payload: dict[str, Any], auto_view_all: bool = True) -> None:
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
        self.plot.addItem(self.raw_scatter)
        self.plot.addItem(self.corrected_scatter)
        self.plot.addItem(self.average_curve)
        self.plot.addItem(self.average_scatter)
        self.plot.addItem(self.afr_scatter)
        self.plot.addItem(self.afr_target_scatter)
        self.plot.addItem(self.afr_error_scatter)
        self.plot.addItem(self.knock_scatter)
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
        self.raw_scatter.setZValue(10)
        self.corrected_scatter.setZValue(10)
        self.afr_scatter.setZValue(10)
        self.afr_target_scatter.setZValue(10)
        self.afr_error_scatter.setZValue(10)
        self.knock_scatter.setZValue(10)
        self.average_curve.setZValue(30)
        self.average_scatter.setZValue(31)
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
        entries = [
            ("table", self.table_curve, "Selected Row", self._show_table),
            ("raw", self.raw_scatter, "Raw VE1", self._show_raw),
            ("corrected", self.corrected_scatter, "EGO Corrected VE", self._show_corrected),
            ("average", self.average_curve, "Predicted VE", self._show_average),
            ("afr", self.afr_scatter, "AFR", self._show_afr),
            ("afr_target", self.afr_target_scatter, "AFR Target", self._show_afr_target),
            ("afr_error", self.afr_error_scatter, "AFR Error", self._show_afr_error),
            ("knock", self.knock_scatter, "Knock In", self._show_knock),
        ]

        for series_id, item, label, is_enabled in entries:
            if is_enabled and self._series_has_points(series_id):
                self.legend.addItem(item, label)

    def _get_series(self, series_id: str) -> dict[str, Any] | None:
        for series in self._point_sets:
            if series.get("series_id") == series_id:
                return series
        return None

    def _refresh_point_styles(self) -> None:
        table_series = self._get_series("table")
        raw_series = self._get_series("raw")
        corrected_series = self._get_series("corrected")
        average_series = self._get_series("average")
        afr_series = self._get_series("afr")
        afr_target_series = self._get_series("afr_target")
        afr_error_series = self._get_series("afr_error")
        knock_series = self._get_series("knock")

        if table_series is not None:
            self.table_curve.setData(table_series["rpm"], table_series["ve"])
            self.table_scatter.setData(x=table_series["rpm"], y=table_series["ve"])
        else:
            self.table_curve.setData([], [])
            self.table_scatter.setData([], [])

        if raw_series is not None:
            self.raw_scatter.setData(x=raw_series["rpm"], y=raw_series["ve"])
        else:
            self.raw_scatter.setData([], [])

        if corrected_series is not None:
            self.corrected_scatter.setData(x=corrected_series["rpm"], y=corrected_series["ve"])
        else:
            self.corrected_scatter.setData([], [])

        if average_series is not None:
            self.average_curve.setData(average_series["rpm"], average_series["ve"])
            self.average_scatter.setData(x=average_series["rpm"], y=average_series["ve"])
        else:
            self.average_curve.setData([], [])
            self.average_scatter.setData([], [])

        if afr_series is not None:
            self.afr_scatter.setData(x=afr_series["rpm"], y=afr_series["ve"])
        else:
            self.afr_scatter.setData([], [])

        if afr_target_series is not None:
            self.afr_target_scatter.setData(x=afr_target_series["rpm"], y=afr_target_series["ve"])
        else:
            self.afr_target_scatter.setData([], [])

        if afr_error_series is not None:
            self.afr_error_scatter.setData(x=afr_error_series["rpm"], y=afr_error_series["ve"])
        else:
            self.afr_error_scatter.setData([], [])

        if knock_series is not None:
            self.knock_scatter.setData(x=knock_series["rpm"], y=knock_series["ve"])
        else:
            self.knock_scatter.setData([], [])

        self._refresh_selected_marker()
        self._refresh_legend()

    def _refresh_selected_marker(self) -> None:
        if self._selected_point is None:
            self.selected_marker.setData([], [])
            self.selected_marker.hide()
            return

        series_id, index = self._selected_point
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
        def value_at(key: str) -> float | None:
            values = series.get(key, [])
            if isinstance(values, list) and 0 <= index < len(values):
                return values[index]
            return None

        name = str(series.get("name", "Point"))
        rpm = float(series["rpm"][index])
        map_val = float(series["map"][index])
        ve = float(series["ve"][index])
        ve_raw = value_at("ve_raw")
        ve_scaled = value_at("ve_scaled")
        afr = value_at("afr")
        afr_predicted = value_at("afr_predicted")
        afr_target = value_at("afr_target")
        afr_error = value_at("afr_error")
        time_elapsed = value_at("time_elapsed")
        cranking_status = value_at("cranking_status")
        warmup_enrichment = value_at("warmup_enrichment")
        startup_enrichment = value_at("startup_enrichment")
        ignition_advance = value_at("ignition_advance")

        lines = [f"Selected point ({name})"]
        detail_rows: list[tuple[str, str]] = [
            ("RPM:", f"{rpm:.1f}"),
            ("MAP:", f"{map_val:.1f} kPa"),
        ]
        series_id = str(series.get("series_id"))
        if series_id == "corrected":
            if self._table_type == "ve" and ve_scaled is not None and ve_raw is not None and abs(float(ve_raw)) > 1e-9:
                ve_offset_pct = ((float(ve_scaled) - float(ve_raw)) / float(ve_raw)) * 100.0
                detail_rows.append(("Offset vs Raw VE:", f"{ve_offset_pct:+.2f}%"))
            detail_rows.append(("VE scaled:", self._format_value(ve_scaled, '%')))
            detail_rows.append(("VE unscaled:", self._format_value(ve_raw, '%')))
        elif series_id == "table":
            detail_rows.append(("Selected row value:", f"{ve:.2f}"))
            if afr_predicted is not None:
                detail_rows.append(("AFR predicted:", self._format_value(afr_predicted)))
        elif series_id in {"afr", "afr_target", "afr_error", "knock"}:
            detail_rows.append(("Value:", f"{ve:.2f}"))
        else:
            detail_rows.append(("VE:", f"{ve:.2f}%"))
        detail_rows.append(("AFR Actual:", self._format_value(afr)))
        detail_rows.append(("AFR target:", self._format_value(afr_target)))
        detail_rows.append(("AFR error:", self._format_value(afr_error)))
        if series_id == "knock":
            detail_rows.append(("Time elapsed:", self._format_value(time_elapsed, ' s')))
            if cranking_status is None:
                detail_rows.append(("Cranking status:", "n/a"))
            else:
                detail_rows.append(("Cranking status:", "On" if float(cranking_status) >= 0.5 else "Off"))
            detail_rows.append(("Warm-up enrichment:", self._format_value(warmup_enrichment, '%')))
            detail_rows.append(("Start-up enrichment:", self._format_value(startup_enrichment, '%')))
            detail_rows.append(("Ignition advance:", self._format_value(ignition_advance, ' deg')))

        rows_html = "".join(
            f"<tr><td style='padding-right:10px'>{label}</td><td>{value}</td></tr>"
            for label, value in detail_rows
        )
        header = lines[0] if lines else ""
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
        "average": "Predicted VE",
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


class MainWindow(QMainWindow):
    MAX_RECENT_FILES = 10
    WINDOW_GEOMETRY_KEY = "main_window_geometry"
    WINDOW_DOCK_STATE_KEY = "main_window_dock_state_v2"
    DEFAULT_WINDOW_GEOMETRY_KEY = "default_main_window_geometry_v1"
    DEFAULT_WINDOW_DOCK_STATE_KEY = "default_main_window_dock_state_v1"
    OPEN_TUNE_PLACEHOLDER_KEY = "__open_tune_file__"
    ROW_VIZ_SERIES_BY_TABLE_TYPE: dict[str, list[str]] = {
        "ve": ["table", "raw", "corrected", "average"],
        "afr": ["table", "afr", "afr_target", "afr_error"],
        "knock": ["table", "knock"],
        "generic": ["table"],
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
        self.save_tune_action: Any | None = None
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
        self._history_is_applying = False
        self.undo_action: QAction | None = None
        self.redo_action: QAction | None = None
        self.average_line_data: dict[str, Any] | None = None
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
        self.show_tunerstudio_names = True
        self.row_viz_preferences: dict[str, dict[str, bool]] = self._default_row_viz_preferences()

        self._load_recent_files()
        self._load_favorites()
        self._load_table_filter_preferences()
        self._load_row_viz_preferences()
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

        open_tune_action = file_menu.addAction("Open &Tune File...")
        open_tune_action.triggered.connect(self._open_tune_file)

        select_working_folder_action = file_menu.addAction("Select &Working Folder...")
        select_working_folder_action.triggered.connect(self._on_select_working_folder)

        self.recent_tunes_menu = file_menu.addMenu("Recent &Tune Files")
        self._update_recent_tunes_menu()

        load_recent_tune_shortcut_action = QAction(self)
        load_recent_tune_shortcut_action.setShortcut(QKeySequence("Ctrl+T"))
        load_recent_tune_shortcut_action.triggered.connect(self._on_load_most_recent_tune_shortcut)
        self.addAction(load_recent_tune_shortcut_action)

        load_matching_log_shortcut_action = QAction(self)
        load_matching_log_shortcut_action.setShortcut(QKeySequence("Ctrl+L"))
        load_matching_log_shortcut_action.triggered.connect(self._on_load_matching_log_shortcut)
        self.addAction(load_matching_log_shortcut_action)

        self.save_tune_action = file_menu.addAction("&Save")
        self.save_tune_action.setShortcut(QKeySequence("Ctrl+S"))
        self.save_tune_action.triggered.connect(self._save_tune_revision)
        self.save_tune_action.setEnabled(False)

        self.save_tune_as_action = file_menu.addAction("Save Tune &As...")
        self.save_tune_as_action.triggered.connect(self._save_tune_as)
        self.save_tune_as_action.setEnabled(False)

        open_log_action = file_menu.addAction("Open &Log File...")
        open_log_action.triggered.connect(self._open_log_file)

        self.recent_logs_menu = file_menu.addMenu("Recent &Log Files")
        self._update_recent_logs_menu()

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

        edit_menu.addSeparator()
        preferences_action = edit_menu.addAction("&Preferences...")
        preferences_action.triggered.connect(self._open_preferences)

        self.view_menu = self.menuBar().addMenu("&View")
        self.tunerstudio_names_action = self.view_menu.addAction("&TunerStudio Table Names")
        self.tunerstudio_names_action.setCheckable(True)
        self.tunerstudio_names_action.setChecked(True)
        self.tunerstudio_names_action.triggered.connect(self._on_tunerstudio_names_toggled)

        self.only_show_favorited_tables_action = self.view_menu.addAction("&Favorite Tables Only")
        self.only_show_favorited_tables_action.setCheckable(True)
        self.only_show_favorited_tables_action.setChecked(self.only_show_favorited_tables)
        self.only_show_favorited_tables_action.triggered.connect(self._on_only_show_favorited_toggled)

        self.show_2d_tables_action = self.view_menu.addAction("Show &2D Tables")
        self.show_2d_tables_action.setCheckable(True)
        self.show_2d_tables_action.setChecked(self.show_2d_tables)
        self.show_2d_tables_action.triggered.connect(self._on_show_2d_tables_toggled)

        self.show_1d_tables_action = self.view_menu.addAction("Show &1D Tables")
        self.show_1d_tables_action.setCheckable(True)
        self.show_1d_tables_action.setChecked(self.show_1d_tables)
        self.show_1d_tables_action.triggered.connect(self._on_show_1d_tables_toggled)
        # Grey out dimensional filters when showing only favorited tables
        self.show_2d_tables_action.setEnabled(not self.only_show_favorited_tables)
        self.show_1d_tables_action.setEnabled(not self.only_show_favorited_tables)

    def _set_tune_save_actions_enabled(self, enabled: bool) -> None:
        if self.save_tune_action is not None:
            self.save_tune_action.setEnabled(enabled)
        if self.save_tune_as_action is not None:
            self.save_tune_as_action.setEnabled(enabled)

        self.view_menu.addSeparator()
        save_default_layout_action = self.view_menu.addAction("Save Current Layout As Default")
        save_default_layout_action.triggered.connect(self._save_default_window_layout)

        load_default_layout_action = self.view_menu.addAction("Load Default Window Layout")
        load_default_layout_action.triggered.connect(self._load_default_window_layout)

        reset_layout_action = self.view_menu.addAction("Reset Window Layout")
        reset_layout_action.triggered.connect(self._reset_window_layout)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._save_dock_layout()
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
        show_apply_predictions: bool,
        show_write: bool,
        show_revert: bool,
    ) -> None:
        self.generate_average_button.setVisible(show_generate)
        self.apply_average_to_row_button.setVisible(show_apply_predictions)
        _ = show_write
        self.revert_row_changes_button.setVisible(show_revert)

    def _refresh_history_action_state(self) -> None:
        if self.undo_action is not None:
            self.undo_action.setEnabled(bool(self.global_undo_stack))
        if self.redo_action is not None:
            self.redo_action.setEnabled(bool(self.global_redo_stack))

    def _clear_global_history(self) -> None:
        self.global_undo_stack.clear()
        self.global_redo_stack.clear()
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

        self.global_undo_stack.append(
            {
                "table": table_name,
                "row": int(row_index),
                "col": int(column_index),
                "old": float(old_value),
                "new": float(new_value),
            }
        )
        if len(self.global_undo_stack) > 5000:
            self.global_undo_stack = self.global_undo_stack[-5000:]
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
            {"rows": {}, "one_d": None, "average_line_data": None},
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
        if not self.global_undo_stack:
            return
        entry = self.global_undo_stack.pop()
        applied = self._apply_history_entry(entry, use_new_value=False)
        if applied:
            self.global_redo_stack.append(entry)
        self._refresh_history_action_state()

    def _trigger_global_redo(self) -> None:
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
        self.only_show_favorited_tables = bool(settings.value("only_show_favorited_tables", False, type=bool))
        self.show_1d_tables = bool(settings.value("show_1d_tables", True, type=bool))
        self.show_2d_tables = bool(settings.value("show_2d_tables", True, type=bool))

    def _save_table_filter_preferences(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("only_show_favorited_tables", self.only_show_favorited_tables)
        settings.setValue("show_1d_tables", self.show_1d_tables)
        settings.setValue("show_2d_tables", self.show_2d_tables)

    def _default_row_viz_preferences(self) -> dict[str, dict[str, bool]]:
        return {
            "ve": {
                "table": True,
                "raw": True,
                "corrected": True,
                "average": True,
            },
            "afr": {
                "table": True,
                "afr": True,
                "afr_target": True,
                "afr_error": True,
            },
            "knock": {
                "table": True,
                "knock": True,
            },
            "generic": {
                "table": True,
            },
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
            if series_id == "table" or series_id in existing:
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

        if self.log_df is not None and not self.log_df.empty and "knock" in table_name.lower():
            rpm_channel = None
            knock_channel = None
            afr_channel = None
            afr_target_channel = None
            afr_error_channel = None
            time_elapsed_channel = None
            cranking_channel = None
            warmup_enrichment_channel = None
            startup_enrichment_channel = None
            ignition_advance_channel = None
            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if rpm_channel is None and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if knock_channel is None and ('knock in' in col_lower or 'knock_in' in col_lower or 'knockin' in col_lower):
                    knock_channel = str(col)
                if (
                    not afr_channel
                    and ('afr' in col_lower or 'lambda' in col_lower)
                    and 'target' not in col_lower
                    and 'tgt' not in col_lower
                    and 'error' not in col_lower
                    and 'err' not in col_lower
                ):
                    afr_channel = str(col)
                if not afr_target_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('target' in col_lower or 'tgt' in col_lower):
                    afr_target_channel = str(col)
                if not afr_error_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('error' in col_lower or 'err' in col_lower):
                    afr_error_channel = str(col)
                if not time_elapsed_channel and (
                    ('time' in col_lower and ('sec' in col_lower or 'elapsed' in col_lower))
                    or col_lower in {'time', 'seconds', 'sec'}
                ):
                    time_elapsed_channel = str(col)
                if not cranking_channel and 'crank' in col_lower:
                    cranking_channel = str(col)
                if not warmup_enrichment_channel and (
                    'warmup' in col_lower
                    or 'warm up' in col_lower
                    or 'wue' in col_lower
                ):
                    warmup_enrichment_channel = str(col)
                if not startup_enrichment_channel and (
                    'startup' in col_lower
                    or 'start up' in col_lower
                    or 'afterstart' in col_lower
                    or 'ase' in col_lower
                ):
                    startup_enrichment_channel = str(col)
                if (
                    not ignition_advance_channel
                    and ('advance' in col_lower or ('ign' in col_lower and 'adv' in col_lower) or ('spark' in col_lower and 'adv' in col_lower))
                    and 'table' not in col_lower
                ):
                    ignition_advance_channel = str(col)
            if knock_channel is None:
                for col in self.log_df.columns:
                    col_lower = str(col).lower()
                    if 'knock' in col_lower and 'threshold' not in col_lower:
                        knock_channel = str(col)
                        break

            if rpm_channel and knock_channel:
                rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
                knock_series = self._to_numeric_series(self.log_df[knock_channel])
                afr_series = self._to_numeric_series(self.log_df[afr_channel]) if afr_channel else None
                afr_target_series = self._to_numeric_series(self.log_df[afr_target_channel]) if afr_target_channel else None
                afr_error_series = self._to_numeric_series(self.log_df[afr_error_channel]) if afr_error_channel else None
                time_elapsed_series = self._to_numeric_series(self.log_df[time_elapsed_channel]) if time_elapsed_channel else None
                cranking_series = self._to_numeric_series(self.log_df[cranking_channel]) if cranking_channel else None
                warmup_enrichment_series = self._to_numeric_series(self.log_df[warmup_enrichment_channel]) if warmup_enrichment_channel else None
                startup_enrichment_series = self._to_numeric_series(self.log_df[startup_enrichment_channel]) if startup_enrichment_channel else None
                ignition_advance_series = self._to_numeric_series(self.log_df[ignition_advance_channel]) if ignition_advance_channel else None
                if rpm_series is not None and knock_series is not None:
                    import numpy as np

                    rpm_vals = np.asarray(rpm_series, dtype=float)
                    knock_vals = np.asarray(knock_series, dtype=float)
                    finite = np.isfinite(rpm_vals) & np.isfinite(knock_vals)
                    rpm_vals = rpm_vals[finite]
                    knock_vals = knock_vals[finite]

                    def to_optional_values(series: Any) -> list[float | None]:
                        if series is None:
                            return [None] * len(rpm_vals)
                        values = np.asarray(series, dtype=float)[finite]
                        return [float(v) if np.isfinite(v) else None for v in values]

                    afr_vals = to_optional_values(afr_series)
                    afr_target_vals = to_optional_values(afr_target_series)
                    afr_error_vals = to_optional_values(afr_error_series)
                    if afr_error_series is None and afr_series is not None and afr_target_series is not None:
                        afr_error_vals = [
                            (float(a) - float(t)) if (a is not None and t is not None) else None
                            for a, t in zip(afr_vals, afr_target_vals)
                        ]
                    time_elapsed_vals = to_optional_values(time_elapsed_series)
                    cranking_vals = to_optional_values(cranking_series)
                    warmup_vals = to_optional_values(warmup_enrichment_series)
                    startup_vals = to_optional_values(startup_enrichment_series)
                    ignition_advance_vals = to_optional_values(ignition_advance_series)

                    if len(rpm_vals) > 0:
                        point_sets.append(
                            {
                                "series_id": "knock",
                                "name": "Knock In",
                                "rpm": [float(v) for v in rpm_vals],
                                "ve": [float(v) for v in knock_vals],
                                "map": [0.0] * len(rpm_vals),
                                "ve_raw": [float(v) for v in knock_vals],
                                "ve_scaled": [float(v) for v in knock_vals],
                                "afr": afr_vals,
                                "afr_target": afr_target_vals,
                                "afr_error": afr_error_vals,
                                "time_elapsed": time_elapsed_vals,
                                "cranking_status": cranking_vals,
                                "warmup_enrichment": warmup_vals,
                                "startup_enrichment": startup_vals,
                                "ignition_advance": ignition_advance_vals,
                            }
                        )
                        stats += f"\nKnock points: {len(rpm_vals)}"

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

    def _open_recent_tune_file(self, file_path: Path) -> None:
        if not file_path.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {file_path} no longer exists.")
            self.recent_tune_files.remove(file_path)
            self._save_recent_files()
            self._update_recent_tunes_menu()
            return

        try:
            self.tune_data = self.tune_loader.load(file_path)
            self.selected_rows_per_table = {}  # Clear saved row selections for new tune
            self.pending_edits_per_table = {}  # Clear saved pending edits for new tune
            self.loaded_table_snapshots = self._snapshot_table_values(self.tune_data)
            self._clear_global_history()
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

    def _open_recent_log_file(self, file_path: Path) -> None:
        if not file_path.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {file_path} no longer exists.")
            self.recent_log_files.remove(file_path)
            self._save_recent_files()
            self._update_recent_logs_menu()
            return

        # Show wait cursor while loading
        self.setCursor(Qt.CursorShape.WaitCursor)
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            self.setCursor(Qt.CursorShape.ArrowCursor)
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return

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

        self.show_tunerstudio_names = self.tunerstudio_names_action.isChecked()
        current_item = self.table_list.currentItem()
        current_table_name = current_item.data(Qt.ItemDataRole.UserRole) if current_item else None

        self.table_list.clear()
        all_items = []
        for table_name in sorted(self.tune_data.tables.keys()):
            # Apply filters
            if self.only_show_favorited_tables:
                # Only show favorited tables (ignore 1D filter)
                if table_name not in self.favorite_tables:
                    continue
            else:
                # Apply dimensional filters when not in favorites-only mode
                is_1d = self._is_1d_table(table_name)
                if is_1d and not self.show_1d_tables:
                    continue
                if (not is_1d) and not self.show_2d_tables:
                    continue
            
            display_name = self._get_tunerstudio_name(table_name) if self.show_tunerstudio_names else table_name
            if table_name in self.favorite_tables and not self.only_show_favorited_tables:
                display_name = f"★ {display_name}"
            all_items.append((table_name, display_name))

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
            elif not self.show_1d_tables and not self.show_2d_tables:
                self.table_list.addItem("All table types are hidden")
            elif not self.show_1d_tables:
                self.table_list.addItem("No 2D tables found")
            elif not self.show_2d_tables:
                self.table_list.addItem("No 1D tables found")
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

    def _on_tunerstudio_names_toggled(self) -> None:
        """Handle View menu TunerStudio names toggle."""
        self._update_table_display()

    def _on_only_show_favorited_toggled(self) -> None:
        """Handle 'Favorite Tables Only' toggle."""
        self.only_show_favorited_tables = self.only_show_favorited_tables_action.isChecked()
        # Grey out table-type filters when showing only favorited tables
        self.show_2d_tables_action.setEnabled(not self.only_show_favorited_tables)
        self.show_1d_tables_action.setEnabled(not self.only_show_favorited_tables)
        self._save_table_filter_preferences()
        self._update_table_display()

    def _on_show_2d_tables_toggled(self) -> None:
        """Handle 'Show 2D Tables' toggle."""
        self.show_2d_tables = self.show_2d_tables_action.isChecked()
        self._save_table_filter_preferences()
        self._update_table_display()

    def _on_show_1d_tables_toggled(self) -> None:
        """Handle 'Show 1D Tables' toggle."""
        self.show_1d_tables = self.show_1d_tables_action.isChecked()
        self._save_table_filter_preferences()
        self._update_table_display()

    def _open_preferences(self) -> None:
        dialog = RowVisualizationPreferencesDialog(
            table_type_series=self.ROW_VIZ_SERIES_BY_TABLE_TYPE,
            current_preferences=self.row_viz_preferences,
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        self.row_viz_preferences = dialog.preferences()
        self._save_row_viz_preferences()
        self.statusBar().showMessage("Saved row visualization preferences", 3000)
        self._update_table_grid_row_visualization()

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

        menu.exec(list_widget.mapToGlobal(position))

    def _toggle_favorite(self, table_name: str) -> None:
        """Toggle favorite status for a table."""
        if table_name in self.favorite_tables:
            self.favorite_tables.remove(table_name)
        else:
            self.favorite_tables.add(table_name)
        self._save_favorites()
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

        # Left panel - Tune Tables
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
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
                margin: 4px 6px;
                padding: 8px 10px;
            }
            QListWidget::item:hover {
                background-color: palette(alternate-base);
                border: 1px solid palette(midlight);
            }
            QListWidget::item:selected {
                background-color: palette(highlight);
                color: palette(highlighted-text);
                border: 1px solid palette(light);
            }
            """
        )
        self._add_open_tune_file_placeholder()
        self.table_list.currentItemChanged.connect(self._on_table_selected)
        self.table_list.itemClicked.connect(self._on_table_list_item_clicked)
        self.table_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_list.customContextMenuRequested.connect(self._show_table_context_menu)
        left_layout.addWidget(self.table_list)

        # Selected table panel
        selected_table_panel = QWidget()
        selected_table_layout = QVBoxLayout(selected_table_panel)
        selected_table_layout.setContentsMargins(0, 0, 0, 0)

        self.table_meta = QLabel("No table selected")
        selected_table_layout.addWidget(self.table_meta)

        self.axis_meta = QLabel("Axes: X (index), Y (index)")
        selected_table_layout.addWidget(self.axis_meta)

        controls_layout = QHBoxLayout()
        self.transpose_checkbox = QCheckBox("Transpose Table")
        self.transpose_checkbox.toggled.connect(self._refresh_current_table_view)
        controls_layout.addWidget(self.transpose_checkbox)

        self.swap_axes_checkbox = QCheckBox("Swap Axes Labels")
        self.swap_axes_checkbox.toggled.connect(self._refresh_current_table_view)
        controls_layout.addWidget(self.swap_axes_checkbox)

        controls_layout.addStretch(1)
        selected_table_layout.addLayout(controls_layout)

        self.table_grid = CopyPasteTableWidget()
        self.table_grid.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table_grid.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
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
        self.table_row_panel.on_point_adjust_requested = self._on_table_row_plot_point_adjust_requested
        self.table_row_panel.on_visibility_changed = self._on_row_viz_visibility_changed

        self.table_row_editor_group = QGroupBox("")
        row_editor_layout = QVBoxLayout(self.table_row_editor_group)

        row_editor_controls = QHBoxLayout()
        self.table_row_label = QLabel("Selected row: none")
        row_editor_controls.addWidget(self.table_row_label)
        self.generate_average_button = QPushButton("Predict VE Value")
        self.generate_average_button.setToolTip(
            "Predict VE values from nearby logged VE points using EGO-corrected VE and matching AFR readings."
        )
        self.generate_average_button.clicked.connect(self._generate_average_line_from_log)
        row_editor_controls.addWidget(self.generate_average_button)
        self.apply_average_to_row_button = QPushButton("Apply Predictions")
        self.apply_average_to_row_button.clicked.connect(self._apply_average_to_selected_row)
        row_editor_controls.addWidget(self.apply_average_to_row_button)
        self.revert_row_changes_button = QPushButton("Revert Row")
        self.revert_row_changes_button.clicked.connect(self._revert_pending_row_values)
        row_editor_controls.addWidget(self.revert_row_changes_button)
        row_editor_controls.addStretch(1)
        row_editor_layout.addLayout(row_editor_controls)

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
            show_apply_predictions=False,
            show_write=False,
            show_revert=False,
        )

        self.table_row_status = QLabel("Select a VE table and click a row to view and edit it.")
        row_editor_container_layout.addWidget(self.table_row_status)

        # Wrap row visualization panel to avoid it inheriting dock margins directly.
        row_viz_container = QWidget()
        row_viz_layout = QVBoxLayout(row_viz_container)
        row_viz_layout.setContentsMargins(0, 0, 0, 0)
        row_viz_layout.addWidget(self.table_row_panel)

        self.tune_tables_dock = self._create_section_dock(
            "Tune Tables",
            "tuneTablesDock",
            self._create_outlined_panel(left),
        )
        self.selected_table_dock = self._create_section_dock(
            "Selected Table",
            "selectedTableDock",
            self._create_outlined_panel(selected_table_panel),
        )
        self.row_viz_dock = self._create_section_dock("Row Data Visualization", "rowVizDock", row_viz_container)
        self.row_editor_dock = self._create_section_dock(
            "Selected Row Editor",
            "rowEditorDock",
            self._create_outlined_panel(row_editor_container),
        )

        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.tune_tables_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.selected_table_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_viz_dock)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.row_editor_dock)

        self._apply_standard_window_layout()

        self.view_menu.addSeparator()
        self.view_menu.addAction(self.tune_tables_dock.toggleViewAction())
        self.view_menu.addAction(self.selected_table_dock.toggleViewAction())
        self.view_menu.addAction(self.row_viz_dock.toggleViewAction())
        self.view_menu.addAction(self.row_editor_dock.toggleViewAction())

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
        table_ve_values: list[float],
    ) -> dict[str, Any]:
        rpm_point_values = [float(v) for v in rpm_values]
        table_point_values = [float(v) for v in table_ve_values]
        point_sets: list[dict[str, Any]] = [
            {
                "series_id": "table",
                "name": "Table VE",
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

        title = f"MAP {map_value} kPa - Table VE"
        stats = f"Table VE values for MAP {map_value} kPa\n"
        stats += f"RPM range: {min(rpm_values):.0f} - {max(rpm_values):.0f}\n"
        stats += f"VE range: {min(table_ve_values):.1f}% - {max(table_ve_values):.1f}%"

        if self.log_df is None or self.log_df.empty:
            return {"title": title, "stats": stats + "\n\nNo log data loaded.", "point_sets": point_sets}

        try:
            rpm_channel = None
            map_channel = None
            ve1_channel = None
            ego_cor1_channel = None
            afr_channel = None
            afr_target_channel = None
            afr_error_channel = None

            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if not map_channel and ('map' in col_lower):
                    map_channel = str(col)
                if not ve1_channel and ('ve1' in col_lower or 've 1' in col_lower):
                    ve1_channel = str(col)
                if not ego_cor1_channel and ('ego' in col_lower and 'cor1' in col_lower):
                    ego_cor1_channel = str(col)
                if not afr_target_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('target' in col_lower or 'tgt' in col_lower):
                    afr_target_channel = str(col)
                if not afr_error_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('error' in col_lower or 'err' in col_lower):
                    afr_error_channel = str(col)
                if (
                    not afr_channel
                    and ('afr' in col_lower or 'lambda' in col_lower)
                    and 'target' not in col_lower
                    and 'tgt' not in col_lower
                    and 'error' not in col_lower
                    and 'err' not in col_lower
                ):
                    afr_channel = str(col)

            if not rpm_channel or not map_channel or not ve1_channel:
                return {
                    "title": f"MAP {map_value} kPa - Required channels not found",
                    "stats": stats + "\n\nRequired channels not found in log data.\nNeed: RPM, MAP, VE1",
                    "point_sets": point_sets,
                }

            rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
            map_series = self._to_numeric_series(self.log_df[map_channel])
            ve_series = self._to_numeric_series(self.log_df[ve1_channel])
            afr_series = self._to_numeric_series(self.log_df[afr_channel]) if afr_channel else None
            afr_target_series = self._to_numeric_series(self.log_df[afr_target_channel]) if afr_target_channel else None
            afr_error_series = self._to_numeric_series(self.log_df[afr_error_channel]) if afr_error_channel else None

            if rpm_series is None or map_series is None or ve_series is None:
                return {
                    "title": f"MAP {map_value} kPa - Could not convert data",
                    "stats": stats + "\n\nCould not convert log data to numeric values.",
                    "point_sets": point_sets,
                }

            ve_corrected_series = ve_series.copy()
            if ego_cor1_channel is not None:
                ego_series = self._to_numeric_series(self.log_df[ego_cor1_channel])
                if ego_series is not None:
                    ve_corrected_series = ve_series * (ego_series / 100.0)

            map_tolerance = 10.0
            if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
                sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
                map_tolerance = max(10.0, min_spacing * 0.75)
            map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
            filtered_rpm = rpm_series[map_mask]
            filtered_map = map_series[map_mask]
            filtered_ve_raw = ve_series[map_mask]
            filtered_ve_corrected = ve_corrected_series[map_mask]
            filtered_afr = afr_series[map_mask] if afr_series is not None else None
            filtered_afr_target = afr_target_series[map_mask] if afr_target_series is not None else None
            filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None

            if len(filtered_rpm) == 0 and map_tolerance < 20.0:
                map_tolerance = 20.0
                map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                filtered_rpm = rpm_series[map_mask]
                filtered_map = map_series[map_mask]
                filtered_ve_raw = ve_series[map_mask]
                filtered_ve_corrected = ve_corrected_series[map_mask]
                filtered_afr = afr_series[map_mask] if afr_series is not None else None
                filtered_afr_target = afr_target_series[map_mask] if afr_target_series is not None else None
                filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None

            if filtered_afr_error is None and filtered_afr is not None and filtered_afr_target is not None:
                filtered_afr_error = filtered_afr - filtered_afr_target

            if len(filtered_rpm) == 0:
                return {
                    "title": f"MAP {map_value} kPa - No data near MAP value",
                    "stats": stats + f"\n\nNo log data found near MAP {map_value} kPa (±{map_tolerance} kPa).",
                    "point_sets": point_sets,
                }

            filtered_rpm_values = [float(v) for v in filtered_rpm]
            filtered_map_values = [float(v) for v in filtered_map]
            filtered_ve_raw_values = [float(v) for v in filtered_ve_raw]
            filtered_ve_corrected_values = [float(v) for v in filtered_ve_corrected]
            filtered_afr_values = [float(v) for v in filtered_afr] if filtered_afr is not None else [None] * len(filtered_rpm_values)
            filtered_afr_target_values = [float(v) for v in filtered_afr_target] if filtered_afr_target is not None else [None] * len(filtered_rpm_values)
            filtered_afr_error_values = [float(v) for v in filtered_afr_error] if filtered_afr_error is not None else [None] * len(filtered_rpm_values)

            table_afr_values: list[float | None] = []
            table_afr_target_values: list[float | None] = []
            table_afr_error_values: list[float | None] = []
            table_afr_predicted_values: list[float | None] = []
            for rpm_value in rpm_point_values:
                nearest_index = min(range(len(filtered_rpm_values)), key=lambda idx: abs(filtered_rpm_values[idx] - rpm_value))
                table_afr_values.append(filtered_afr_values[nearest_index])
                table_afr_target_values.append(filtered_afr_target_values[nearest_index])
                table_afr_error_values.append(filtered_afr_error_values[nearest_index])

            for rpm_value, table_ve_value in zip(rpm_point_values, table_point_values):
                predicted = self._predict_afr_from_neighbors(
                    target_rpm=float(rpm_value),
                    target_map=float(map_value),
                    target_ve=float(table_ve_value),
                    rpm_values=filtered_rpm_values,
                    map_values=filtered_map_values,
                    ve_corrected_values=filtered_ve_corrected_values,
                    afr_values=filtered_afr_values,
                )
                table_afr_predicted_values.append(predicted)

            point_sets[0]["afr"] = table_afr_values
            point_sets[0]["afr_predicted"] = table_afr_predicted_values
            point_sets[0]["afr_target"] = table_afr_target_values
            point_sets[0]["afr_error"] = table_afr_error_values

            point_sets.append(
                {
                    "series_id": "raw",
                    "name": "Raw VE1",
                    "rpm": filtered_rpm_values,
                    "ve": filtered_ve_raw_values,
                    "map": filtered_map_values,
                    "ve_raw": filtered_ve_raw_values,
                    "ve_scaled": filtered_ve_raw_values,
                    "afr": filtered_afr_values,
                    "afr_predicted": [None] * len(filtered_rpm_values),
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                }
            )
            point_sets.append(
                {
                    "series_id": "corrected",
                    "name": "EGO Corrected VE",
                    "rpm": filtered_rpm_values,
                    "ve": filtered_ve_corrected_values,
                    "map": filtered_map_values,
                    "ve_raw": filtered_ve_raw_values,
                    "ve_scaled": filtered_ve_corrected_values,
                    "afr": filtered_afr_values,
                    "afr_predicted": [None] * len(filtered_rpm_values),
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                }
            )

            raw_mean = filtered_ve_raw.mean()
            raw_std = filtered_ve_raw.std()
            corrected_mean = filtered_ve_corrected.mean()
            corrected_std = filtered_ve_corrected.std()

            stats += f"\n\nLogged Data ({len(filtered_rpm)} points):\n"
            stats += f"Raw VE1: {raw_mean:.1f}% (σ={raw_std:.1f})\n"
            stats += f"EGO Corrected: {corrected_mean:.1f}% (σ={corrected_std:.1f})\n"
            predicted_only = [v for v in table_afr_predicted_values if v is not None]
            target_only = [v for v in table_afr_target_values if v is not None]
            if predicted_only:
                stats += f"Predicted AFR (row avg): {sum(predicted_only) / len(predicted_only):.2f}\n"
            if target_only:
                stats += f"AFR target (row avg): {sum(target_only) / len(target_only):.2f}\n"
            stats += f"RPM range: {filtered_rpm.min():.0f} - {filtered_rpm.max():.0f}"

            return {
                "title": f"MAP {map_value} kPa - {len(filtered_rpm)} points",
                "stats": stats,
                "point_sets": point_sets,
            }

        except Exception as exc:
            return {
                "title": f"MAP {map_value} kPa - Error: {exc}",
                "stats": stats + f"\n\nError plotting logged data: {exc}",
                "point_sets": point_sets,
            }

    def _build_afr_row_visualization_payload(
        self,
        map_value: float,
        rpm_values: list[float],
        table_afr_target_values: list[float],
    ) -> dict[str, Any]:
        rpm_point_values = [float(v) for v in rpm_values]
        table_point_values = [float(v) for v in table_afr_target_values]
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
                "afr_target": table_point_values,
                "afr_error": [None] * len(rpm_point_values),
            }
        ]

        title = f"MAP {map_value} kPa - AFR Target Row"
        stats = f"AFR target row for MAP {map_value} kPa\n"
        stats += f"RPM range: {min(rpm_values):.0f} - {max(rpm_values):.0f}\n"
        stats += f"AFR target range: {min(table_afr_target_values):.2f} - {max(table_afr_target_values):.2f}"

        if self.log_df is None or self.log_df.empty:
            return {
                "title": title,
                "stats": stats + "\n\nNo log data loaded.",
                "point_sets": point_sets,
                "y_label": "AFR / Error",
            }

        try:
            rpm_channel = None
            map_channel = None
            afr_channel = None
            afr_target_channel = None
            afr_error_channel = None

            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if not map_channel and ('map' in col_lower):
                    map_channel = str(col)
                if not afr_target_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('target' in col_lower or 'tgt' in col_lower):
                    afr_target_channel = str(col)
                if not afr_error_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('error' in col_lower or 'err' in col_lower):
                    afr_error_channel = str(col)
                if (
                    not afr_channel
                    and ('afr' in col_lower or 'lambda' in col_lower)
                    and 'target' not in col_lower
                    and 'tgt' not in col_lower
                    and 'error' not in col_lower
                    and 'err' not in col_lower
                ):
                    afr_channel = str(col)

            if not rpm_channel or not map_channel or not afr_target_channel:
                return {
                    "title": f"MAP {map_value} kPa - Required channels not found",
                    "stats": stats + "\n\nRequired channels not found in log data.\nNeed: RPM, MAP, AFR Target",
                    "point_sets": point_sets,
                    "y_label": "AFR / Error",
                }

            rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
            map_series = self._to_numeric_series(self.log_df[map_channel])
            afr_series = self._to_numeric_series(self.log_df[afr_channel]) if afr_channel else None
            afr_target_series = self._to_numeric_series(self.log_df[afr_target_channel])
            afr_error_series = self._to_numeric_series(self.log_df[afr_error_channel]) if afr_error_channel else None

            if rpm_series is None or map_series is None or afr_target_series is None:
                return {
                    "title": f"MAP {map_value} kPa - Could not convert data",
                    "stats": stats + "\n\nCould not convert log data to numeric values.",
                    "point_sets": point_sets,
                    "y_label": "AFR / Error",
                }

            if afr_error_series is None and afr_series is not None:
                afr_error_series = afr_series - afr_target_series

            map_tolerance = 10.0
            if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
                sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
                map_tolerance = max(10.0, min_spacing * 0.75)

            map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
            filtered_rpm = rpm_series[map_mask]
            filtered_map = map_series[map_mask]
            filtered_afr = afr_series[map_mask] if afr_series is not None else None
            filtered_afr_target = afr_target_series[map_mask]
            filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None

            if len(filtered_rpm) == 0 and map_tolerance < 20.0:
                map_tolerance = 20.0
                map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                filtered_rpm = rpm_series[map_mask]
                filtered_map = map_series[map_mask]
                filtered_afr = afr_series[map_mask] if afr_series is not None else None
                filtered_afr_target = afr_target_series[map_mask]
                filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None

            if len(filtered_rpm) == 0:
                return {
                    "title": f"MAP {map_value} kPa - No data near MAP value",
                    "stats": stats + f"\n\nNo log data found near MAP {map_value} kPa (±{map_tolerance} kPa).",
                    "point_sets": point_sets,
                    "y_label": "AFR / Error",
                }

            filtered_rpm_values = [float(v) for v in filtered_rpm]
            filtered_map_values = [float(v) for v in filtered_map]
            filtered_afr_values = [float(v) for v in filtered_afr] if filtered_afr is not None else [None] * len(filtered_rpm_values)
            filtered_afr_target_values = [float(v) for v in filtered_afr_target]
            filtered_afr_error_values = [float(v) for v in filtered_afr_error] if filtered_afr_error is not None else [None] * len(filtered_rpm_values)

            point_sets[0]["afr"] = filtered_afr_values[:len(point_sets[0]["afr"])]
            point_sets[0]["afr_error"] = filtered_afr_error_values[:len(point_sets[0]["afr_error"])]

            point_sets.append(
                {
                    "series_id": "afr",
                    "name": "AFR",
                    "rpm": filtered_rpm_values,
                    "ve": [float(v) if v is not None else 0.0 for v in filtered_afr_values],
                    "map": filtered_map_values,
                    "ve_raw": filtered_afr_values,
                    "ve_scaled": filtered_afr_values,
                    "afr": filtered_afr_values,
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                }
            )
            point_sets.append(
                {
                    "series_id": "afr_target",
                    "name": "AFR Target",
                    "rpm": filtered_rpm_values,
                    "ve": filtered_afr_target_values,
                    "map": filtered_map_values,
                    "ve_raw": filtered_afr_target_values,
                    "ve_scaled": filtered_afr_target_values,
                    "afr": filtered_afr_values,
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                }
            )
            point_sets.append(
                {
                    "series_id": "afr_error",
                    "name": "AFR Error",
                    "rpm": filtered_rpm_values,
                    "ve": [float(v) if v is not None else 0.0 for v in filtered_afr_error_values],
                    "map": filtered_map_values,
                    "ve_raw": filtered_afr_error_values,
                    "ve_scaled": filtered_afr_error_values,
                    "afr": filtered_afr_values,
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                }
            )

            afr_values_only = [v for v in filtered_afr_values if v is not None]
            afr_error_only = [v for v in filtered_afr_error_values if v is not None]

            stats += f"\n\nLogged Data ({len(filtered_rpm_values)} points):\n"
            if afr_values_only:
                stats += f"AFR mean: {sum(afr_values_only) / len(afr_values_only):.2f}\n"
            stats += f"AFR target mean: {sum(filtered_afr_target_values) / len(filtered_afr_target_values):.2f}\n"
            if afr_error_only:
                stats += f"AFR error mean: {sum(afr_error_only) / len(afr_error_only):.2f}\n"
            stats += f"RPM range: {min(filtered_rpm_values):.0f} - {max(filtered_rpm_values):.0f}"

            return {
                "title": f"MAP {map_value} kPa - {len(filtered_rpm_values)} points",
                "stats": stats,
                "point_sets": point_sets,
                "y_label": "AFR / Error",
            }
        except Exception as exc:
            return {
                "title": f"MAP {map_value} kPa - Error: {exc}",
                "stats": stats + f"\n\nError plotting logged data: {exc}",
                "point_sets": point_sets,
                "y_label": "AFR / Error",
            }

    def _build_knock_row_visualization_payload(
        self,
        map_value: float,
        rpm_values: list[float],
        table_knock_threshold_values: list[float],
    ) -> dict[str, Any]:
        rpm_point_values = [float(v) for v in rpm_values]
        table_point_values = [float(v) for v in table_knock_threshold_values]
        point_sets: list[dict[str, Any]] = [
            {
                "series_id": "table",
                "name": "Knock Threshold (Selected Row)",
                "rpm": rpm_point_values,
                "ve": table_point_values,
                "map": [float(map_value)] * len(rpm_point_values),
                "ve_raw": table_point_values,
                "ve_scaled": table_point_values,
                "afr": [None] * len(rpm_point_values),
                "afr_target": [None] * len(rpm_point_values),
                "afr_error": [None] * len(rpm_point_values),
                "time_elapsed": [None] * len(rpm_point_values),
                "cranking_status": [None] * len(rpm_point_values),
                "warmup_enrichment": [None] * len(rpm_point_values),
                "startup_enrichment": [None] * len(rpm_point_values),
                "ignition_advance": [None] * len(rpm_point_values),
            }
        ]

        title = f"MAP {map_value} kPa - Knock Threshold Row"
        stats = f"Knock threshold row for MAP {map_value} kPa\n"
        stats += f"RPM range: {min(rpm_values):.0f} - {max(rpm_values):.0f}\n"
        stats += f"Threshold range: {min(table_knock_threshold_values):.2f} - {max(table_knock_threshold_values):.2f}"

        if self.log_df is None or self.log_df.empty:
            return {
                "title": title,
                "stats": stats + "\n\nNo log data loaded.",
                "point_sets": point_sets,
                "y_label": "Knock",
            }

        try:
            rpm_channel = None
            map_channel = None
            knock_channel = None
            afr_channel = None
            afr_target_channel = None
            afr_error_channel = None
            time_elapsed_channel = None
            cranking_channel = None
            warmup_enrichment_channel = None
            startup_enrichment_channel = None
            ignition_advance_channel = None

            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if not map_channel and ('map' in col_lower):
                    map_channel = str(col)
                if knock_channel is None and ('knock in' in col_lower or 'knock_in' in col_lower or 'knockin' in col_lower):
                    knock_channel = str(col)
                if (
                    not afr_channel
                    and ('afr' in col_lower or 'lambda' in col_lower)
                    and 'target' not in col_lower
                    and 'tgt' not in col_lower
                    and 'error' not in col_lower
                    and 'err' not in col_lower
                ):
                    afr_channel = str(col)
                if not afr_target_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('target' in col_lower or 'tgt' in col_lower):
                    afr_target_channel = str(col)
                if not afr_error_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('error' in col_lower or 'err' in col_lower):
                    afr_error_channel = str(col)
                if not time_elapsed_channel and (
                    ('time' in col_lower and ('sec' in col_lower or 'elapsed' in col_lower))
                    or col_lower in {'time', 'seconds', 'sec'}
                ):
                    time_elapsed_channel = str(col)
                if not cranking_channel and 'crank' in col_lower:
                    cranking_channel = str(col)
                if not warmup_enrichment_channel and (
                    'warmup' in col_lower
                    or 'warm up' in col_lower
                    or 'wue' in col_lower
                ):
                    warmup_enrichment_channel = str(col)
                if not startup_enrichment_channel and (
                    'startup' in col_lower
                    or 'start up' in col_lower
                    or 'afterstart' in col_lower
                    or 'ase' in col_lower
                ):
                    startup_enrichment_channel = str(col)
                if (
                    not ignition_advance_channel
                    and ('advance' in col_lower or ('ign' in col_lower and 'adv' in col_lower) or ('spark' in col_lower and 'adv' in col_lower))
                    and 'table' not in col_lower
                ):
                    ignition_advance_channel = str(col)

            if knock_channel is None:
                for col in self.log_df.columns:
                    col_lower = str(col).lower()
                    if 'knock' in col_lower and 'threshold' not in col_lower:
                        knock_channel = str(col)
                        break

            if not rpm_channel or not map_channel or not knock_channel:
                return {
                    "title": f"MAP {map_value} kPa - Required channels not found",
                    "stats": stats + "\n\nRequired channels not found in log data.\nNeed: RPM, MAP, Knock In",
                    "point_sets": point_sets,
                    "y_label": "Knock",
                }

            rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
            map_series = self._to_numeric_series(self.log_df[map_channel])
            knock_series = self._to_numeric_series(self.log_df[knock_channel])
            afr_series = self._to_numeric_series(self.log_df[afr_channel]) if afr_channel else None
            afr_target_series = self._to_numeric_series(self.log_df[afr_target_channel]) if afr_target_channel else None
            afr_error_series = self._to_numeric_series(self.log_df[afr_error_channel]) if afr_error_channel else None
            time_elapsed_series = self._to_numeric_series(self.log_df[time_elapsed_channel]) if time_elapsed_channel else None
            cranking_series = self._to_numeric_series(self.log_df[cranking_channel]) if cranking_channel else None
            warmup_enrichment_series = self._to_numeric_series(self.log_df[warmup_enrichment_channel]) if warmup_enrichment_channel else None
            startup_enrichment_series = self._to_numeric_series(self.log_df[startup_enrichment_channel]) if startup_enrichment_channel else None
            ignition_advance_series = self._to_numeric_series(self.log_df[ignition_advance_channel]) if ignition_advance_channel else None

            if rpm_series is None or map_series is None or knock_series is None:
                return {
                    "title": f"MAP {map_value} kPa - Could not convert data",
                    "stats": stats + "\n\nCould not convert log data to numeric values.",
                    "point_sets": point_sets,
                    "y_label": "Knock",
                }

            map_tolerance = 10.0
            if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
                sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
                map_tolerance = max(10.0, min_spacing * 0.75)

            map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
            filtered_rpm = rpm_series[map_mask]
            filtered_map = map_series[map_mask]
            filtered_knock = knock_series[map_mask]
            filtered_afr = afr_series[map_mask] if afr_series is not None else None
            filtered_afr_target = afr_target_series[map_mask] if afr_target_series is not None else None
            filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None
            filtered_time_elapsed = time_elapsed_series[map_mask] if time_elapsed_series is not None else None
            filtered_cranking = cranking_series[map_mask] if cranking_series is not None else None
            filtered_warmup_enrichment = warmup_enrichment_series[map_mask] if warmup_enrichment_series is not None else None
            filtered_startup_enrichment = startup_enrichment_series[map_mask] if startup_enrichment_series is not None else None
            filtered_ignition_advance = ignition_advance_series[map_mask] if ignition_advance_series is not None else None

            if len(filtered_rpm) == 0 and map_tolerance < 20.0:
                map_tolerance = 20.0
                map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                filtered_rpm = rpm_series[map_mask]
                filtered_map = map_series[map_mask]
                filtered_knock = knock_series[map_mask]
                filtered_afr = afr_series[map_mask] if afr_series is not None else None
                filtered_afr_target = afr_target_series[map_mask] if afr_target_series is not None else None
                filtered_afr_error = afr_error_series[map_mask] if afr_error_series is not None else None
                filtered_time_elapsed = time_elapsed_series[map_mask] if time_elapsed_series is not None else None
                filtered_cranking = cranking_series[map_mask] if cranking_series is not None else None
                filtered_warmup_enrichment = warmup_enrichment_series[map_mask] if warmup_enrichment_series is not None else None
                filtered_startup_enrichment = startup_enrichment_series[map_mask] if startup_enrichment_series is not None else None
                filtered_ignition_advance = ignition_advance_series[map_mask] if ignition_advance_series is not None else None

            if len(filtered_rpm) == 0:
                return {
                    "title": f"MAP {map_value} kPa - No data near MAP value",
                    "stats": stats + f"\n\nNo log data found near MAP {map_value} kPa (±{map_tolerance} kPa).",
                    "point_sets": point_sets,
                    "y_label": "Knock",
                }

            filtered_rpm_values = [float(v) for v in filtered_rpm]
            filtered_map_values = [float(v) for v in filtered_map]
            filtered_knock_values = [float(v) for v in filtered_knock]
            filtered_afr_values = [float(v) for v in filtered_afr] if filtered_afr is not None else [None] * len(filtered_rpm_values)
            filtered_afr_target_values = [float(v) for v in filtered_afr_target] if filtered_afr_target is not None else [None] * len(filtered_rpm_values)
            filtered_afr_error_values = [float(v) for v in filtered_afr_error] if filtered_afr_error is not None else [None] * len(filtered_rpm_values)
            if filtered_afr_error is None and filtered_afr is not None and filtered_afr_target is not None:
                filtered_afr_error_values = [
                    (float(a) - float(t)) if (a is not None and t is not None) else None
                    for a, t in zip(filtered_afr_values, filtered_afr_target_values)
                ]
            filtered_time_elapsed_values = [float(v) for v in filtered_time_elapsed] if filtered_time_elapsed is not None else [None] * len(filtered_rpm_values)
            filtered_cranking_values = [float(v) for v in filtered_cranking] if filtered_cranking is not None else [None] * len(filtered_rpm_values)
            filtered_warmup_enrichment_values = [float(v) for v in filtered_warmup_enrichment] if filtered_warmup_enrichment is not None else [None] * len(filtered_rpm_values)
            filtered_startup_enrichment_values = [float(v) for v in filtered_startup_enrichment] if filtered_startup_enrichment is not None else [None] * len(filtered_rpm_values)
            filtered_ignition_advance_values = [float(v) for v in filtered_ignition_advance] if filtered_ignition_advance is not None else [None] * len(filtered_rpm_values)

            table_afr_values: list[float | None] = []
            table_afr_target_values: list[float | None] = []
            table_afr_error_values: list[float | None] = []
            table_time_elapsed_values: list[float | None] = []
            table_cranking_values: list[float | None] = []
            table_warmup_values: list[float | None] = []
            table_startup_values: list[float | None] = []
            table_ignition_advance_values: list[float | None] = []
            for rpm_value in rpm_point_values:
                nearest_index = min(range(len(filtered_rpm_values)), key=lambda idx: abs(filtered_rpm_values[idx] - rpm_value))
                table_afr_values.append(filtered_afr_values[nearest_index])
                table_afr_target_values.append(filtered_afr_target_values[nearest_index])
                table_afr_error_values.append(filtered_afr_error_values[nearest_index])
                table_time_elapsed_values.append(filtered_time_elapsed_values[nearest_index])
                table_cranking_values.append(filtered_cranking_values[nearest_index])
                table_warmup_values.append(filtered_warmup_enrichment_values[nearest_index])
                table_startup_values.append(filtered_startup_enrichment_values[nearest_index])
                table_ignition_advance_values.append(filtered_ignition_advance_values[nearest_index])

            point_sets[0]["afr"] = table_afr_values
            point_sets[0]["afr_target"] = table_afr_target_values
            point_sets[0]["afr_error"] = table_afr_error_values
            point_sets[0]["time_elapsed"] = table_time_elapsed_values
            point_sets[0]["cranking_status"] = table_cranking_values
            point_sets[0]["warmup_enrichment"] = table_warmup_values
            point_sets[0]["startup_enrichment"] = table_startup_values
            point_sets[0]["ignition_advance"] = table_ignition_advance_values

            point_sets.append(
                {
                    "series_id": "knock",
                    "name": "Knock In",
                    "rpm": filtered_rpm_values,
                    "ve": filtered_knock_values,
                    "map": filtered_map_values,
                    "ve_raw": filtered_knock_values,
                    "ve_scaled": filtered_knock_values,
                    "afr": filtered_afr_values,
                    "afr_target": filtered_afr_target_values,
                    "afr_error": filtered_afr_error_values,
                    "time_elapsed": filtered_time_elapsed_values,
                    "cranking_status": filtered_cranking_values,
                    "warmup_enrichment": filtered_warmup_enrichment_values,
                    "startup_enrichment": filtered_startup_enrichment_values,
                    "ignition_advance": filtered_ignition_advance_values,
                }
            )

            knock_mean = sum(filtered_knock_values) / len(filtered_knock_values)
            stats += f"\n\nLogged Data ({len(filtered_rpm_values)} points):\n"
            stats += f"Knock mean: {knock_mean:.3f}\n"
            stats += f"RPM range: {min(filtered_rpm_values):.0f} - {max(filtered_rpm_values):.0f}"

            return {
                "title": f"MAP {map_value} kPa - {len(filtered_rpm_values)} points",
                "stats": stats,
                "point_sets": point_sets,
                "y_label": "Knock",
            }
        except Exception as exc:
            return {
                "title": f"MAP {map_value} kPa - Error: {exc}",
                "stats": stats + f"\n\nError plotting logged data: {exc}",
                "point_sets": point_sets,
                "y_label": "Knock",
            }

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

    def _generate_average_line_from_log(self) -> None:
        """Predict VE values from AFR and EGO-corrected VE log data for the current MAP row."""
        if self.current_table is None or "ve" not in self.current_table.name.lower():
            self.table_row_status.setText("VE prediction is currently available for VE rows.")
            return

        if self.log_df is None or self.log_df.empty:
            self.table_row_status.setText("No log data loaded. Cannot predict VE.")
            return

        if self.selected_table_row_idx is None or self.current_y_axis is None or self.current_x_axis is None:
            self.table_row_status.setText("Select a row first.")
            return

        map_value = float(self.current_y_axis.values[self.selected_table_row_idx])

        # Find relevant columns in log data
        rpm_channel = None
        map_channel = None
        ve1_channel = None
        ego_cor1_channel = None
        afr_channel = None
        afr_target_channel = None

        for col in self.log_df.columns:
            col_lower = str(col).lower()
            if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                rpm_channel = str(col)
            if not map_channel and ('map' in col_lower):
                map_channel = str(col)
            if not ve1_channel and ('ve1' in col_lower or 've 1' in col_lower):
                ve1_channel = str(col)
            if not ego_cor1_channel and ('ego' in col_lower and 'cor1' in col_lower):
                ego_cor1_channel = str(col)
            if not afr_target_channel and ('afr' in col_lower or 'lambda' in col_lower) and ('target' in col_lower or 'tgt' in col_lower):
                afr_target_channel = str(col)
            if (
                not afr_channel
                and ('afr' in col_lower or 'lambda' in col_lower)
                and 'target' not in col_lower
                and 'tgt' not in col_lower
                and 'error' not in col_lower
                and 'err' not in col_lower
            ):
                afr_channel = str(col)

        if not rpm_channel or not map_channel or not ve1_channel:
            self.table_row_status.setText("Required log columns not found (RPM, MAP, VE1).")
            return

        if not afr_channel:
            self.table_row_status.setText("AFR channel not found in log. Cannot predict VE without AFR readings.")
            return

        # Convert to numeric series
        rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
        map_series = self._to_numeric_series(self.log_df[map_channel])
        ve_series = self._to_numeric_series(self.log_df[ve1_channel])
        afr_series = self._to_numeric_series(self.log_df[afr_channel])
        afr_target_series = self._to_numeric_series(self.log_df[afr_target_channel]) if afr_target_channel else None

        if rpm_series is None or map_series is None or ve_series is None or afr_series is None:
            self.table_row_status.setText("Could not convert log data to numeric values.")
            return

        # Apply EGO correction to VE
        ve_corrected_series = ve_series.copy()
        if ego_cor1_channel is not None:
            ego_series = self._to_numeric_series(self.log_df[ego_cor1_channel])
            if ego_series is not None:
                ve_corrected_series = ve_series * (ego_series / 100.0)

        # Calculate adaptive MAP tolerance
        map_tolerance = 10.0
        if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
            sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
            min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
            map_tolerance = max(10.0, min_spacing * 0.75)

        # Filter to current MAP bin
        map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
        filtered_rpm = rpm_series[map_mask]
        filtered_map = map_series[map_mask]
        filtered_ve_corrected = ve_corrected_series[map_mask]
        filtered_afr = afr_series[map_mask]
        filtered_afr_target = afr_target_series[map_mask] if afr_target_series is not None else None

        if len(filtered_rpm) == 0:
            self.table_row_status.setText(f"No log data near MAP {map_value} kPa.")
            return

        filtered_rpm_values = [float(v) for v in filtered_rpm]
        filtered_map_values = [float(v) for v in filtered_map]
        filtered_ve_corrected_values = [float(v) for v in filtered_ve_corrected]
        filtered_afr_values: list[float | None] = [float(v) for v in filtered_afr]
        filtered_afr_target_values: list[float | None] = (
            [float(v) for v in filtered_afr_target] if filtered_afr_target is not None else [None] * len(filtered_rpm_values)
        )

        # Predict VE for each RPM bin using AFR-based weighted neighbor averaging
        rpm_bins = [float(v) for v in self.current_x_axis.values]
        average_rpm_values: list[float] = []
        average_ve_values: list[float] = []
        average_afr_target_values: list[float | None] = []

        for target_rpm in rpm_bins:
            # Determine the target AFR for this bin from nearby log points
            rpm_tolerance = 500.0
            if len(rpm_bins) > 1:
                sorted_rpms = sorted(rpm_bins)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_rpms, sorted_rpms[1:]))
                rpm_tolerance = max(300.0, min_spacing * 0.4)

            rpm_mask_bin = [abs(r - target_rpm) <= rpm_tolerance for r in filtered_rpm_values]
            nearby_afr_target = [v for v, m in zip(filtered_afr_target_values, rpm_mask_bin) if m and v is not None]
            if afr_target_channel is not None and len(nearby_afr_target) < 2:
                continue
            target_afr = float(sum(nearby_afr_target) / len(nearby_afr_target)) if nearby_afr_target else 14.7

            predicted_ve = self._predict_ve_from_neighbors(
                target_rpm=target_rpm,
                target_map=map_value,
                target_afr=target_afr,
                rpm_values=filtered_rpm_values,
                map_values=filtered_map_values,
                ve_corrected_values=filtered_ve_corrected_values,
                afr_values=filtered_afr_values,
            )

            if predicted_ve is not None:
                average_rpm_values.append(target_rpm)
                average_ve_values.append(self._round_ve_prediction(predicted_ve))
                average_afr_target_values.append(target_afr)

        if not average_rpm_values:
            self.table_row_status.setText("Could not predict VE values from log data for this MAP.")
            return

        skipped_points = len(rpm_bins) - len(average_rpm_values)

        # Store the prediction data for display
        self.average_line_data = {
            "series_id": "average",
            "name": "Predicted VE",
            "rpm": average_rpm_values,
            "ve": average_ve_values,
            "map": [map_value] * len(average_rpm_values),
            "ve_raw": average_ve_values,
            "ve_scaled": average_ve_values,
            "afr": [None] * len(average_rpm_values),
            "afr_target": average_afr_target_values,
            "afr_error": [None] * len(average_rpm_values),
        }

        self.table_row_status.setText(
            f"Predicted VE from {len(filtered_rpm)} log points near MAP {map_value} kPa "
            f"using AFR{'+ target' if afr_target_channel else ' (14.7 default target)'}; "
            f"predicted {len(average_rpm_values)}/{len(rpm_bins)} row points"
            f"{f', skipped {skipped_points} for low data' if skipped_points > 0 else ''}."
        )
        self._set_row_editor_button_visibility(
            show_generate=True,
            show_apply_predictions=True,
            show_write=bool(self.pending_row_values),
            show_revert=bool(self.pending_row_values),
        )
        self._update_table_grid_row_visualization()


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
        """Show context menu on a row editor cell when prediction data is available for that RPM bin."""
        if self.average_line_data is None or self.current_x_axis is None:
            return

        item = self.table_row_editor.itemAt(pos)
        if item is None:
            return

        col = self.table_row_editor.column(item)
        if col < 0 or col >= len(self.current_x_axis.values):
            return

        target_rpm = float(self.current_x_axis.values[col])
        avg_rpm_values = [float(v) for v in self.average_line_data.get("rpm", [])]
        avg_ve_values = [float(v) for v in self.average_line_data.get("ve", [])]
        avg_by_rpm = dict(zip(avg_rpm_values, avg_ve_values))

        if target_rpm not in avg_by_rpm:
            return

        avg_value = avg_by_rpm[target_rpm]
        menu = QMenu(self)
        apply_action = menu.addAction(f"Apply Prediction ({avg_value:.4g}) to This Cell")
        result = menu.exec(self.table_row_editor.viewport().mapToGlobal(pos))
        if result is apply_action:
            old_value = float(self.pending_row_values[col])
            self.pending_row_values[col] = avg_value
            table_name = self._current_table_name()
            cell_coords = self._current_editor_cell_coordinates(col)
            if table_name is not None and cell_coords is not None:
                row_index, source_col = cell_coords
                self._record_global_cell_edit(table_name, row_index, source_col, old_value, float(avg_value))
            self._refresh_table_row_editor()
            self._update_table_grid_row_visualization(include_row_plot=False)
            self._schedule_table_grid_row_visualization_update()

    def _apply_average_to_selected_row(self) -> None:
        if (
            self.average_line_data is None
            or self.current_x_axis is None
            or not self.pending_row_values
            or self.current_table is None
            or "ve" not in self.current_table.name.lower()
        ):
            self.table_row_status.setText("Generate a predicted VE line first.")
            return

        avg_rpm_values = [float(v) for v in self.average_line_data.get("rpm", [])]
        avg_ve_values = [float(v) for v in self.average_line_data.get("ve", [])]
        if not avg_rpm_values or not avg_ve_values or len(avg_rpm_values) != len(avg_ve_values):
            self.table_row_status.setText("Prediction data is not available to apply.")
            return

        # Build a lookup of RPM bins that actually have averaged data
        avg_by_rpm = dict(zip(avg_rpm_values, avg_ve_values))

        new_pending = list(self.pending_row_values)
        table_name = self._current_table_name()
        for i, target_rpm in enumerate(self.current_x_axis.values):
            rpm_key = float(target_rpm)
            if rpm_key in avg_by_rpm:
                old_value = float(new_pending[i])
                new_value = float(avg_by_rpm[rpm_key])
                new_pending[i] = new_value
                cell_coords = self._current_editor_cell_coordinates(i)
                if table_name is not None and cell_coords is not None:
                    row_index, source_col = cell_coords
                    self._record_global_cell_edit(table_name, row_index, source_col, old_value, new_value)
        self.pending_row_values = new_pending
        self._refresh_table_row_editor()
        self._update_table_grid_row_visualization(include_row_plot=False)
        self._schedule_table_grid_row_visualization_update()
        applied_count = len(avg_rpm_values)
        self.table_row_status.setText(f"Applied predictions to {applied_count} of {len(self.current_x_axis.values)} cells with log data.")

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
                show_apply_predictions=False,
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
        if "knock" in table_name:
            payload = self._build_knock_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
        elif "afr" in table_name:
            payload = self._build_afr_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
        else:
            payload = self._build_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
            payload["y_label"] = "VE %"
        
        # Add predicted VE line if it has been generated
        if self.average_line_data is not None and "ve" in table_name:
            point_sets = payload.get("point_sets", [])
            point_sets.append(self.average_line_data)
            payload["point_sets"] = point_sets

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
        status_message = f"Editing MAP {map_value:g} kPa for {self.current_table.name}. Changes are previewed below as a modified table."
        prediction_summary = self._current_editor_prediction_summary(payload)
        if prediction_summary:
            status_message = f"{status_message} {prediction_summary}"
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

    def _current_editor_prediction_summary(self, payload: dict[str, Any]) -> str:
        if self._row_table_type() != "ve":
            return ""
        if self.table_row_editor.columnCount() <= 0:
            return ""
        current_column = self.table_row_editor.currentColumn()
        if current_column < 0:
            return ""

        table_series: dict[str, Any] | None = None
        for series in payload.get("point_sets", []):
            if isinstance(series, dict) and str(series.get("series_id")) == "table":
                table_series = series
                break
        if table_series is None:
            return ""

        afr_predicted = None
        afr_target = None
        predicted_values = table_series.get("afr_predicted", [])
        target_values = table_series.get("afr_target", [])
        if isinstance(predicted_values, list) and 0 <= current_column < len(predicted_values):
            afr_predicted = predicted_values[current_column]
        if isinstance(target_values, list) and 0 <= current_column < len(target_values):
            afr_target = target_values[current_column]

        if afr_predicted is None and afr_target is None:
            return ""

        predicted_text = "n/a" if afr_predicted is None else f"{float(afr_predicted):.2f}"
        target_text = "n/a" if afr_target is None else f"{float(afr_target):.2f}"
        return f"Selected-cell AFR predicted: {predicted_text} | AFR target: {target_text}."

    def _table_row_editing_available(self) -> tuple[bool, str]:
        if not self.current_table or not self.current_x_axis or not self.current_y_axis:
            return False, "Select a table and click a row to view and edit it."

        if self._is_current_table_1d():
            return True, ""

        table_name = self.current_table.name.lower()
        if "ve" not in table_name and "afr" not in table_name and "knock" not in table_name:
            return False, "Row visualization is currently available for VE, AFR, and knock threshold tables."
        if self.transpose_checkbox.isChecked() or self.swap_axes_checkbox.isChecked():
            return False, "Disable Transpose Table and Swap Axes Labels to edit rows from the table grid."
        return True, ""

    def _clear_table_row_editor(self, status_message: str) -> None:
        self.selected_table_row_idx = None
        self.pending_row_values = []
        self.row_default_values = []
        self.row_edit_undo_stack = []
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None
        self.average_line_data = None
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
            show_apply_predictions=False,
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

        current_column = self.table_grid.currentColumn()
        preferred_column = current_column if current_column >= 0 else None
        self._load_selected_table_row(source_row, preferred_column=preferred_column)

    def _on_table_grid_cell_clicked(self, row: int, column: int) -> None:
        if self._is_current_table_1d():
            self._update_table_grid_row_visualization()
            return
        source_row = self._display_row_to_source_row(row)
        if source_row is None:
            return
        self._load_selected_table_row(source_row, preferred_column=column)

    def _on_table_grid_current_cell_changed(
        self,
        current_row: int,
        current_column: int,
        previous_row: int,
        previous_column: int,
    ) -> None:
        _ = (previous_row, previous_column)
        if self._is_current_table_1d():
            self._update_table_grid_row_visualization()
            return
        source_row = self._display_row_to_source_row(current_row)
        if source_row is None:
            return
        self._load_selected_table_row(source_row, preferred_column=current_column)

    def _load_selected_table_row(self, source_row: int, preferred_column: int | None = None) -> None:
        if self.current_table is None or self.current_x_axis is None or self.current_y_axis is None:
            return

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
            saved = self.pending_edits_per_table.get(table_name, {})
            self.average_line_data = saved.get("average_line_data")
            self.table_row_label.setText("Selected row: full 1D table")
            self._refresh_table_row_editor(preferred_column=preferred_column)
            self._update_table_grid_row_visualization()
            return

        self.selected_table_row_idx = source_row
        baseline_values = self._row_values_for_table(self.current_table, source_row)
        staged_values = self._staged_row_values(table_name, self.current_table, source_row)
        self.pending_row_values = staged_values if staged_values is not None else baseline_values
        self.row_default_values = [float(value) for value in baseline_values]
        self.row_edit_undo_stack = []
        self._active_row_edit_column = None
        self._active_row_edit_snapshot = None
        saved = self.pending_edits_per_table.get(table_name, {})
        self.average_line_data = saved.get("average_line_data")
        map_value = float(self.current_y_axis.values[source_row])
        self.table_row_label.setText(f"Selected row: MAP {map_value:g} kPa")
        self._refresh_table_row_editor(preferred_column=preferred_column)
        self._update_table_grid_row_visualization()

    def _refresh_table_row_editor(self, preferred_column: int | None = None) -> None:
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
        if self.pending_row_values:
            if preferred_column is not None and 0 <= preferred_column < len(self.pending_row_values):
                target_column = preferred_column
            else:
                current_column = self.table_row_editor.currentColumn()
                target_column = current_column if 0 <= current_column < len(self.pending_row_values) else 0
            self.table_row_editor.setCurrentCell(0, target_column)
        self.table_row_editor.resizeColumnsToContents()
        self.table_row_editor.resizeRowsToContents()
        self.table_row_editor.setEnabled(True)
        can_predict_ve = (
            self.current_table is not None
            and "ve" in self.current_table.name.lower()
            and not self._is_current_table_1d()
        )
        self._set_row_editor_button_visibility(
            show_generate=can_predict_ve,
            show_apply_predictions=can_predict_ve and self.average_line_data is not None,
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

    def _save_tune_revision(self) -> None:
        if self.tune_data is None or self.loaded_tune_path is None:
            QMessageBox.warning(self, "No Tune Loaded", "Load a tune file before saving.")
            return

        if self.save_target_path is None:
            # First save after loading a tune behaves like Save As.
            self._save_tune_as()
            return

        self._save_tune_to_path(self.save_target_path)

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
        
        # Show wait cursor while loading
        self.setCursor(Qt.CursorShape.WaitCursor)
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            self.setCursor(Qt.CursorShape.ArrowCursor)
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return

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
            {"rows": {}, "one_d": None, "average_line_data": None},
        )
        rows_state = table_state.setdefault("rows", {})

        if self._is_current_table_1d():
            baseline_values = self._row_values_for_table(self.current_table)
            staged_values = [float(value) for value in self.pending_row_values]
            if self._values_differ(staged_values, baseline_values):
                table_state["one_d"] = staged_values
            else:
                table_state["one_d"] = None
            table_state["average_line_data"] = self.average_line_data
        else:
            if self.selected_table_row_idx is None:
                return
            baseline_values = self._row_values_for_table(self.current_table, self.selected_table_row_idx)
            staged_values = [float(value) for value in self.pending_row_values]
            if self._values_differ(staged_values, baseline_values):
                rows_state[self.selected_table_row_idx] = staged_values
            else:
                rows_state.pop(self.selected_table_row_idx, None)
            table_state["average_line_data"] = self.average_line_data

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
            saved = self.pending_edits_per_table[table_name]
            self.average_line_data = saved.get("average_line_data")
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
            and not self.transpose_checkbox.isChecked()
            and not self.swap_axes_checkbox.isChecked()
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
            and not self.transpose_checkbox.isChecked()
            and not self.swap_axes_checkbox.isChecked()
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

        if self.transpose_checkbox.isChecked():
            matrix = [list(col) for col in zip(*matrix)]
            display_x_axis, display_y_axis = display_y_axis, display_x_axis

        if self.swap_axes_checkbox.isChecked():
            display_x_axis, display_y_axis = display_y_axis, display_x_axis

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
    def _predict_ve_from_neighbors(
        target_rpm: float,
        target_map: float,
        target_afr: float,
        rpm_values: list[float],
        map_values: list[float],
        ve_corrected_values: list[float],
        afr_values: list[float | None],
        neighbor_count: int = 24,
        min_neighbors: int = 4,
        max_distance: float = 6.0,
    ) -> float | None:
        """Predict the VE needed to achieve target_afr using EGO-corrected log VE and actual AFR neighbors."""
        if abs(target_afr) < 1e-9:
            return None

        candidates: list[tuple[float, float]] = []
        for rpm, map_val, ve_corr, afr in zip(rpm_values, map_values, ve_corrected_values, afr_values):
            if afr is None or abs(float(afr)) < 1e-9 or abs(float(ve_corr)) < 1e-9:
                continue
            rpm_dist = abs(float(rpm) - target_rpm) / 500.0
            map_dist = abs(float(map_val) - target_map) / 10.0
            distance = (rpm_dist * rpm_dist) + (map_dist * map_dist)
            # Scale EGO-corrected VE by actual/target AFR ratio to predict VE needed for target AFR
            ve_predicted = float(ve_corr) * (float(afr) / target_afr)
            candidates.append((distance, ve_predicted))

        if not candidates:
            return None

        candidates.sort(key=lambda item: item[0])
        selected = candidates[: max(1, neighbor_count)]
        close_selected = [item for item in selected if item[0] <= max_distance]
        if len(close_selected) < max(1, min_neighbors):
            return None

        selected = close_selected
        predicted_values = [value for _, value in selected]
        weights = [1.0 / (distance + 0.05) for distance, _ in selected]
        return MainWindow._weighted_average(predicted_values, weights)

    @staticmethod
    def _round_ve_prediction(value: float) -> float:
        return round(float(value), 1)

    @staticmethod
    def _default_data_dir() -> Path:
        cwd_data = Path.cwd() / "tuning_data"
        return cwd_data if cwd_data.exists() else Path.cwd()
