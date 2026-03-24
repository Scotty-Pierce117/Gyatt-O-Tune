from __future__ import annotations

from pathlib import Path
from typing import Any

import pyqtgraph as pg
from PySide6.QtCore import QEvent, QObject, Qt, QTimer
from PySide6.QtGui import QColor, QGuiApplication, QKeySequence
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
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
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import QSettings

from gyatt_o_tune.core.io import AxisVector, LogLoader, TableData, TuneData, TuneLoader
from gyatt_o_tune.ui.log_viewer_window import LogViewerWindow


class CopyPasteTableWidget(QTableWidget):
    """Spreadsheet-style copy/paste with TSV payload."""

    def __init__(self) -> None:
        super().__init__()
        self.footer_rows = 0

    def set_footer_rows(self, count: int) -> None:
        self.footer_rows = max(0, count)

    def _data_row_count(self) -> int:
        return max(0, self.rowCount() - self.footer_rows)

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        if event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            self.paste_selection()
            return
        super().keyPressEvent(event)

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


class RowVisualizationPanel(QGroupBox):
    """Reusable row data visualization with crosshair and point selection."""

    def __init__(self, title: str = "Row Data Visualization") -> None:
        super().__init__(title)
        self._selected_point: tuple[str, int] | None = None
        self._point_sets: list[dict[str, Any]] = []
        self.on_point_selected: Any = None
        self._y_label_text = "Value"
        
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

        self.plot = pg.PlotWidget()
        self.plot.showGrid(x=True, y=True, alpha=0.2)
        self.plot.getPlotItem().setLabel("bottom", "RPM")
        self.plot.getPlotItem().setLabel("left", "Value")

        self.table_curve = pg.PlotCurveItem(pen=pg.mkPen('r', width=3), name="Selected Row Data")
        self.table_scatter = pg.ScatterPlotItem(size=8, name="Selected Row Points")
        self.raw_scatter = pg.ScatterPlotItem(pen=pg.mkPen('b', width=1), brush=pg.mkBrush('b'), size=6, name="Raw VE1")
        self.corrected_scatter = pg.ScatterPlotItem(pen=pg.mkPen('g', width=1), brush=pg.mkBrush('g'), size=6, name="EGO Corrected VE")
        self.average_curve = pg.PlotCurveItem(pen=pg.mkPen(255, 165, 0, width=2), name="Average from Log")
        self.afr_scatter = pg.ScatterPlotItem(pen=pg.mkPen(35, 140, 220, width=1), brush=pg.mkBrush(35, 140, 220), size=6, name="AFR")
        self.afr_target_scatter = pg.ScatterPlotItem(pen=pg.mkPen(180, 90, 220, width=1), brush=pg.mkBrush(180, 90, 220), size=6, name="AFR Target")
        self.afr_error_scatter = pg.ScatterPlotItem(pen=pg.mkPen(240, 165, 40, width=1), brush=pg.mkBrush(240, 165, 40), size=6, name="AFR Error")
        self.knock_scatter = pg.ScatterPlotItem(pen=pg.mkPen(230, 60, 110, width=1), brush=pg.mkBrush(230, 60, 110), size=6, name="Knock In")
        self.crosshair_vline = pg.InfiniteLine(angle=90, movable=False, pen=pg.mkPen(200, 200, 200, 140, width=1))
        self.crosshair_hline = pg.InfiniteLine(angle=0, movable=False, pen=pg.mkPen(200, 200, 200, 140, width=1))

        layout.addWidget(self.plot)
        
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
        
        self.check_average = QCheckBox("Average from Log")
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

        self.stats_text = QTextEdit()
        self.stats_text.setMaximumHeight(100)
        self.stats_text.setReadOnly(True)
        layout.addWidget(self.stats_text)

        self.cursor_label = QLabel("Cursor: RPM -, Value -")
        layout.addWidget(self.cursor_label)

        self.selected_point_label = QLabel("Selected point: none")
        self.selected_point_label.setWordWrap(True)
        layout.addWidget(self.selected_point_label)

        self._mouse_proxy = pg.SignalProxy(
            self.plot.scene().sigMouseMoved,
            rateLimit=60,
            slot=self._on_mouse_moved,
        )
        self.plot.scene().sigMouseClicked.connect(self._on_mouse_clicked)
        self.clear_visualization()

    def _on_toggle_table(self, checked: bool) -> None:
        self._show_table = checked
        self._refresh_visibility()

    def _on_toggle_raw(self, checked: bool) -> None:
        self._show_raw = checked
        self._refresh_visibility()

    def _on_toggle_corrected(self, checked: bool) -> None:
        self._show_corrected = checked
        self._refresh_visibility()

    def _on_toggle_average(self, checked: bool) -> None:
        self._show_average = checked
        self._refresh_visibility()

    def _on_toggle_afr(self, checked: bool) -> None:
        self._show_afr = checked
        self._refresh_visibility()

    def _on_toggle_afr_target(self, checked: bool) -> None:
        self._show_afr_target = checked
        self._refresh_visibility()

    def _on_toggle_afr_error(self, checked: bool) -> None:
        self._show_afr_error = checked
        self._refresh_visibility()

    def _on_toggle_knock(self, checked: bool) -> None:
        self._show_knock = checked
        self._refresh_visibility()

    def _refresh_visibility(self) -> None:
        """Update plot item visibility based on checkboxes."""
        self.table_curve.show() if self._show_table else self.table_curve.hide()
        self.table_scatter.show() if self._show_table else self.table_scatter.hide()
        self.raw_scatter.show() if self._show_raw else self.raw_scatter.hide()
        self.corrected_scatter.show() if self._show_corrected else self.corrected_scatter.hide()
        self.average_curve.show() if self._show_average else self.average_curve.hide()
        self.afr_scatter.show() if self._show_afr else self.afr_scatter.hide()
        self.afr_target_scatter.show() if self._show_afr_target else self.afr_target_scatter.hide()
        self.afr_error_scatter.show() if self._show_afr_error else self.afr_error_scatter.hide()
        self.knock_scatter.show() if self._show_knock else self.knock_scatter.hide()

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

    def clear_visualization(self, title: str = "No data selected", stats_text: str = "") -> None:
        self._selected_point = None
        self._point_sets = []
        self._y_label_text = "Value"
        self.plot.clear()
        self.plot.setTitle(title)
        self.plot.setLabel('left', self._y_label_text)
        self.plot.setLabel('bottom', 'RPM')
        self.stats_text.setPlainText(stats_text)
        self.cursor_label.setText("Cursor: RPM -, Value -")
        self.selected_point_label.setText("Selected point: none")
        self._add_plot_items()

    def set_row_data(self, payload: dict[str, Any]) -> None:
        self._selected_point = None
        self._point_sets = list(payload.get("point_sets", []))
        self._y_label_text = str(payload.get("y_label", "Value"))
        x_label_text = str(payload.get("x_label", "RPM"))
        available_series = [str(s) for s in payload.get("available_series", [])]
        preferred_visibility = payload.get("series_visibility", {})
        self.configure_series_controls(available_series, preferred_visibility)
        self.plot.clear()
        self.plot.setTitle(str(payload.get("title", "Row Data Visualization")))
        self.plot.setLabel('left', self._y_label_text)
        self.plot.setLabel('bottom', x_label_text)
        self.stats_text.setPlainText(str(payload.get("stats", "")))
        self.cursor_label.setText("Cursor: RPM -, Value -")
        self.selected_point_label.setText("Selected point: none")
        self._add_plot_items()
        self._refresh_point_styles()
        self._refresh_visibility()

    def _add_plot_items(self) -> None:
        self.plot.addItem(self.table_curve)
        self.plot.addItem(self.table_scatter)
        self.plot.addItem(self.raw_scatter)
        self.plot.addItem(self.corrected_scatter)
        self.plot.addItem(self.average_curve)
        self.plot.addItem(self.afr_scatter)
        self.plot.addItem(self.afr_target_scatter)
        self.plot.addItem(self.afr_error_scatter)
        self.plot.addItem(self.knock_scatter)
        self.plot.addItem(self.crosshair_vline, ignoreBounds=True)
        self.plot.addItem(self.crosshair_hline, ignoreBounds=True)
        self.crosshair_vline.hide()
        self.crosshair_hline.hide()

        legend = pg.LegendItem(offset=(50, 30))
        legend.setParentItem(self.plot.graphicsItem())
        legend.addItem(self.table_curve, "Selected Row Data")
        legend.addItem(self.raw_scatter, "Raw VE1")
        legend.addItem(self.corrected_scatter, "EGO Corrected VE")
        legend.addItem(self.average_curve, "Average from Log")
        legend.addItem(self.afr_scatter, "AFR")
        legend.addItem(self.afr_target_scatter, "AFR Target")
        legend.addItem(self.afr_error_scatter, "AFR Error")
        legend.addItem(self.knock_scatter, "Knock In")

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
            self.table_scatter.setData(spots=self._build_spots(table_series))
        else:
            self.table_curve.setData([], [])
            self.table_scatter.setData(spots=[])

        if raw_series is not None:
            self.raw_scatter.setData(spots=self._build_spots(raw_series))
        else:
            self.raw_scatter.setData(spots=[])

        if corrected_series is not None:
            self.corrected_scatter.setData(spots=self._build_spots(corrected_series))
        else:
            self.corrected_scatter.setData(spots=[])

        if average_series is not None:
            self.average_curve.setData(average_series["rpm"], average_series["ve"])
        else:
            self.average_curve.setData([], [])

        if afr_series is not None:
            self.afr_scatter.setData(spots=self._build_spots(afr_series))
        else:
            self.afr_scatter.setData(spots=[])

        if afr_target_series is not None:
            self.afr_target_scatter.setData(spots=self._build_spots(afr_target_series))
        else:
            self.afr_target_scatter.setData(spots=[])

        if afr_error_series is not None:
            self.afr_error_scatter.setData(spots=self._build_spots(afr_error_series))
        else:
            self.afr_error_scatter.setData(spots=[])

        if knock_series is not None:
            self.knock_scatter.setData(spots=self._build_spots(knock_series))
        else:
            self.knock_scatter.setData(spots=[])

    def _build_spots(self, series: dict[str, Any]) -> list[dict[str, Any]]:
        series_id = str(series.get("series_id", ""))
        rpm_values = series.get("rpm", [])
        ve_values = series.get("ve", [])
        has_selection = self._selected_point is not None
        selected_series_id, selected_index = self._selected_point if self._selected_point is not None else (None, None)

        base_brushes = {
            "table": pg.mkBrush(255, 80, 80, 220),
            "raw": pg.mkBrush(60, 120, 255, 220),
            "corrected": pg.mkBrush(60, 190, 90, 220),
            "average": pg.mkBrush(255, 165, 0, 220),
            "afr": pg.mkBrush(35, 140, 220, 220),
            "afr_target": pg.mkBrush(180, 90, 220, 220),
            "afr_error": pg.mkBrush(240, 165, 40, 220),
            "knock": pg.mkBrush(230, 60, 110, 220),
        }
        base_pens = {
            "table": pg.mkPen(180, 40, 40, 230, width=1.2),
            "raw": pg.mkPen(40, 80, 210, 230, width=1.0),
            "corrected": pg.mkPen(40, 140, 70, 230, width=1.0),
            "average": pg.mkPen(230, 140, 0, 230, width=1.2),
            "afr": pg.mkPen(35, 140, 220, 230, width=1.0),
            "afr_target": pg.mkPen(180, 90, 220, 230, width=1.0),
            "afr_error": pg.mkPen(240, 165, 40, 230, width=1.0),
            "knock": pg.mkPen(230, 60, 110, 230, width=1.0),
        }
        dim_brush = pg.mkBrush(150, 150, 150, 70)
        dim_pen = pg.mkPen(140, 140, 140, 90, width=1.0)
        highlight_brush = pg.mkBrush(255, 225, 120, 255)
        highlight_pen = pg.mkPen(255, 255, 255, 255, width=2.0)

        spots: list[dict[str, Any]] = []
        for index, (rpm, ve) in enumerate(zip(rpm_values, ve_values)):
            is_selected = selected_series_id == series_id and selected_index == index
            brush = base_brushes.get(series_id, pg.mkBrush(220, 220, 220, 220))
            pen = base_pens.get(series_id, pg.mkPen(220, 220, 220, 220, width=1.0))
            size = 8 if series_id == "table" else 7 if series_id == "average" else 6
            if has_selection and not is_selected:
                brush = dim_brush
                pen = dim_pen
            if is_selected:
                brush = highlight_brush
                pen = highlight_pen
                size = 11
            spots.append({"pos": (float(rpm), float(ve)), "brush": brush, "pen": pen, "size": size})
        return spots

    @staticmethod
    def _format_value(value: float | None, suffix: str = "") -> str:
        if value is None:
            return "n/a"
        return f"{value:.2f}{suffix}"

    def _format_selected_point_text(self, series: dict[str, Any], index: int) -> str:
        name = str(series.get("name", "Point"))
        rpm = float(series["rpm"][index])
        map_val = float(series["map"][index])
        ve = float(series["ve"][index])
        ve_raw = series.get("ve_raw", [None])[index]
        ve_scaled = series.get("ve_scaled", [None])[index]
        afr = series.get("afr", [None])[index]
        afr_target = series.get("afr_target", [None])[index]
        afr_error = series.get("afr_error", [None])[index]

        lines = [
            f"Selected point ({name})",
            f"RPM {rpm:.1f} | MAP {map_val:.1f} kPa",
        ]
        series_id = str(series.get("series_id"))
        if series_id == "corrected":
            lines.append(f"VE scaled: {self._format_value(ve_scaled, '%')}")
            lines.append(f"VE unscaled: {self._format_value(ve_raw, '%')}")
        elif series_id == "table":
            lines.append(f"Selected row value: {ve:.2f}")
        elif series_id in {"afr", "afr_target", "afr_error", "knock"}:
            lines.append(f"Value: {ve:.2f}")
        else:
            lines.append(f"VE: {ve:.2f}%")
        lines.append(f"AFR: {self._format_value(afr)}")
        lines.append(f"AFR target: {self._format_value(afr_target)}")
        lines.append(f"AFR error: {self._format_value(afr_error)}")
        return "\n".join(lines)

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
            self.cursor_label.setText("Cursor: RPM -, Value -")
            return

        mouse_point = vb.mapSceneToView(pos)
        rpm = float(mouse_point.x())
        ve = float(mouse_point.y())
        self.crosshair_vline.setPos(rpm)
        self.crosshair_hline.setPos(ve)
        self.crosshair_vline.show()
        self.crosshair_hline.show()
        self.cursor_label.setText(f"Cursor: RPM {rpm:.1f}, Value {ve:.2f}")

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

        nearest: tuple[float, str, int] | None = None
        for series in self._point_sets:
            series_id = str(series.get("series_id", ""))
            for index, (rpm, ve) in enumerate(zip(series.get("rpm", []), series.get("ve", []))):
                scene_point = vb.mapViewToScene(pg.Point(float(rpm), float(ve)))
                distance = (scene_point.x() - click_pos.x()) ** 2 + (scene_point.y() - click_pos.y()) ** 2
                if nearest is None or distance < nearest[0]:
                    nearest = (distance, series_id, index)

        max_pick_distance_squared = 12.0 ** 2
        if nearest is None or nearest[0] > max_pick_distance_squared:
            self._selected_point = None
            self._refresh_point_styles()
            self.selected_point_label.setText("Selected point: none")
            if callable(self.on_point_selected):
                self.on_point_selected(None, None, None)
            return

        _, series_id, index = nearest
        selected_series = self._get_series(series_id)
        if selected_series is None:
            return

        self._selected_point = (series_id, index)
        self._refresh_point_styles()
        self.selected_point_label.setText(self._format_selected_point_text(selected_series, index))
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
        self.setRowCount(1)
        self.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.verticalHeader().setVisible(False)

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        current_column = self.currentColumn()
        if event.matches(QKeySequence.StandardKey.Undo):
            if callable(self.on_undo):
                self.on_undo(current_column)
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
        amount = float(steps * 0.1 if delta > 0 else -steps * 0.1)
        self.on_adjust_value(current_column, amount)
        event.accept()


class RowVisualizationPreferencesDialog(QDialog):
    """Preferences dialog for row visualization series visibility by table type."""

    SERIES_LABELS: dict[str, str] = {
        "table": "Selected Row Data",
        "raw": "Raw VE1",
        "corrected": "EGO Corrected VE",
        "average": "Average from Log",
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
        self.log_file: Path | None = None
        self.log_df: Any | None = None
        self.current_table: TableData | None = None
        self.current_x_axis: AxisVector | None = None
        self.current_y_axis: AxisVector | None = None
        self.log_viewer_windows: list[LogViewerWindow] = []
        self.selected_table_row_idx: int | None = None
        self.pending_row_values: list[float] = []
        self.row_default_values: list[float] = []
        self.row_edit_undo_stack: list[list[float]] = []
        self.average_line_data: dict[str, Any] | None = None

        self.recent_tune_files: list[Path] = []
        self.recent_log_files: list[Path] = []

        self.favorite_tables: set[str] = set()
        self.show_favorites_only = True
        self.show_tunerstudio_names = True
        self.row_viz_preferences: dict[str, dict[str, bool]] = self._default_row_viz_preferences()

        self._load_recent_files()
        self._load_favorites()
        self._load_row_viz_preferences()

        self._create_menu()
        self._create_layout()
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        QTimer.singleShot(0, self._auto_size_table_and_row_viz)

    def _create_menu(self) -> None:
        file_menu = self.menuBar().addMenu("&File")

        open_tune_action = file_menu.addAction("Open &Tune File...")
        open_tune_action.triggered.connect(self._open_tune_file)

        self.recent_tunes_menu = file_menu.addMenu("Recent &Tune Files")
        self._update_recent_tunes_menu()

        open_log_action = file_menu.addAction("Open &Log File...")
        open_log_action.triggered.connect(self._open_log_file)

        self.recent_logs_menu = file_menu.addMenu("Recent &Log Files")
        self._update_recent_logs_menu()

        open_log_viewer_action = file_menu.addAction("Open Log &Viewer...")
        open_log_viewer_action.triggered.connect(self._open_log_viewer)

        file_menu.addSeparator()
        exit_action = file_menu.addAction("E&xit")
        exit_action.triggered.connect(self.close)

        edit_menu = self.menuBar().addMenu("&Edit")
        preferences_action = edit_menu.addAction("&Preferences...")
        preferences_action.triggered.connect(self._open_preferences)

        view_menu = self.menuBar().addMenu("&View")
        self.tunerstudio_names_action = view_menu.addAction("&TunerStudio Table Names")
        self.tunerstudio_names_action.setCheckable(True)
        self.tunerstudio_names_action.setChecked(True)
        self.tunerstudio_names_action.triggered.connect(self._on_tunerstudio_names_toggled)

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

    def _load_favorites(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        favorites = settings.value("favorite_tables", [])
        self.favorite_tables = set(favorites) if favorites else set()

    def _save_favorites(self) -> None:
        settings = QSettings("GyattOTune", "GyattOTune")
        settings.setValue("favorite_tables", list(self.favorite_tables))

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

    def _one_d_table_plot_data(self) -> tuple[list[float], list[float], str]:
        if self.current_table is None:
            return [], [], "Index"

        if self.current_table.rows == 1:
            y_values = [float(v) for v in self.current_table.values[0]]
            if self.current_x_axis is not None and len(self.current_x_axis.values) == len(y_values):
                x_values = [float(v) for v in self.current_x_axis.values]
                x_label = self.current_x_axis.name or "X"
            else:
                x_values = [float(i + 1) for i in range(len(y_values))]
                x_label = "Index"
            return x_values, y_values, x_label

        y_values = [float(row[0]) for row in self.current_table.values]
        if self.current_y_axis is not None and len(self.current_y_axis.values) == len(y_values):
            x_values = [float(v) for v in self.current_y_axis.values]
            x_label = self.current_y_axis.name or "Y"
        else:
            x_values = [float(i + 1) for i in range(len(y_values))]
            x_label = "Index"
        return x_values, y_values, x_label

    def _build_1d_table_visualization_payload(self) -> dict[str, Any]:
        x_values, y_values, x_label = self._one_d_table_plot_data()
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
            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if rpm_channel is None and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if knock_channel is None and ('knock in' in col_lower or 'knock_in' in col_lower or 'knockin' in col_lower):
                    knock_channel = str(col)
            if knock_channel is None:
                for col in self.log_df.columns:
                    col_lower = str(col).lower()
                    if 'knock' in col_lower and 'threshold' not in col_lower:
                        knock_channel = str(col)
                        break

            if rpm_channel and knock_channel:
                rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
                knock_series = self._to_numeric_series(self.log_df[knock_channel])
                if rpm_series is not None and knock_series is not None:
                    import numpy as np

                    rpm_vals = np.asarray(rpm_series, dtype=float)
                    knock_vals = np.asarray(knock_series, dtype=float)
                    finite = np.isfinite(rpm_vals) & np.isfinite(knock_vals)
                    rpm_vals = rpm_vals[finite]
                    knock_vals = knock_vals[finite]
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
                                "afr": [None] * len(rpm_vals),
                                "afr_target": [None] * len(rpm_vals),
                                "afr_error": [None] * len(rpm_vals),
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
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Tune Load Error", f"Could not load tune file:\n{exc}")
            return

        self.table_list.clear()
        if self.tune_data.tables:
            self._update_table_display()
        else:
            self.table_list.addItem("No tables found")
            self.table_grid.clear()
            self.table_grid.setRowCount(0)
            self.table_grid.setColumnCount(0)
            self.table_meta.setText("No tables found in this tune file")

        self.statusBar().showMessage(f"Loaded tune: {file_path.name}")
        self._refresh_workspace_text()
        self._add_recent_tune_file(file_path)

    def _open_recent_log_file(self, file_path: Path) -> None:
        if not file_path.exists():
            QMessageBox.warning(self, "File Not Found", f"The file {file_path} no longer exists.")
            self.recent_log_files.remove(file_path)
            self._save_recent_files()
            self._update_recent_logs_menu()
            return

        # Show progress dialog for large files
        progress = QProgressBar()
        progress.setRange(0, 0)  # Indeterminate progress
        progress.setMinimumWidth(300)
        
        # Create a modal dialog with the progress bar
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel
        progress_dialog = QDialog(self)
        progress_dialog.setWindowTitle("Loading Log File")
        progress_dialog.setModal(True)
        progress_dialog.setFixedSize(350, 100)
        
        layout = QVBoxLayout(progress_dialog)
        layout.addWidget(QLabel(f"Loading {file_path.name}..."))
        layout.addWidget(progress)
        progress_dialog.show()
        
        # Force UI update
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            progress_dialog.close()
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return

        progress_dialog.close()

        self.log_file = file_path
        self.statusBar().showMessage(
            f"Loaded log: {file_path.name} ({len(log_df):,} rows, parser={parse_result.parser_used}, enc={parse_result.encoding})"
        )
        self.log_df = log_df
        self._add_recent_log_file(file_path)

    def _update_table_display(self) -> None:
        """Update the table list display based on current settings."""
        if not self.tune_data:
            return

        self.show_tunerstudio_names = self.tunerstudio_names_action.isChecked()

        # Update Favorites tab
        self.favorites_list.clear()
        favorites_items = []
        for table_name in sorted(self.tune_data.tables.keys()):
            if table_name not in self.favorite_tables:
                continue
            display_name = self._get_tunerstudio_name(table_name) if self.show_tunerstudio_names else table_name
            favorites_items.append((table_name, display_name))

        if favorites_items:
            for table_name, display_name in favorites_items:
                item = QListWidgetItem(display_name)
                item.setData(Qt.ItemDataRole.UserRole, table_name)
                self.favorites_list.addItem(item)
            self.favorites_list.setCurrentRow(0)
        else:
            self.favorites_list.addItem("No favorite tables")

        # Update All Tables tab
        self.table_list.clear()
        all_items = []
        for table_name in sorted(self.tune_data.tables.keys()):
            display_name = self._get_tunerstudio_name(table_name) if self.show_tunerstudio_names else table_name
            if table_name in self.favorite_tables:
                display_name = f"★ {display_name}"
            all_items.append((table_name, display_name))

        if all_items:
            for table_name, display_name in all_items:
                item = QListWidgetItem(display_name)
                item.setData(Qt.ItemDataRole.UserRole, table_name)
                self.table_list.addItem(item)
            self.table_list.setCurrentRow(0)
        else:
            self.table_list.addItem("No tables found")

    def _on_tunerstudio_names_toggled(self) -> None:
        """Handle View menu TunerStudio names toggle."""
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
        # Determine which list was clicked
        if self.table_tabs.currentIndex() == 0:
            # Favorites tab
            list_widget = self.favorites_list
        else:
            # All tables tab
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

    def _open_log_viewer(self) -> None:
        viewer = LogViewerWindow()
        viewer.show()
        self.log_viewer_windows.append(viewer)

    def _create_layout(self) -> None:
        root = QWidget(self)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(8, 8, 8, 8)

        # Create main splitter for left panel and right panel
        splitter = QSplitter(Qt.Orientation.Horizontal)
        root_layout.addWidget(splitter)

        # Left panel - Tabbed Tune Tables (Favorites/All)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(QLabel("Tune Tables"))

        # Create tabbed widget for Favorites and All tables
        self.table_tabs = QTabWidget()

        # Favorites tab
        favorites_tab = QWidget()
        favorites_layout = QVBoxLayout(favorites_tab)
        favorites_layout.setContentsMargins(0, 0, 0, 0)
        
        self.favorites_list = QListWidget()
        self.favorites_list.addItem("Load a tune file to see favorites")
        self.favorites_list.currentItemChanged.connect(self._on_table_selected)
        self.favorites_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.favorites_list.customContextMenuRequested.connect(self._show_table_context_menu)
        favorites_layout.addWidget(self.favorites_list)
        self.table_tabs.addTab(favorites_tab, "Favorites")

        # All tables tab
        all_tables_tab = QWidget()
        all_tables_layout = QVBoxLayout(all_tables_tab)
        all_tables_layout.setContentsMargins(0, 0, 0, 0)
        
        self.table_list = QListWidget()
        self.table_list.addItem("Load a tune file to see parsed tables")
        self.table_list.currentItemChanged.connect(self._on_table_selected)
        self.table_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_list.customContextMenuRequested.connect(self._show_table_context_menu)
        all_tables_layout.addWidget(self.table_list)
        self.table_tabs.addTab(all_tables_tab, "All Tables")

        left_layout.addWidget(self.table_tabs)

        # Right panel - Table display with editor
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.table_meta = QLabel("No table selected")
        right_layout.addWidget(self.table_meta)

        controls_layout = QHBoxLayout()
        self.transpose_checkbox = QCheckBox("Transpose Table")
        self.transpose_checkbox.toggled.connect(self._refresh_current_table_view)
        controls_layout.addWidget(self.transpose_checkbox)

        self.swap_axes_checkbox = QCheckBox("Swap Axes Labels")
        self.swap_axes_checkbox.toggled.connect(self._refresh_current_table_view)
        controls_layout.addWidget(self.swap_axes_checkbox)

        controls_layout.addStretch(1)
        right_layout.addLayout(controls_layout)

        self.table_grid = CopyPasteTableWidget()
        self.table_grid.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table_grid.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_grid.itemSelectionChanged.connect(self._on_table_grid_row_selection_changed)
        self.table_grid.cellClicked.connect(self._on_table_grid_cell_clicked)
        self.table_grid.currentCellChanged.connect(self._on_table_grid_current_cell_changed)
        self.table_grid.horizontalHeader().setVisible(False)
        self.table_grid.verticalHeader().setVisible(True)  # leftmost axis labels
        
        self.table_sections_splitter = QSplitter(Qt.Orientation.Horizontal)
        right_layout.addWidget(self.table_sections_splitter, 1)

        table_grid_container = QWidget()
        self.table_grid_container = table_grid_container
        table_grid_layout = QVBoxLayout(table_grid_container)
        table_grid_layout.setContentsMargins(0, 0, 0, 0)
        table_grid_layout.addWidget(self.table_grid)
        self.table_sections_splitter.addWidget(table_grid_container)

        self.table_details_splitter = QSplitter(Qt.Orientation.Vertical)
        self.table_sections_splitter.addWidget(self.table_details_splitter)

        self.table_row_panel = RowVisualizationPanel()
        self.table_row_panel.on_point_selected = self._on_table_row_plot_point_selected
        self.table_details_splitter.addWidget(self.table_row_panel)

        self.table_row_editor_group = QGroupBox("Selected Row Editor")
        row_editor_layout = QVBoxLayout(self.table_row_editor_group)

        row_editor_controls = QHBoxLayout()
        self.table_row_label = QLabel("Selected row: none")
        row_editor_controls.addWidget(self.table_row_label)
        row_editor_controls.addStretch(1)
        self.generate_average_button = QPushButton("Generate Average from Log")
        self.generate_average_button.clicked.connect(self._generate_average_line_from_log)
        row_editor_controls.addWidget(self.generate_average_button)
        self.apply_row_changes_button = QPushButton("Write Row To Table Grid")
        self.apply_row_changes_button.clicked.connect(self._apply_pending_row_changes)
        row_editor_controls.addWidget(self.apply_row_changes_button)
        self.revert_row_changes_button = QPushButton("Revert Row")
        self.revert_row_changes_button.clicked.connect(self._revert_pending_row_values)
        row_editor_controls.addWidget(self.revert_row_changes_button)
        row_editor_layout.addLayout(row_editor_controls)

        self.table_row_editor = RowEditorTableWidget()
        self.table_row_editor.on_adjust_value = self._adjust_pending_row_value
        self.table_row_editor.on_undo = self._undo_pending_row_edit
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
        self.generate_average_button.setEnabled(False)
        self.apply_row_changes_button.setEnabled(False)
        self.revert_row_changes_button.setEnabled(False)

        self.table_row_status = QLabel("Select a VE table and click a row to view and edit it.")
        row_editor_container_layout.addWidget(self.table_row_status)

        self.table_details_splitter.addWidget(row_editor_container)

        self.table_sections_splitter.setStretchFactor(0, 6)
        self.table_sections_splitter.setStretchFactor(1, 4)
        self.table_details_splitter.setStretchFactor(0, 3)
        self.table_details_splitter.setStretchFactor(1, 2)
        self.table_sections_splitter.setSizes([900, 600])
        self.table_details_splitter.setSizes([420, 260])

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)

        self.setCentralWidget(root)

    def _create_ve_comparison_tab(self) -> None:
        """Create the VE comparison tab for analyzing VE table vs log data."""
        layout = QVBoxLayout(self.ve_comparison_tab)

        # Controls
        controls_layout = QHBoxLayout()
        self.ve_table_combo = QComboBox()
        self.ve_table_combo.currentTextChanged.connect(self._on_ve_table_selection_changed)
        controls_layout.addWidget(QLabel("VE Table:"))
        controls_layout.addWidget(self.ve_table_combo, 1)

        controls_layout.addStretch(1)
        layout.addLayout(controls_layout)

        # VE correction grid
        self.ve_correction_grid = QTableWidget()
        self.ve_correction_grid.itemSelectionChanged.connect(self._on_ve_grid_selection_changed)
        self.ve_correction_grid.installEventFilter(self)
        layout.addWidget(self.ve_correction_grid, 3)

        # Row adjustment controls
        adjustment_layout = QHBoxLayout()
        adjustment_layout.addWidget(QLabel("Selected Row:"))
        self.selected_row_label = QLabel("None")
        adjustment_layout.addWidget(self.selected_row_label)

        adjustment_layout.addWidget(QLabel("Adjustment:"))
        self.adjustment_spin = QDoubleSpinBox()
        self.adjustment_spin.setRange(-2.0, 2.0)
        self.adjustment_spin.setSingleStep(0.01)
        self.adjustment_spin.setValue(1.0)
        self.adjustment_spin.setSuffix("x")
        self.adjustment_spin.valueChanged.connect(self._apply_row_adjustment)
        adjustment_layout.addWidget(self.adjustment_spin)

        self.apply_adjustment_button = QPushButton("Apply to Row")
        self.apply_adjustment_button.clicked.connect(self._apply_row_adjustment)
        adjustment_layout.addWidget(self.apply_adjustment_button)

        adjustment_layout.addStretch(1)
        layout.addLayout(adjustment_layout)

        self.ve_row_panel = RowVisualizationPanel()
        layout.addWidget(self.ve_row_panel)

        # Status and info
        self.ve_comparison_status = QLabel("Select a VE table to view correction factors.")
        layout.addWidget(self.ve_comparison_status)

        # Initialize the row visualization
        self._clear_row_visualization()

    def _update_ve_table_combo(self) -> None:
        """Update the VE table combo box with available VE tables."""
        if not self.tune_data:
            return

        self.ve_table_combo.clear()
        ve_tables = []

        # Find VE tables (those containing 've' in the name, case insensitive)
        for table_name in self.tune_data.tables.keys():
            if 've' in table_name.lower():
                display_name = self._get_tunerstudio_name(table_name)
                ve_tables.append((table_name, display_name))

        # Sort by display name
        ve_tables.sort(key=lambda x: x[1])

        for table_name, display_name in ve_tables:
            self.ve_table_combo.addItem(display_name, table_name)

        if ve_tables:
            self.ve_table_combo.setCurrentIndex(0)

        # If table grid has a selected table, pick it for VE comparison by default
        if self.current_table:
            idx = self.ve_table_combo.findData(self.current_table.name)
            if idx != -1:
                self.ve_table_combo.setCurrentIndex(idx)

        # Update the correction grid
        self._update_ve_correction_grid()

    def _on_ve_table_selection_changed(self) -> None:
        self._update_ve_correction_grid()

    def _update_ve_correction_grid(self) -> None:
        """Update the VE correction grid with correction factors for the entire table."""
        if not self.tune_data or self.log_df is None or self.log_df.empty:
            self.ve_comparison_status.setText("Load both a tune file and log file to see VE correction factors.")
            return

        # Get the selected table
        current_table_name = self.ve_table_combo.currentData()
        if not current_table_name or current_table_name not in self.tune_data.tables:
            self.ve_comparison_status.setText("Select a VE table to view correction factors.")
            return

        table = self.tune_data.tables[current_table_name]

        # Get table axes
        x_axis, y_axis = self.tune_data.resolve_table_axes(table)
        if not x_axis or not y_axis:
            self.ve_comparison_status.setText("Selected table is missing axis information.")
            return

        # Calculate correction factors for entire grid
        correction_factors = self._calculate_ve_correction_factors(table, x_axis, y_axis)

        # Display the correction grid
        self._render_ve_correction_grid(table, x_axis, y_axis, correction_factors)

        if correction_factors:
            self.ve_comparison_status.setText(
                f"VE correction factors for {table.name} | "
                f"Range: {min(min(row) for row in correction_factors):.2f}x - {max(max(row) for row in correction_factors):.2f}x"
            )
        else:
            self.ve_comparison_status.setText(f"VE table values shown for {table.name}. No logged data available for correction calculation.")

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

    def _calculate_ve_correction_factors(self, table: TableData, x_axis: AxisVector, y_axis: AxisVector) -> list[list[float]] | None:
        """Calculate correction factors for the entire VE table grid."""
        if self.log_df is None or self.log_df.empty:
            return None

        try:
            # Get required channels
            rpm_channel = None
            map_channel = None
            ve1_channel = None
            ego_cor1_channel = None

            # Auto-detect channels
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

            if not rpm_channel or not map_channel or not ve1_channel:
                return None

            # Convert to numeric series
            rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
            map_series = self._to_numeric_series(self.log_df[map_channel])
            ve_series = self._to_numeric_series(self.log_df[ve1_channel])

            if rpm_series is None or map_series is None or ve_series is None:
                return None

            # Apply EGO correction
            if ego_cor1_channel is not None:
                ego_series = self._to_numeric_series(self.log_df[ego_cor1_channel])
                if ego_series is not None:
                    ve_series = ve_series * (ego_series / 100.0)

            # Initialize correction factor grid
            correction_factors = [[1.0 for _ in range(table.cols)] for _ in range(table.rows)]

            # For each table cell, calculate average correction factor from nearby log data
            for row_idx in range(table.rows):
                map_value = float(y_axis.values[row_idx])
                for col_idx in range(table.cols):
                    rpm_value = float(x_axis.values[col_idx])
                    table_ve = float(table.values[row_idx][col_idx])

                    # Find log data points near this RPM/MAP combination
                    rpm_tolerance = rpm_value * 0.1  # 10% tolerance
                    map_tolerance = 10  # 10 kPa tolerance

                    rpm_mask = (rpm_series >= rpm_value - rpm_tolerance) & (rpm_series <= rpm_value + rpm_tolerance)
                    map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                    combined_mask = rpm_mask & map_mask

                    nearby_ve_values = ve_series[combined_mask]

                    if len(nearby_ve_values) > 0:
                        # Calculate correction factor: logged VE / table VE
                        avg_logged_ve = float(nearby_ve_values.mean())
                        if table_ve > 0:
                            correction_factors[row_idx][col_idx] = avg_logged_ve / table_ve
                        else:
                            correction_factors[row_idx][col_idx] = 1.0

            return correction_factors

        except Exception:
            return None

    def _render_ve_correction_grid(self, table: TableData, x_axis: AxisVector, y_axis: AxisVector, correction_factors: list[list[float]] | None) -> None:
        """Render the VE correction factor grid."""
        row_count = len(y_axis.values)
        col_count = len(x_axis.values)

        self.ve_correction_grid.clear()
        self.ve_correction_grid.setRowCount(row_count + 1)
        self.ve_correction_grid.setColumnCount(col_count)
        self.ve_correction_grid.setVerticalHeaderLabels([f"{float(v):g}" for v in reversed(y_axis.values)] + ["X axis"])

        # Set column headers (RPM values)
        for c, rpm in enumerate(x_axis.values):
            self.ve_correction_grid.setHorizontalHeaderItem(c, QTableWidgetItem(f"{float(rpm):g}"))

        # Fill grid with correction factors or table values
        display_matrix = list(reversed(table.values)) if correction_factors is None else list(reversed(correction_factors))

        for r, row in enumerate(display_matrix):
            for c, value in enumerate(row):
                if correction_factors is not None:
                    # Show correction factor
                    item = QTableWidgetItem(f"{value:.2f}x")
                    # Color based on correction factor
                    if value < 0.9:
                        item.setBackground(QColor(255, 200, 200))  # Red for too low
                    elif value > 1.1:
                        item.setBackground(QColor(200, 255, 200))  # Green for too high
                    else:
                        item.setBackground(QColor(200, 200, 255))  # Blue for close
                else:
                    # Show table VE value
                    item = QTableWidgetItem(f"{value:g}")
                    item.setBackground(QColor(240, 240, 240))  # Gray for no correction data

                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.ve_correction_grid.setItem(r, c, item)

        # Add X axis row
        axis_row = row_count
        for c, label in enumerate([f"{float(rpm):g}" for rpm in x_axis.values]):
            axis_item = QTableWidgetItem(label)
            axis_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            axis_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            axis_item.setBackground(QColor(60, 60, 60))
            self.ve_correction_grid.setItem(axis_row, c, axis_item)

        self.ve_correction_grid.resizeColumnsToContents()
        self.ve_correction_grid.resizeRowsToContents()

    def _on_ve_grid_selection_changed(self) -> None:
        """Handle selection changes in the VE correction grid."""
        selected_items = self.ve_correction_grid.selectedItems()
        if not selected_items:
            self.selected_row_label.setText("None")
            self._clear_row_visualization()
            return

        # Find the row of the first selected item
        selected_row = selected_items[0].row()

        # Check if this is a data row (not the X axis row at the bottom)
        if selected_row >= self.ve_correction_grid.rowCount() - 1:
            self.selected_row_label.setText("None")
            self._clear_row_visualization()
            return

        # Get the MAP value for this row
        row_header = self.ve_correction_grid.verticalHeaderItem(selected_row)
        if not row_header:
            return

        map_value = float(row_header.text())
        self.selected_row_label.setText(f"MAP {map_value} kPa")

        # Update the adjustment spinbox to current correction factor
        # For now, set to 1.0 as default
        self.adjustment_spin.setValue(1.0)

        # Show data visualization for this row
        self._update_row_data_visualization(selected_row, map_value)

    def _update_row_data_visualization(self, row_idx: int, map_value: float) -> None:
        """Update the data visualization for the selected row."""
        if not self.tune_data or self.log_df is None or self.log_df.empty:
            self._clear_row_visualization()
            return

        current_table_name = self.ve_table_combo.currentData()
        if not current_table_name or current_table_name not in self.tune_data.tables:
            self._clear_row_visualization()
            return

        table = self.tune_data.tables[current_table_name]
        x_axis, y_axis = self.tune_data.resolve_table_axes(table)
        if not x_axis or not y_axis:
            self._clear_row_visualization()
            return

        # Get table VE values for this row
        table_row_idx = len(y_axis.values) - 1 - row_idx  # Reverse indexing
        table_ve_values = [float(val) for val in table.values[table_row_idx]]

        payload = self._build_row_visualization_payload(map_value, x_axis.values, table_ve_values)
        self.ve_row_panel.set_row_data(payload)

    def _clear_row_visualization(self) -> None:
        """Clear the row data visualization."""
        self.table_row_panel.clear_visualization()

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
            for rpm_value in rpm_point_values:
                nearest_index = min(range(len(filtered_rpm_values)), key=lambda idx: abs(filtered_rpm_values[idx] - rpm_value))
                table_afr_values.append(filtered_afr_values[nearest_index])
                table_afr_target_values.append(filtered_afr_target_values[nearest_index])
                table_afr_error_values.append(filtered_afr_error_values[nearest_index])

            point_sets[0]["afr"] = table_afr_values
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

            for col in self.log_df.columns:
                col_lower = str(col).lower()
                if not rpm_channel and ('rpm' in col_lower or 'engine speed' in col_lower):
                    rpm_channel = str(col)
                if not map_channel and ('map' in col_lower):
                    map_channel = str(col)
                if knock_channel is None and ('knock in' in col_lower or 'knock_in' in col_lower or 'knockin' in col_lower):
                    knock_channel = str(col)

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

            if len(filtered_rpm) == 0 and map_tolerance < 20.0:
                map_tolerance = 20.0
                map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
                filtered_rpm = rpm_series[map_mask]
                filtered_map = map_series[map_mask]
                filtered_knock = knock_series[map_mask]

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

            point_sets.append(
                {
                    "series_id": "knock",
                    "name": "Knock In",
                    "rpm": filtered_rpm_values,
                    "ve": filtered_knock_values,
                    "map": filtered_map_values,
                    "ve_raw": filtered_knock_values,
                    "ve_scaled": filtered_knock_values,
                    "afr": [None] * len(filtered_rpm_values),
                    "afr_target": [None] * len(filtered_rpm_values),
                    "afr_error": [None] * len(filtered_rpm_values),
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
            self._update_table_grid_row_visualization()
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
        """Generate an average VE line from log data for the current MAP."""
        if self.current_table is None or "ve" not in self.current_table.name.lower():
            self.table_row_status.setText("Average line generation is currently available for VE rows.")
            return

        if self.log_df is None or self.log_df.empty:
            self.table_row_status.setText("No log data loaded. Cannot generate average.")
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
        
        if not rpm_channel or not map_channel or not ve1_channel:
            self.table_row_status.setText("Required log columns not found (RPM, MAP, VE1).")
            return
        
        # Convert to numeric series
        rpm_series = self._to_numeric_series(self.log_df[rpm_channel])
        map_series = self._to_numeric_series(self.log_df[map_channel])
        ve_series = self._to_numeric_series(self.log_df[ve1_channel])
        
        if rpm_series is None or map_series is None or ve_series is None:
            self.table_row_status.setText("Could not convert log data to numeric values.")
            return
        
        # Apply EGO correction if available
        ve_corrected_series = ve_series.copy()
        if ego_cor1_channel is not None:
            ego_series = self._to_numeric_series(self.log_df[ego_cor1_channel])
            if ego_series is not None:
                ve_corrected_series = ve_series * (ego_series / 100.0)
        
        # Calculate adaptive tolerance like in _build_row_visualization_payload
        map_tolerance = 10.0
        if self.current_y_axis is not None and len(self.current_y_axis.values) > 1:
            sorted_bins = sorted(float(v) for v in self.current_y_axis.values)
            min_spacing = min(abs(b - a) for a, b in zip(sorted_bins, sorted_bins[1:]))
            map_tolerance = max(10.0, min_spacing * 0.75)
        
        # Filter to current MAP bin
        map_mask = (map_series >= map_value - map_tolerance) & (map_series <= map_value + map_tolerance)
        filtered_rpm = rpm_series[map_mask]
        filtered_ve = ve_corrected_series[map_mask]
        
        if len(filtered_rpm) == 0:
            self.table_row_status.setText(f"No log data near MAP {map_value} kPa.")
            return
        
        # Group by RPM bins and calculate averages
        rpm_bins = [float(v) for v in self.current_x_axis.values]
        average_rpm_values = []
        average_ve_values = []
        
        for target_rpm in rpm_bins:
            # Use bins centered on each RPM value (±500 RPM range, or proportional to spacing)
            rpm_tolerance = 500.0
            if len(rpm_bins) > 1:
                sorted_rpms = sorted(rpm_bins)
                min_spacing = min(abs(b - a) for a, b in zip(sorted_rpms, sorted_rpms[1:]))
                rpm_tolerance = max(300.0, min_spacing * 0.4)
            
            rpm_mask = (filtered_rpm >= target_rpm - rpm_tolerance) & (filtered_rpm <= target_rpm + rpm_tolerance)
            matching_ve = filtered_ve[rpm_mask]
            
            if len(matching_ve) > 0:
                avg_ve = float(matching_ve.mean())
                average_rpm_values.append(target_rpm)
                average_ve_values.append(avg_ve)
        
        if not average_rpm_values:
            self.table_row_status.setText("Could not calculate averages from log data for this MAP.")
            return
        
        # Store the average data for display
        self.average_line_data = {
            "series_id": "average",
            "name": "Average from Log",
            "rpm": average_rpm_values,
            "ve": average_ve_values,
            "map": [map_value] * len(average_rpm_values),
            "ve_raw": average_ve_values,
            "ve_scaled": average_ve_values,
            "afr": [None] * len(average_rpm_values),
            "afr_target": [None] * len(average_rpm_values),
            "afr_error": [None] * len(average_rpm_values),
        }
        
        self.table_row_status.setText(f"Generated average from {len(filtered_rpm)} log points near MAP {map_value} kPa.")
        self._update_table_grid_row_visualization()

    def _update_table_grid_row_visualization(self) -> None:
        available, message = self._table_row_editing_available()
        if not available:
            self._clear_table_row_editor(message)
            return

        if self.current_table is not None and self._is_current_table_1d():
            table_type = self._row_table_type()
            payload = self._build_1d_table_visualization_payload()
            payload["available_series"] = self._row_viz_available_series(table_type, payload)
            payload["series_visibility"] = self.row_viz_preferences.get(table_type, {})
            self.table_row_panel.set_row_data(payload)
            self.table_row_label.setText("Selected row: n/a (1D table)")
            self.table_row_editor.clear()
            self.table_row_editor.setRowCount(1)
            self.table_row_editor.setColumnCount(0)
            self.table_row_editor.setEnabled(False)
            self.generate_average_button.setEnabled(False)
            self.apply_row_changes_button.setEnabled(False)
            self.revert_row_changes_button.setEnabled(False)
            self.table_row_status.setText(f"Viewing full 1D data for {self.current_table.name}. Selected-row editing is disabled.")
            return

        if self.selected_table_row_idx is None or not self.pending_row_values:
            self._clear_table_row_editor("Click a row in the table grid to view and edit it.")
            return

        if self.current_y_axis is None or self.current_x_axis is None or self.current_table is None:
            self._clear_table_row_editor("Select a table row in the table grid to view and edit it.")
            return

        table_name = self.current_table.name.lower()
        table_type = self._row_table_type()
        map_value = float(self.current_y_axis.values[self.selected_table_row_idx])
        if "knock" in table_name:
            payload = self._build_knock_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
        elif "afr" in table_name:
            payload = self._build_afr_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
        else:
            payload = self._build_row_visualization_payload(map_value, self.current_x_axis.values, self.pending_row_values)
            payload["y_label"] = "VE %"
        
        # Add average line if it has been generated
        if self.average_line_data is not None and "ve" in table_name:
            point_sets = payload.get("point_sets", [])
            point_sets.append(self.average_line_data)
            payload["point_sets"] = point_sets

        payload["available_series"] = self._row_viz_available_series(table_type, payload)
        payload["series_visibility"] = self.row_viz_preferences.get(table_type, {})
        
        self.table_row_panel.set_row_data(payload)
        self.table_row_status.setText(
            f"Editing MAP {map_value:g} kPa for {self.current_table.name}. Click 'Write Row To Table Grid' to commit changes."
        )

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
        self.average_line_data = None
        self.table_row_label.setText("Selected row: none")
        self.table_row_editor.clear()
        self.table_row_editor.setRowCount(1)
        self.table_row_editor.setColumnCount(0)
        self.table_row_editor.setEnabled(False)
        self.generate_average_button.setEnabled(False)
        self.apply_row_changes_button.setEnabled(False)
        self.revert_row_changes_button.setEnabled(False)
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

        self._load_selected_table_row(source_row)

    def _on_table_grid_cell_clicked(self, row: int, column: int) -> None:
        _ = column
        if self._is_current_table_1d():
            self._update_table_grid_row_visualization()
            return
        source_row = self._display_row_to_source_row(row)
        if source_row is None:
            return
        self._load_selected_table_row(source_row)

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
        self._load_selected_table_row(source_row)

    def _load_selected_table_row(self, source_row: int) -> None:
        if self.current_table is None or self.current_x_axis is None or self.current_y_axis is None:
            return

        self.selected_table_row_idx = source_row
        self.pending_row_values = [float(value) for value in self.current_table.values[source_row]]
        self.row_default_values = [float(value) for value in self.pending_row_values]
        self.row_edit_undo_stack = []
        self.average_line_data = None
        map_value = float(self.current_y_axis.values[source_row])
        self.table_row_label.setText(f"Selected row: MAP {map_value:g} kPa")
        self._refresh_table_row_editor()
        self._update_table_grid_row_visualization()

    def _refresh_table_row_editor(self) -> None:
        self.table_row_editor.blockSignals(True)
        self.table_row_editor.clear()
        self.table_row_editor.setRowCount(1)
        self.table_row_editor.setColumnCount(len(self.pending_row_values))
        minimum, maximum = self._matrix_min_max([self.pending_row_values])
        span = max(maximum - minimum, 1e-9)
        if self.current_x_axis is not None:
            self.table_row_editor.setHorizontalHeaderLabels([f"{float(value):g}" for value in self.current_x_axis.values])
        for column_index, value in enumerate(self.pending_row_values):
            item = QTableWidgetItem(f"{value:g}")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(self._cell_color(value, minimum, span))
            self.table_row_editor.setItem(0, column_index, item)
        if self.pending_row_values:
            current_column = self.table_row_editor.currentColumn()
            target_column = current_column if 0 <= current_column < len(self.pending_row_values) else 0
            self.table_row_editor.setCurrentCell(0, target_column)
        self.table_row_editor.resizeColumnsToContents()
        self.table_row_editor.resizeRowsToContents()
        self.table_row_editor.setEnabled(True)
        self.generate_average_button.setEnabled(self.current_table is not None and "ve" in self.current_table.name.lower())
        self.apply_row_changes_button.setEnabled(True)
        self.revert_row_changes_button.setEnabled(True)
        self.table_row_editor.blockSignals(False)

    def _commit_pending_row_undo_state(self) -> None:
        self.row_edit_undo_stack.append([float(value) for value in self.pending_row_values])
        if len(self.row_edit_undo_stack) > 100:
            self.row_edit_undo_stack = self.row_edit_undo_stack[-100:]

    def _adjust_pending_row_value(self, column_index: int, amount: float) -> None:
        if column_index < 0 or column_index >= len(self.pending_row_values):
            return
        self._commit_pending_row_undo_state()
        self.pending_row_values[column_index] = float(self.pending_row_values[column_index] + amount)
        item = self.table_row_editor.item(0, column_index)
        if item is None:
            item = QTableWidgetItem()
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_row_editor.setItem(0, column_index, item)
        item.setText(f"{self.pending_row_values[column_index]:g}")
        self._update_table_grid_row_visualization()

    def _undo_pending_row_edit(self, column_index: int) -> None:
        if (
            column_index < 0
            or column_index >= len(self.pending_row_values)
            or column_index >= len(self.row_default_values)
        ):
            return
        self.pending_row_values[column_index] = float(self.row_default_values[column_index])
        item = self.table_row_editor.item(0, column_index)
        if item is None:
            item = QTableWidgetItem()
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_row_editor.setItem(0, column_index, item)
        item.setText(f"{self.pending_row_values[column_index]:g}")
        minimum, maximum = self._matrix_min_max([self.pending_row_values])
        span = max(maximum - minimum, 1e-9)
        for c, value in enumerate(self.pending_row_values):
            row_item = self.table_row_editor.item(0, c)
            if row_item is not None:
                row_item.setBackground(self._cell_color(value, minimum, span))
        self._update_table_grid_row_visualization()

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

        target_column: int | None = None
        if series_id == "table" and index is not None and 0 <= index < self.table_row_editor.columnCount():
            target_column = index
        elif rpm_value is not None and self.current_x_axis is not None and self.current_x_axis.values:
            target_column = min(
                range(len(self.current_x_axis.values)),
                key=lambda i: abs(float(self.current_x_axis.values[i]) - float(rpm_value)),
            )

        if target_column is None or target_column < 0 or target_column >= self.table_row_editor.columnCount():
            return

        self.table_row_editor.setCurrentCell(0, target_column)
        self.table_row_editor.setFocus()

    def _revert_pending_row_values(self) -> None:
        if not self.row_default_values:
            return
        self._commit_pending_row_undo_state()
        self.pending_row_values = [float(value) for value in self.row_default_values]
        self._refresh_table_row_editor()
        self._update_table_grid_row_visualization()

    def _apply_pending_row_changes(self) -> None:
        if (
            self.current_table is None
            or self.current_x_axis is None
            or self.current_y_axis is None
            or self.selected_table_row_idx is None
            or not self.pending_row_values
        ):
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



    def _apply_row_adjustment(self) -> None:
        """Apply the current adjustment factor to the selected row."""
        if not self.tune_data:
            return

        current_table_name = self.ve_table_combo.currentData()
        if not current_table_name or current_table_name not in self.tune_data.tables:
            return

        table = self.tune_data.tables[current_table_name]
        x_axis, y_axis = self.tune_data.resolve_table_axes(table)
        if not x_axis or not y_axis:
            return

        # Get selected row
        selected_items = self.ve_correction_grid.selectedItems()
        if not selected_items:
            return

        selected_row = selected_items[0].row()
        if selected_row >= self.ve_correction_grid.rowCount() - 1:
            return

        # Get the MAP value and find corresponding table row
        row_header = self.ve_correction_grid.verticalHeaderItem(selected_row)
        if not row_header:
            return

        map_value = float(row_header.text())
        table_row_idx = len(y_axis.values) - 1 - selected_row  # Reverse indexing

        # Apply adjustment to entire row
        adjustment = self.adjustment_spin.value()
        for col_idx in range(len(table.values[table_row_idx])):
            table.values[table_row_idx][col_idx] *= adjustment

        # Refresh the grid and summary
        self._update_ve_correction_grid()
        self._update_row_data_visualization(selected_row, map_value)

        # Refresh table grid if same table is selected
        if self.current_table and self.current_table.name == table.name:
            self._refresh_current_table_view()

    def eventFilter(self, source: QObject, event: QEvent) -> bool:
        """Handle keyboard events for fine-tuning adjustments."""
        if source == self.ve_correction_grid and event.type() == QEvent.Type.KeyPress:
            key_event = event  # type: ignore
            if key_event.key() == Qt.Key.Key_Left:
                self.adjustment_spin.setValue(self.adjustment_spin.value() - 0.01)
                self._apply_row_adjustment()
                return True
            elif key_event.key() == Qt.Key.Key_Right:
                self.adjustment_spin.setValue(self.adjustment_spin.value() + 0.01)
                self._apply_row_adjustment()
                return True
            elif key_event.key() == Qt.Key.Key_Up:
                self.adjustment_spin.setValue(self.adjustment_spin.value() + 0.1)
                self._apply_row_adjustment()
                return True
            elif key_event.key() == Qt.Key.Key_Down:
                self.adjustment_spin.setValue(self.adjustment_spin.value() - 0.1)
                self._apply_row_adjustment()
                return True

        return super().eventFilter(source, event)

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
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Tune Load Error", f"Could not load tune file:\n{exc}")
            return

        self.table_list.clear()
        if self.tune_data.tables:
            self._update_table_display()
        else:
            self.table_list.addItem("No tables found")
            self.table_grid.clear()
            self.table_grid.setRowCount(0)
            self.table_grid.setColumnCount(0)
            self.table_meta.setText("No tables found in this tune file")

        self.statusBar().showMessage(f"Loaded tune: {file_path.name}")
        self._refresh_workspace_text()
        self._add_recent_tune_file(file_path)

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
        
        # Show progress dialog for large files
        progress = QProgressBar()
        progress.setRange(0, 0)  # Indeterminate progress
        progress.setMinimumWidth(300)
        
        # Create a modal dialog with the progress bar
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel
        progress_dialog = QDialog(self)
        progress_dialog.setWindowTitle("Loading Log File")
        progress_dialog.setModal(True)
        progress_dialog.setFixedSize(350, 100)
        
        layout = QVBoxLayout(progress_dialog)
        layout.addWidget(QLabel(f"Loading {file_path.name}..."))
        layout.addWidget(progress)
        progress_dialog.show()
        
        # Force UI update
        QGuiApplication.processEvents()

        try:
            parse_result = self.log_loader.load_log_with_report(file_path)
            log_df = parse_result.dataframe
        except Exception as exc:  # noqa: BLE001
            progress_dialog.close()
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return

        progress_dialog.close()

        self.log_file = file_path
        self.statusBar().showMessage(
            f"Loaded log: {file_path.name} ({len(log_df):,} rows, parser={parse_result.parser_used}, enc={parse_result.encoding})"
        )
        self.log_df = log_df
        self._add_recent_log_file(file_path)

    def _refresh_workspace_text(self) -> None:
        """Workspace text refresh - disabled since workspace is hidden by default."""
        pass

    def _on_table_selected(self, current_item: QListWidgetItem | None, previous_item: QListWidgetItem | None) -> None:
        if not current_item or not self.tune_data:
            self.current_table = None
            self.current_x_axis = None
            self.current_y_axis = None
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

    def _refresh_current_table_view(self) -> None:
        if not self.current_table:
            return
        self._render_table(self.current_table, self.current_x_axis, self.current_y_axis)
        self._update_table_grid_row_controls()

    def _render_table(self, table: TableData, x_axis: AxisVector | None, y_axis: AxisVector | None) -> None:
        matrix, display_x_axis, display_y_axis = self._table_display_state(table, x_axis, y_axis)
        row_count = len(matrix)
        col_count = len(matrix[0]) if matrix else 0
        display_matrix = list(reversed(matrix))
        y_labels = list(reversed(self._header_labels(display_y_axis, row_count)))
        x_labels = self._header_labels(display_x_axis, col_count)

        self.table_grid.clear()
        self.table_grid.setRowCount(row_count + 1)
        self.table_grid.setColumnCount(col_count)
        self.table_grid.set_footer_rows(1)
        self.table_grid.setVerticalHeaderLabels(y_labels + ["X axis"])

        minimum, maximum = self._matrix_min_max(display_matrix)
        span = max(maximum - minimum, 1e-9)

        for r, row in enumerate(display_matrix):
            for c, value in enumerate(row):
                item = QTableWidgetItem(f"{value:g}")
                item.setBackground(self._cell_color(value, minimum, span))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_grid.setItem(r, c, item)

        axis_row = row_count
        for c, label in enumerate(x_labels):
            axis_item = QTableWidgetItem(label)
            axis_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            axis_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            axis_item.setBackground(QColor(60, 60, 60))
            self.table_grid.setItem(axis_row, c, axis_item)

        self.table_grid.resizeColumnsToContents()
        self.table_grid.resizeRowsToContents()
        self._auto_size_table_and_row_viz()

        units = table.units or "-"
        self.table_meta.setText(
            f"{table.name} | {row_count}x{col_count} | units: {units} | "
            f"x(bottom): {self._axis_title('X', display_x_axis)} | y(left): {self._axis_title('Y', display_y_axis)}"
        )

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
            min_editor_height = max(150, self.table_row_editor_group.sizeHint().height() + self.table_row_status.sizeHint().height() + 8)
            row_viz_height = max(220, details_height - min_editor_height)
            editor_height = max(min_editor_height, details_height - row_viz_height)
            self.table_details_splitter.setSizes([row_viz_height, editor_height])

    def _table_display_state(
        self,
        table: TableData,
        x_axis: AxisVector | None,
        y_axis: AxisVector | None,
    ) -> tuple[list[list[float]], AxisVector | None, AxisVector | None]:
        matrix = [row[:] for row in table.values]
        display_x_axis = x_axis
        display_y_axis = y_axis

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
        units = f" ({axis.units})" if axis.units else ""
        return f"{axis.name}{units}"

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

    def _populate_log_channel_controls(self) -> None:
        if self.log_df is None or self.log_df.empty:
            return

        numeric_columns = []
        for column in self.log_df.columns:
            try:
                series = self.log_df[column]
                numeric_series = self._to_numeric_series(series)
                if numeric_series is not None:
                    numeric_columns.append(str(column))
            except Exception:
                continue
        if not numeric_columns:
            self.scatter_status.setText(
                "Log loaded, but no numeric channels were detected for plotting."
            )
            return

        for combo in (self.x_channel_combo, self.y_channel_combo, self.correction_channel_combo):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(numeric_columns)
            combo.blockSignals(False)

        self._set_preferred_channel(self.x_channel_combo, ["rpm", "engine speed"])
        self._set_preferred_channel(self.y_channel_combo, ["map", "load"])
        self._set_preferred_channel(self.correction_channel_combo, ["ve", "ego", "corr", "trim"])

    def _set_preferred_channel(self, combo: QComboBox, keywords: list[str]) -> None:
        if combo.count() == 0:
            return
        lower_items = [combo.itemText(i).lower() for i in range(combo.count())]
        for keyword in keywords:
            for idx, text in enumerate(lower_items):
                if keyword in text:
                    combo.setCurrentIndex(idx)
                    return
        combo.setCurrentIndex(0)

    def _update_scatter_plot(self) -> None:
        if self.log_df is None or self.log_df.empty:
            self.scatter_status.setText("Load a log file to view scatter data.")
            return
        if self.x_channel_combo.count() == 0 or self.y_channel_combo.count() == 0 or self.correction_channel_combo.count() == 0:
            self.scatter_status.setText("Log loaded, but no numeric channels are currently selected.")
            return
        x_name = self.x_channel_combo.currentText()
        y_name = self.y_channel_combo.currentText()
        corr_name = self.correction_channel_combo.currentText()
        if not x_name or not y_name or not corr_name:
            return

        try:
            x_series = self._to_numeric_series(self.log_df[x_name])
            y_series = self._to_numeric_series(self.log_df[y_name])
            c_series = self._to_numeric_series(self.log_df[corr_name])
            if x_series is None or y_series is None or c_series is None:
                self.scatter_status.setText("Selected channels are not numeric enough to plot.")
                return
            import numpy as np

            x_vals = np.asarray(x_series, dtype=float)
            y_vals = np.asarray(y_series, dtype=float)
            c_vals = np.asarray(c_series, dtype=float)
            finite = np.isfinite(x_vals) & np.isfinite(y_vals) & np.isfinite(c_vals)
            x_vals = x_vals[finite]
            y_vals = y_vals[finite]
            c_vals = c_vals[finite]
            if len(x_vals) == 0:
                self.scatter_status.setText("No valid numeric rows found in selected channels.")
                return
        except Exception:
            self.scatter_status.setText("Could not parse selected channels as numeric values.")
            return

        self.all_points_item.setData(x_vals, y_vals)

        target = float(self.rpm_target_spin.value())
        tolerance = float(self.rpm_tolerance_spin.value())
        slice_mask = (x_vals >= target - tolerance) & (x_vals <= target + tolerance)
        slice_x = x_vals[slice_mask]
        slice_y = y_vals[slice_mask]
        slice_c = c_vals[slice_mask]

        if len(slice_x) == 0:
            self.slice_points_item.setData([], [])
            self.scatter_status.setText("No points in selected RPM slice.")
            return

        c_min = float(slice_c.min())
        c_max = float(slice_c.max())
        c_span = max(c_max - c_min, 1e-9)
        brushes = []
        for value in slice_c:
            ratio = float((value - c_min) / c_span)
            brushes.append(pg.mkBrush(self._correction_color(ratio)))
        self.slice_points_item.setData(slice_x, slice_y, brush=brushes)
        self.scatter_plot.getPlotItem().setLabel("bottom", x_name)
        self.scatter_plot.getPlotItem().setLabel("left", y_name)
        self.scatter_status.setText(
            f"Slice points: {len(slice_x)} | {corr_name} avg: {float(slice_c.mean()):.2f} | range: {c_min:.2f} to {c_max:.2f}"
        )

    def _apply_ve_slice_correction(self) -> None:
        if self.current_table is None or self.current_x_axis is None or self.current_y_axis is None:
            QMessageBox.information(self, "No Table", "Load and select a VE table first.")
            return
        if self.log_df is None or self.log_df.empty:
            QMessageBox.information(self, "No Log", "Load a log file first.")
            return

        table_name_lower = self.current_table.name.lower()
        if "ve" not in table_name_lower:
            response = QMessageBox.question(
                self,
                "Non-VE Table Selected",
                "Current table does not look like a VE table. Apply correction anyway?",
            )
            if response != QMessageBox.StandardButton.Yes:
                return

        x_name = self.x_channel_combo.currentText()
        y_name = self.y_channel_combo.currentText()
        corr_name = self.correction_channel_combo.currentText()
        if not x_name or not y_name or not corr_name:
            return

        import numpy as np

        try:
            x_series = self._to_numeric_series(self.log_df[x_name])
            y_series = self._to_numeric_series(self.log_df[y_name])
            corr_series = self._to_numeric_series(self.log_df[corr_name])
            if x_series is None or y_series is None or corr_series is None:
                QMessageBox.warning(self, "Channel Error", "Selected channels are not numeric enough.")
                return

            x_vals = np.asarray(x_series, dtype=float)
            y_vals = np.asarray(y_series, dtype=float)
            corr_vals = np.asarray(corr_series, dtype=float)
            finite = np.isfinite(x_vals) & np.isfinite(y_vals) & np.isfinite(corr_vals)
            x_vals = x_vals[finite]
            y_vals = y_vals[finite]
            corr_vals = corr_vals[finite]
        except Exception:
            QMessageBox.warning(self, "Channel Error", "Selected channels could not be converted to numbers.")
            return

        target_rpm = float(self.rpm_target_spin.value())
        tolerance = float(self.rpm_tolerance_spin.value())
        slice_mask = (x_vals >= target_rpm - tolerance) & (x_vals <= target_rpm + tolerance)
        x_slice = x_vals[slice_mask]
        y_slice = y_vals[slice_mask]
        corr_slice = corr_vals[slice_mask]
        if len(x_slice) == 0:
            QMessageBox.information(self, "No Data", "No log data points in selected RPM slice.")
            return

        rpm_bins = self.current_x_axis.values
        load_bins = self.current_y_axis.values
        if not rpm_bins or not load_bins:
            QMessageBox.warning(self, "Missing Axes", "Current table is missing axis vectors.")
            return

        col_idx = min(range(len(rpm_bins)), key=lambda i: abs(float(rpm_bins[i]) - target_rpm))
        blend = float(self.blend_spin.value()) / 100.0
        updates = 0
        for row_idx, load_bin in enumerate(load_bins):
            nearest_mask = self._nearest_bin_mask(y_slice, load_bins, row_idx)
            row_points = corr_slice[nearest_mask]
            if len(row_points) == 0:
                continue
            factor = float(row_points.mean()) / 100.0
            old_value = float(self.current_table.values[row_idx][col_idx])
            new_value = old_value * (1.0 + (factor - 1.0) * blend)
            self.current_table.values[row_idx][col_idx] = new_value
            updates += 1

        self._render_table(self.current_table, self.current_x_axis, self.current_y_axis)
        self.statusBar().showMessage(
            f"Applied VE correction at ~{target_rpm:.0f} RPM to {updates} load bins (column {col_idx + 1})."
        )

    @staticmethod
    def _nearest_bin_mask(values, bins: list[float], target_index: int):
        import numpy as np

        bin_arr = np.asarray(bins, dtype=float)
        distances = np.abs(values[:, None] - bin_arr[None, :])
        nearest_indices = np.argmin(distances, axis=1)
        return nearest_indices == target_index

    @staticmethod
    def _correction_color(ratio: float) -> QColor:
        ratio = max(0.0, min(1.0, ratio))
        if ratio < 0.5:
            local = ratio / 0.5
            return QColor(int(40 + 80 * local), int(120 + 80 * local), int(220 - 120 * local), 220)
        local = (ratio - 0.5) / 0.5
        return QColor(int(120 + 120 * local), int(200 - 120 * local), int(100 - 70 * local), 220)

    @staticmethod
    def _to_numeric_series(series):
        import pandas as pd

        direct = pd.to_numeric(series, errors="coerce")
        if direct.notna().mean() >= 0.5:
            return direct

        cleaned = (
            series.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(r"[^0-9eE+\-\.]", "", regex=True)
        )
        coerced = pd.to_numeric(cleaned, errors="coerce")
        if coerced.notna().mean() >= 0.5:
            return coerced
        return None

    @staticmethod
    def _default_data_dir() -> Path:
        cwd_data = Path.cwd() / "tuning_data"
        return cwd_data if cwd_data.exists() else Path.cwd()
