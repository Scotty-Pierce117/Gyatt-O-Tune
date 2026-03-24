from __future__ import annotations

from pathlib import Path
from typing import Any

import pyqtgraph as pg
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from scattersyncms.core.io import LogLoader


class LogViewerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("ScatterSyncMS - Log Viewer")
        self.resize(1200, 800)

        self.log_loader = LogLoader()
        self.log_df: Any | None = None

        root = QWidget(self)
        layout = QVBoxLayout(root)

        top = QHBoxLayout()
        self.open_btn = QPushButton("Open Log File...")
        self.open_btn.clicked.connect(self._open_log_file)
        top.addWidget(self.open_btn)
        self.file_label = QLabel("No log loaded")
        top.addWidget(self.file_label, 1)
        layout.addLayout(top)

        controls = QHBoxLayout()
        self.x_combo = QComboBox()
        self.y_combo = QComboBox()
        self.x_combo.currentTextChanged.connect(self._render_scatter)
        self.y_combo.currentTextChanged.connect(self._render_scatter)
        controls.addWidget(QLabel("X"))
        controls.addWidget(self.x_combo, 1)
        controls.addWidget(QLabel("Y"))
        controls.addWidget(self.y_combo, 1)
        layout.addLayout(controls)

        mode_row = QHBoxLayout()
        self.single_mode_radio = QRadioButton("Single Y + Color")
        self.multi_mode_radio = QRadioButton("Multiple Y Channels (checkboxes)")
        self.single_mode_radio.setChecked(True)
        self.plot_mode_group = QButtonGroup(self)
        self.plot_mode_group.addButton(self.single_mode_radio)
        self.plot_mode_group.addButton(self.multi_mode_radio)
        self.single_mode_radio.toggled.connect(self._render_scatter)
        self.multi_mode_radio.toggled.connect(self._render_scatter)
        mode_row.addWidget(self.single_mode_radio)
        mode_row.addWidget(self.multi_mode_radio)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        self.multi_y_list = QListWidget()
        self.multi_y_list.itemChanged.connect(self._render_scatter)
        self.multi_y_list.setMaximumHeight(140)
        layout.addWidget(QLabel("Multi-Y Channels"))
        layout.addWidget(self.multi_y_list)

        self.scatter_plot = pg.PlotWidget()
        self.scatter_plot.showGrid(x=True, y=True, alpha=0.2)
        self.scatter_item = pg.ScatterPlotItem(size=5, pen=pg.mkPen(220, 220, 220, 120))
        self.scatter_plot.addItem(self.scatter_item)
        self.multi_scatter_items: list[pg.ScatterPlotItem] = []
        layout.addWidget(self.scatter_plot, 2)

        self.status_label = QLabel("Open a .mlg/.msl/.csv log to inspect parsing and channels.")
        layout.addWidget(self.status_label)

        self.diagnostics = QTextEdit()
        self.diagnostics.setReadOnly(True)
        layout.addWidget(self.diagnostics, 1)

        self.setCentralWidget(root)

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

        try:
            result = self.log_loader.load_log_with_report(file_path)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Log Load Error", f"Could not load log file:\n{exc}")
            return

        self.log_df = result.dataframe
        self.file_label.setText(file_path.name)
        self.status_label.setText(f"Rows: {len(self.log_df):,} | Columns: {len(self.log_df.columns):,}")
        self.diagnostics.setPlainText(
            f"Parser: {result.parser_used}\n"
            f"Encoding: {result.encoding}\n"
            f"Notes: {result.notes or '-'}\n\n"
            f"Columns:\n- " + "\n- ".join(map(str, self.log_df.columns[:120]))
        )

        self._populate_channel_controls()
        self._render_scatter()

    def _populate_channel_controls(self) -> None:
        if self.log_df is None:
            return
        numeric_columns: list[str] = []
        for column in self.log_df.columns:
            series = self._to_numeric_series(self.log_df[column])
            if series is not None:
                numeric_columns.append(str(column))
        for combo in (self.x_combo, self.y_combo):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(numeric_columns)
            combo.blockSignals(False)

        self.multi_y_list.blockSignals(True)
        self.multi_y_list.clear()
        for channel in numeric_columns:
            item = QListWidgetItem(channel)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.multi_y_list.addItem(item)
        self.multi_y_list.blockSignals(False)

        self._set_preferred(self.x_combo, ["rpm", "engine speed", "time"])
        self._set_preferred(self.y_combo, ["map", "load", "afr"])
        self._precheck_multi_y(["map", "load", "afr", "ve1", "ego cor1"])

    def _render_scatter(self) -> None:
        if self.log_df is None:
            return

        for item in self.multi_scatter_items:
            self.scatter_plot.removeItem(item)
        self.multi_scatter_items.clear()

        if self.multi_mode_radio.isChecked():
            self._render_multi_channel_scatter()
            return

        x_name = self.x_combo.currentText()
        y_name = self.y_combo.currentText()
        if not x_name or not y_name:
            return

        x_series = self._to_numeric_series(self.log_df[x_name])
        y_series = self._to_numeric_series(self.log_df[y_name])
        if x_series is None or y_series is None:
            self.status_label.setText("Selected channels are not numeric enough to plot.")
            return

        import numpy as np

        x = np.asarray(x_series, dtype=float)
        y = np.asarray(y_series, dtype=float)
        finite = np.isfinite(x) & np.isfinite(y)
        x = x[finite]
        y = y[finite]
        if len(x) == 0:
            self.status_label.setText("No valid numeric rows for selected channels.")
            return

        y_min = float(y.min())
        y_max = float(y.max())
        span = max(y_max - y_min, 1e-9)
        brushes = [pg.mkBrush(self._color_from_ratio((float(v) - y_min) / span)) for v in y]
        self.scatter_item.setData(x, y, brush=brushes)
        self.scatter_plot.getPlotItem().setLabel("bottom", x_name)
        self.scatter_plot.getPlotItem().setLabel("left", y_name)
        self._set_plot_bounds([x], [y])
        self.status_label.setText(
            f"Rows: {len(x):,} | Auto-color by {y_name}: min {y_min:.2f}, max {y_max:.2f}"
        )

    def _render_multi_channel_scatter(self) -> None:
        if self.log_df is None:
            return
        x_name = self.x_combo.currentText()
        if not x_name:
            return
        selected_y = self._checked_multi_y_channels()
        if not selected_y:
            self.scatter_item.setData([], [])
            self.status_label.setText("Select one or more Y channels using checkboxes.")
            return

        x_series = self._to_numeric_series(self.log_df[x_name])
        if x_series is None:
            self.status_label.setText("X channel is not numeric enough.")
            return

        import numpy as np

        x = np.asarray(x_series, dtype=float)
        series_count = 0
        total_points = 0
        color_table = [
            (60, 170, 255, 180),
            (255, 170, 70, 180),
            (140, 220, 120, 180),
            (230, 120, 120, 180),
            (200, 120, 255, 180),
            (120, 220, 220, 180),
        ]
        x_sets: list[Any] = []
        y_sets: list[Any] = []
        for idx, y_name in enumerate(selected_y):
            y_series = self._to_numeric_series(self.log_df[y_name])
            if y_series is None:
                continue
            y = np.asarray(y_series, dtype=float)
            finite = np.isfinite(x) & np.isfinite(y)
            x_plot = x[finite]
            y_plot = y[finite]
            if len(x_plot) == 0:
                continue
            y_min = float(np.nanmin(y_plot))
            y_max = float(np.nanmax(y_plot))
            y_span = max(y_max - y_min, 1e-9)
            brushes = [pg.mkBrush(self._color_from_ratio((float(v) - y_min) / y_span)) for v in y_plot]
            scatter = pg.ScatterPlotItem(
                size=4,
                pen=pg.mkPen(color_table[idx % len(color_table)]),
                name=y_name,
            )
            scatter.setData(x_plot, y_plot, brush=brushes)
            self.scatter_plot.addItem(scatter)
            self.multi_scatter_items.append(scatter)
            x_sets.append(x_plot)
            y_sets.append(y_plot)
            series_count += 1
            total_points += len(x_plot)

        self.scatter_item.setData([], [])
        self.scatter_plot.getPlotItem().setLabel("bottom", x_name)
        self.scatter_plot.getPlotItem().setLabel("left", "Selected Y channels")
        self._set_plot_bounds(x_sets, y_sets)
        self.status_label.setText(
            f"Multi-Y mode: {series_count} channel(s), {total_points:,} plotted points."
        )

    def _set_plot_bounds(self, x_sets: list[Any], y_sets: list[Any]) -> None:
        import numpy as np

        valid_x = [np.asarray(values, dtype=float) for values in x_sets if len(values) > 0]
        valid_y = [np.asarray(values, dtype=float) for values in y_sets if len(values) > 0]
        if not valid_x or not valid_y:
            self.scatter_plot.enableAutoRange()
            return

        x_all = np.concatenate(valid_x)
        y_all = np.concatenate(valid_y)
        if len(x_all) == 0 or len(y_all) == 0:
            self.scatter_plot.enableAutoRange()
            return

        x_min = float(np.nanmin(x_all))
        x_max = float(np.nanmax(x_all))
        y_min = float(np.nanmin(y_all))
        y_max = float(np.nanmax(y_all))

        x_pad = max((x_max - x_min) * 0.05, 1e-6)
        y_pad = max((y_max - y_min) * 0.05, 1e-6)
        self.scatter_plot.setXRange(x_min - x_pad, x_max + x_pad, padding=0.0)
        self.scatter_plot.setYRange(y_min - y_pad, y_max + y_pad, padding=0.0)

    @staticmethod
    def _to_numeric_series(series: Any) -> Any | None:
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
    def _set_preferred(combo: QComboBox, keywords: list[str]) -> None:
        if combo.count() == 0:
            return
        lowered = [combo.itemText(i).lower() for i in range(combo.count())]
        for keyword in keywords:
            for idx, text in enumerate(lowered):
                if keyword in text:
                    combo.setCurrentIndex(idx)
                    return
        combo.setCurrentIndex(0)

    def _precheck_multi_y(self, keywords: list[str]) -> None:
        lowered = [self.multi_y_list.item(i).text().lower() for i in range(self.multi_y_list.count())]
        matched = 0
        for keyword in keywords:
            for idx, text in enumerate(lowered):
                if keyword in text:
                    self.multi_y_list.item(idx).setCheckState(Qt.CheckState.Checked)
                    matched += 1
                    break
        if matched == 0 and self.multi_y_list.count() > 0:
            self.multi_y_list.item(0).setCheckState(Qt.CheckState.Checked)

    def _checked_multi_y_channels(self) -> list[str]:
        selected: list[str] = []
        for idx in range(self.multi_y_list.count()):
            item = self.multi_y_list.item(idx)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(item.text())
        return selected

    @staticmethod
    def _color_from_ratio(ratio: float):
        ratio = max(0.0, min(1.0, ratio))
        if ratio < 0.5:
            local = ratio / 0.5
            return (int(30 + 40 * local), int(110 + 100 * local), int(220 - 140 * local), 220)
        local = (ratio - 0.5) / 0.5
        return (int(70 + 170 * local), int(210 - 120 * local), int(80 - 50 * local), 220)

    @staticmethod
    def _default_data_dir() -> Path:
        cwd_data = Path.cwd() / "tuning_data"
        return cwd_data if cwd_data.exists() else Path.cwd()

