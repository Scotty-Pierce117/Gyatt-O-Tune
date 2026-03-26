from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QSettings, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QDialog

from gyatt_o_tune.ui.main_window import MainWindow, StartupTuneDialog


def _asset_path(filename: str) -> Path:
    if getattr(sys, "frozen", False):
        base_path = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        return base_path / "gyatt_o_tune" / "assets" / filename
    return Path(__file__).resolve().parent / "assets" / filename


def _resolve_window_icon_path() -> Path | None:
    for candidate in ("gyatt-o-tune.ico", "gyatt-o-tune.svg"):
        path = _asset_path(candidate)
        if path.exists():
            return path
    return None


def _load_recent_tune_paths() -> list[Path]:
    settings = QSettings("GyattOTune", "GyattOTune")
    recent_tunes = settings.value("recent_tune_files", [])
    if isinstance(recent_tunes, str):
        recent_tunes = [recent_tunes]
    if not isinstance(recent_tunes, list):
        return []
    return [Path(p) for p in recent_tunes if isinstance(p, str) and Path(p).exists()]


def _default_browse_dir(recent_tunes: list[Path]) -> Path:
    if recent_tunes:
        return recent_tunes[0].parent
    tuning_data_dir = Path.cwd() / "tuning_data"
    if tuning_data_dir.exists():
        return tuning_data_dir
    return Path.cwd()


def _open_startup_tune_and_layout(window: MainWindow, selected_tune: Path) -> None:
    if window._open_recent_tune_file(selected_tune):
        window._load_default_window_layout()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Gyatt-O-Tune")

    icon_path = _resolve_window_icon_path()
    if icon_path is not None:
        app_icon = QIcon(str(icon_path))
        app.setWindowIcon(app_icon)

    selected_tune: Path | None = None
    if len(sys.argv) > 1:
        tune_path = Path(sys.argv[1])
        if tune_path.is_file():
            selected_tune = tune_path

    if selected_tune is None:
        recent_tunes = _load_recent_tune_paths()
        startup_dialog = StartupTuneDialog(
            recent_tunes,
            app.windowIcon(),
            default_browse_dir=_default_browse_dir(recent_tunes),
        )
        if startup_dialog.exec() != QDialog.DialogCode.Accepted or startup_dialog.selected_tune_path is None:
            return 0
        selected_tune = startup_dialog.selected_tune_path

    window = MainWindow()
    if icon_path is not None:
        window.setWindowIcon(QIcon(str(icon_path)))
    window.show()

    if selected_tune is not None:
        QTimer.singleShot(0, lambda: _open_startup_tune_and_layout(window, selected_tune))

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
