from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from gyatt_o_tune.ui.main_window import MainWindow


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


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Gyatt-O-Tune")

    icon_path = _resolve_window_icon_path()
    if icon_path is not None:
        app_icon = QIcon(str(icon_path))
        app.setWindowIcon(app_icon)

    window = MainWindow()
    if icon_path is not None:
        window.setWindowIcon(QIcon(str(icon_path)))
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
