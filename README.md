# Gyatt-O-Tune

Desktop app for analyzing MegaSquirt tune files and log files:

- Load a tune file and one or more log files
- Generate scatter plots from log channels
- Build derived lookup tables using log data + tune tables
- Support workflows for VE, spark, knock, and additional table types

## Quick start

1. Create a virtual environment:
   - Windows PowerShell: `python -m venv .venv`
2. Activate it:
   - `.\.venv\Scripts\Activate.ps1`
3. Install the project:
   - `pip install -e .`
4. Run the app:
   - `python -m scattersyncms.main`
   - or `gyatt-o-tune`

## Build a portable Windows `.exe`

This project includes a PyInstaller build script that creates a single-file Windows executable.

1. Install the app and build dependency:
   - `pip install -e .[build]`
2. Build the executable:
   - `powershell -ExecutionPolicy Bypass -File .\build_exe.ps1`
3. Find the output here:
   - `dist\Gyatt-O-Tune.exe`

Notes:

- The build is configured for `--onefile` and `--windowed` so the app runs as a standalone GUI executable.
- The packaged build includes the app SVG asset used for the window icon.
- Windows may still show SmartScreen warnings on unsigned executables. That is separate from portability.

## Initial roadmap

- Tune file import abstraction (MSQ support)
- Log import abstraction (CSV / MSL support)
- Channel mapping + data cleaning pipeline
- Scatter plot and selection tools
- Table alignment, interpolation, and export back to tune
