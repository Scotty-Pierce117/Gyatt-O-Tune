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

## Initial roadmap

- Tune file import abstraction (MSQ support)
- Log import abstraction (CSV / MSL support)
- Channel mapping + data cleaning pipeline
- Scatter plot and selection tools
- Table alignment, interpolation, and export back to tune
