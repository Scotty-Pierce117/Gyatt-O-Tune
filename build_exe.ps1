$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
    throw "Python virtual environment not found at .venv. Create it first and install the project with 'pip install -e .[build]'."
}

Push-Location $projectRoot
try {
    & $pythonExe -m PyInstaller --noconfirm --clean "Gyatt-O-Tune.spec"
}
finally {
    Pop-Location
}