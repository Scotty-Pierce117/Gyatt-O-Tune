$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
    throw "Python virtual environment not found at .venv. Create it first and install the project with 'pip install -e .[build]'."
}

# -- Step 1: Build portable EXE -----------------------------------------------
Push-Location $projectRoot
try {
    Write-Host "==> Building portable EXE with PyInstaller..." -ForegroundColor Cyan
    & $pythonExe -m PyInstaller --noconfirm --clean "Gyatt-O-Tune.spec"
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed (exit code $LASTEXITCODE)." }
    Write-Host "==> Portable EXE ready: dist\Gyatt-O-Tune.exe" -ForegroundColor Green
}
finally {
    Pop-Location
}

# -- Step 2: Build installer with Inno Setup (optional) -----------------------
$isccPaths = @(
    "iscc",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 5\ISCC.exe"
)

$iscc = $null
foreach ($candidate in $isccPaths) {
    if (Get-Command $candidate -ErrorAction SilentlyContinue) {
        $iscc = $candidate
        break
    }
}

if ($null -eq $iscc) {
    Write-Host ""
    Write-Host "==> Inno Setup not found -- skipping installer build." -ForegroundColor Yellow
    Write-Host "    Install Inno Setup 6 from: https://jrsoftware.org/isdl.php" -ForegroundColor Yellow
    Write-Host "    Then re-run this script, or run manually: iscc installer.iss" -ForegroundColor Yellow
} else {
    Push-Location $projectRoot
    try {
        Write-Host ""
        Write-Host "==> Building Windows installer with Inno Setup ($iscc)..." -ForegroundColor Cyan
        & $iscc "installer.iss"
        if ($LASTEXITCODE -ne 0) { throw "Inno Setup compiler failed (exit code $LASTEXITCODE)." }
        Write-Host "==> Installer ready: dist\Gyatt-O-Tune-Setup-0.1.3.exe" -ForegroundColor Green
    }
    finally {
        Pop-Location
    }
}
