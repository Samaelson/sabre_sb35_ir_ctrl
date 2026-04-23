# IRDroid Installer Script
# Run as Administrator

Write-Host "===================================" -ForegroundColor Cyan
Write-Host " IRDroid Installation Starting..." -ForegroundColor Cyan
Write-Host "==================================="

# -----------------------------
# CONFIG
# -----------------------------
$ServiceName = "IRDroidService"
$PythonCheck = "python"
$ScriptDir = "C:\Scripts\IRDroid"
$PSPath = "$ScriptDir\irdroid.ps1"

# -----------------------------
# CREATE FOLDER STRUCTURE
# -----------------------------
Write-Host "Creating directories..." -ForegroundColor Yellow

if (!(Test-Path $ScriptDir)) {
    New-Item -ItemType Directory -Path $ScriptDir -Force | Out-Null
}

# -----------------------------
# CHECK PYTHON
# -----------------------------
Write-Host "Checking Python..." -ForegroundColor Yellow

try {
    $pythonVersion = python --version
    Write-Host "Found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Python not found! Please install Python 3.9+" -ForegroundColor Red
    exit 1
}

# -----------------------------
# INSTALL DEPENDENCIES
# -----------------------------
Write-Host "Installing Python dependencies..." -ForegroundColor Yellow

pip install pywin32 pyserial

# -----------------------------
# REGISTER WINDOWS SERVICE
# -----------------------------
Write-Host "Installing Windows Service..." -ForegroundColor Yellow

try {
    python service.py install
    python service.py start
    Write-Host "Service installed and started." -ForegroundColor Green
} catch {
    Write-Host "Service installation failed." -ForegroundColor Red
}

# -----------------------------
# OPTIONAL: COPY FILES
# -----------------------------
Write-Host "Copying scripts..." -ForegroundColor Yellow

# You can adapt these paths to your repo structure
if (Test-Path ".\irdroid.ps1") {
    Copy-Item ".\irdroid.ps1" $PSPath -Force
}

if (Test-Path ".\service.py") {
    Copy-Item ".\service.py" $ScriptDir -Force
}

# -----------------------------
# TEST PIPE CONNECTION
# -----------------------------
Write-Host "Testing setup..." -ForegroundColor Yellow

Start-Sleep -Seconds 2

try {
    $pipe = New-Object System.IO.Pipes.NamedPipeClientStream(".", "irdroid", "InOut")
    $pipe.Connect(2000)
    Write-Host "Pipe connection OK." -ForegroundColor Green
    $pipe.Close()
} catch {
    Write-Host "Pipe not available yet (service may still start)." -ForegroundColor DarkYellow
}

# -----------------------------
# DONE
# -----------------------------
Write-Host ""
Write-Host "===================================" -ForegroundColor Cyan
Write-Host " IRDroid Installation Complete" -ForegroundColor Cyan
Write-Host "==================================="

Write-Host ""
Write-Host "Usage examples:"
Write-Host "  .\irdroid.ps1 power"
Write-Host "  .\irdroid.ps1 hdmi1 bluetooth"
Write-Host ""