# ============================================
#  EDTECH DOC TEMPLATER — Dependency Installer
#  Called by Inno Setup during installation
# ============================================
param(
    [string]$InstallDir = "C:\Program Files\EDTECH DOC TEMPLATER"
)

$ErrorActionPreference = "Continue"

function Write-Step($num, $total, $msg) {
    Write-Host "[$num/$total] $msg"
}

# 1. Chocolatey
Write-Step 1 6 "Verificando Chocolatey..."
if (-not (Test-Path "C:\ProgramData\chocolatey\bin\choco.exe")) {
    [System.Net.ServicePointManager]::SecurityProtocol = 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    $env:Path = "$env:Path;C:\ProgramData\chocolatey\bin"
}

# 2. Python
Write-Step 2 6 "Verificando Python..."
$hasPython = $false
try { python --version 2>&1 | Out-Null; $hasPython = $true } catch {}
if (-not $hasPython) {
    & choco install python311 -y --no-progress 2>&1 | Out-Null
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# 3. Tesseract
Write-Step 3 6 "Verificando Tesseract OCR..."
if (-not (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe")) {
    & choco install tesseract -y --no-progress 2>&1 | Out-Null
    $env:Path = "$env:Path;C:\Program Files\Tesseract-OCR"
}

# 4. Git + Clone
Write-Step 4 6 "Descargando EDTECH DOC TEMPLATER..."
$hasGit = $false
try { git --version 2>&1 | Out-Null; $hasGit = $true } catch {}
if (-not $hasGit) {
    & choco install git -y --no-progress 2>&1 | Out-Null
    $env:Path = "$env:Path;C:\Program Files\Git\bin"
}

Set-Location $InstallDir
if (-not (Test-Path "$InstallDir\app.py")) {
    git clone https://github.com/Israx1990BO/doc-templater.git . 2>&1 | Out-Null
}

# 5. Python venv + deps
Write-Step 5 6 "Instalando dependencias Python..."
if (-not (Test-Path "$InstallDir\venv")) {
    python -m venv "$InstallDir\venv"
}
& "$InstallDir\venv\Scripts\Activate.ps1"
python -m pip install --upgrade pip --quiet 2>&1 | Out-Null
pip install -r "$InstallDir\requirements.txt" --quiet 2>&1 | Out-Null

# 6. Working dirs
Write-Step 6 6 "Finalizando..."
New-Item -ItemType Directory -Path "$InstallDir\uploads" -Force | Out-Null
New-Item -ItemType Directory -Path "$InstallDir\outputs" -Force | Out-Null

Write-Host "Instalacion completada."
