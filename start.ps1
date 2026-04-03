#Requires -Version 5.1
# makeReportOllama - Auto Start Script (PowerShell)
# Usage: .\start.ps1
# If blocked by execution policy, run first:
#   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

$ROOT          = $PSScriptRoot
$VENV          = Join-Path $ROOT '.venv'
$BACKEND_PORT  = 8000
$FRONTEND_PORT = 5173
$BACKEND_URL   = "http://localhost:$BACKEND_PORT/health"
$FRONTEND_URL  = "http://localhost:$FRONTEND_PORT"
$OLLAMA_URL    = 'http://localhost:11434'
$OLLAMA_MODEL_ANALYST = 'qwen2.5:3b'
$OLLAMA_MODEL_WRITER  = 'qwen3:8b'
$OLLAMA_MODEL_EMBED   = 'nomic-embed-text'
$OLLAMA_DEFAULT_EXE   = Join-Path $env:LOCALAPPDATA 'Programs\Ollama\ollama.exe'

# Helper functions
function Write-Step { param([string]$msg); Write-Host "`n$msg" -ForegroundColor Cyan }
function Write-Ok   { param([string]$msg); Write-Host $msg -ForegroundColor Green }
function Write-Fail {
    param([string]$msg)
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit 1
}

function Test-Url {
    param([string]$url, [int]$timeoutSec = 3)
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $timeoutSec -ErrorAction Stop
        return $r.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Wait-ForUrl {
    param([string]$url, [int]$maxTries = 15, [int]$intervalSec = 2)
    for ($i = 1; $i -le $maxTries; $i++) {
        Start-Sleep -Seconds $intervalSec
        if (Test-Url $url) { return $true }
    }
    return $false
}

function Find-Ollama {
    $cmd = Get-Command ollama -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    if (Test-Path $OLLAMA_DEFAULT_EXE) { return $OLLAMA_DEFAULT_EXE }
    return $null
}

# ============================================================
Write-Host '============================================================'
Write-Host '  makeReportOllama - Auto Start Script'
Write-Host '============================================================'

# [1/6] Node.js
Write-Step '[1/6] Checking Node.js...'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host 'Node.js not found. Installing via winget...'
    winget install -e --id OpenJS.NodeJS --silent
    if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
        Write-Fail 'node not found after install. Please restart terminal and run start.ps1 again.'
    }
}
$nodeVer = node --version 2>$null
Write-Ok "Node.js $nodeVer OK."

# [2/6] Ollama
Write-Step '[2/6] Checking Ollama...'

$ollamaExe = Find-Ollama
if (-not $ollamaExe) {
    Write-Host 'Ollama not found. Installing via winget...'
    winget install -e --id Ollama.Ollama --silent
    Start-Sleep -Seconds 3
    $ollamaExe = Find-Ollama
    if (-not $ollamaExe) {
        Write-Fail 'Ollama not found after install. Please restart terminal and run start.ps1 again.'
    }
}
Write-Ok "Ollama found: $ollamaExe"

# Start Ollama service if not running
$ollamaRunning = Test-Url $OLLAMA_URL
if ($ollamaRunning) {
    Write-Ok 'Ollama service already running.'
} else {
    Write-Host 'Starting Ollama service...'
    $op = @{ FilePath = $ollamaExe; ArgumentList = 'serve'; WindowStyle = 'Minimized' }
    Start-Process @op
    Write-Host 'Waiting for Ollama to start...'
    if (-not (Wait-ForUrl $OLLAMA_URL)) {
        Write-Fail 'Ollama startup timed out.'
    }
    Write-Ok 'Ollama service started.'
}

# Check / pull models
foreach ($model in @($OLLAMA_MODEL_ANALYST, $OLLAMA_MODEL_WRITER, $OLLAMA_MODEL_EMBED)) {
    Write-Host "Checking model '$model'..."
    $listed = & $ollamaExe list 2>$null | Select-String -Pattern ([regex]::Escape($model)) -Quiet
    if ($listed) {
        Write-Ok "Model '$model' OK."
    } else {
        Write-Host "Model '$model' not found. Pulling now (may take a while)..."
        & $ollamaExe pull $model
        if ($LASTEXITCODE -ne 0) { Write-Fail "Failed to pull model '$model'." }
        Write-Ok "Model '$model' ready."
    }
}

# [3/6] Python venv
Write-Step '[3/6] Setting up Python venv...'
$venvActivate = Join-Path $VENV 'Scripts\activate.ps1'
if (-not (Test-Path $venvActivate)) {
    Write-Host 'Creating venv...'
    python -m venv $VENV
    if ($LASTEXITCODE -ne 0) { Write-Fail 'Failed to create venv. Check Python 3.9+ is installed.' }
}
Write-Host 'Installing pip packages...'
$pip = Join-Path $VENV 'Scripts\pip.exe'
& $pip install -q -r (Join-Path $ROOT 'backend\requirements.txt')
if ($LASTEXITCODE -ne 0) { Write-Fail 'pip install failed.' }
Write-Ok 'Python venv ready.'

# [4/6] Frontend npm install
Write-Step '[4/6] Checking frontend dependencies...'
$nodeModules = Join-Path $ROOT 'frontend\node_modules'
if (-not (Test-Path $nodeModules)) {
    Write-Host 'Running npm install...'
    Push-Location (Join-Path $ROOT 'frontend')
    npm install --silent
    $npmExit = $LASTEXITCODE
    Pop-Location
    if ($npmExit -ne 0) { Write-Fail 'npm install failed.' }
}
Write-Ok 'Frontend dependencies OK.'

# [5/6] Start backend
Write-Step '[5/6] Starting backend server...'

# Detect LAN IP and set CORS_ORIGINS so other PCs can access
$lanIp = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
         Where-Object { $_.IPAddress -notmatch '^(127\.|169\.)' } |
         Select-Object -First 1 -ExpandProperty IPAddress
if ($lanIp) {
    $env:CORS_ORIGINS = "http://${lanIp}:$FRONTEND_PORT"
    Write-Host "LAN IP: $lanIp" -ForegroundColor Yellow
    Write-Host "Other PCs can access: http://${lanIp}:$FRONTEND_PORT" -ForegroundColor Yellow
} else {
    $env:CORS_ORIGINS = ''
}

# Stop existing backend process
Get-Process -Name python -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like '*makeReportOllama-Backend*' } |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$pythonExe  = Join-Path $VENV 'Scripts\python.exe'
$backendDir = Join-Path $ROOT 'backend'

$psi = [System.Diagnostics.ProcessStartInfo]::new()
$psi.FileName         = $pythonExe
$psi.Arguments        = "-m uvicorn main:app --host 0.0.0.0 --port $BACKEND_PORT"
$psi.WorkingDirectory = $backendDir
$psi.WindowStyle      = [System.Diagnostics.ProcessWindowStyle]::Minimized
$psi.UseShellExecute  = $false
$psi.EnvironmentVariables['CORS_ORIGINS'] = $env:CORS_ORIGINS
[System.Diagnostics.Process]::Start($psi) | Out-Null

Write-Host 'Waiting for backend to start...'
if (-not (Wait-ForUrl $BACKEND_URL)) {
    Write-Fail "Backend startup timed out. Check $ROOT\backend\app.log"
}
Write-Ok 'Backend started.'

# [6/6] Start frontend
Write-Step '[6/6] Checking frontend server...'
$frontendRunning = Test-Url $FRONTEND_URL
if ($frontendRunning) {
    Write-Ok 'Frontend already running.'
} else {
    Write-Host "Starting frontend on port $FRONTEND_PORT..."
    $fp = @{
        FilePath         = 'cmd.exe'
        ArgumentList     = '/c npm run dev'
        WorkingDirectory = (Join-Path $ROOT 'frontend')
        WindowStyle      = 'Minimized'
    }
    Start-Process @fp
    Start-Sleep -Seconds 4
    Write-Ok 'Frontend started.'
}

# Open browser
Write-Host ''
Write-Host '============================================================' -ForegroundColor Green
Write-Host "  Ready! Opening browser at $FRONTEND_URL"                   -ForegroundColor Green
if ($lanIp) {
    Write-Host "  From other PCs:  http://${lanIp}:$FRONTEND_PORT"       -ForegroundColor Yellow
}
Write-Host '============================================================' -ForegroundColor Green
Start-Sleep -Seconds 2
Start-Process $FRONTEND_URL

Write-Host ''
Write-Host 'Servers are running in background windows.'
Write-Host 'To stop: close the Backend/Frontend/Ollama windows or use Task Manager.'
Read-Host 'Press Enter to exit'
