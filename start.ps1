#Requires -Version 5.1
<#
.SYNOPSIS
    makeReportOllama - Auto Start Script (PowerShell 版)
.DESCRIPTION
    Node.js / Ollama / Python venv / npm の依存関係を確認・インストールし、
    バックエンド (FastAPI) とフロントエンド (Vite) を起動してブラウザを開く。
.NOTES
    初回実行時は管理者権限なしで動作するが、winget でのインストールには
    ネットワーク接続が必要。
    実行ポリシーが制限されている場合は以下を実行してから起動すること:
      Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── 設定 ────────────────────────────────────────────────────────
$ROOT               = $PSScriptRoot
$VENV               = Join-Path $ROOT '.venv'
$BACKEND_PORT       = 8000
$FRONTEND_PORT      = 5173
$BACKEND_URL        = "http://localhost:$BACKEND_PORT/health"
$FRONTEND_URL       = "http://localhost:$FRONTEND_PORT"
$OLLAMA_URL         = 'http://localhost:11434'
$OLLAMA_MODEL_ANALYST = 'qwen2.5:3b'
$OLLAMA_MODEL_WRITER  = 'qwen3:8b'
$OLLAMA_MODEL_EMBED   = 'nomic-embed-text'
$OLLAMA_DEFAULT_EXE = Join-Path $env:LOCALAPPDATA 'Programs\Ollama\ollama.exe'

# ── ヘルパー関数 ─────────────────────────────────────────────────

function Write-Step([string]$msg) {
    Write-Host "`n$msg" -ForegroundColor Cyan
}

function Write-Ok([string]$msg) {
    Write-Host $msg -ForegroundColor Green
}

function Write-Fail([string]$msg) {
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit 1
}

function Test-Url([string]$url, [int]$timeoutSec = 3) {
    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing `
                    -TimeoutSec $timeoutSec -ErrorAction Stop
        return $resp.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Wait-ForUrl([string]$url, [int]$maxTries = 15, [int]$intervalSec = 2) {
    for ($i = 1; $i -le $maxTries; $i++) {
        Start-Sleep -Seconds $intervalSec
        if (Test-Url $url) { return $true }
    }
    return $false
}

function Find-Ollama {
    # PATH 優先、次にデフォルトインストールパス
    $cmd = Get-Command ollama -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    if (Test-Path $OLLAMA_DEFAULT_EXE) { return $OLLAMA_DEFAULT_EXE }
    return $null
}

# ── [1/6] Node.js ────────────────────────────────────────────────
Write-Host '============================================================'
Write-Host '  makeReportOllama - Auto Start Script'
Write-Host '============================================================'

Write-Step '[1/6] Checking Node.js...'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host 'Node.js not found. Installing via winget...'
    winget install -e --id OpenJS.NodeJS --silent
    # winget は "already installed" でも非ゼロを返すことがあるため、直接チェック
    if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
        Write-Fail 'node not found after install. Please restart terminal and run start.ps1 again.'
    }
}
$nodeVer = node --version 2>$null
Write-Ok "Node.js $nodeVer OK."

# ── [2/6] Ollama ─────────────────────────────────────────────────
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

# Ollama サービス起動
if (-not (Test-Url $OLLAMA_URL)) {
    Write-Host 'Starting Ollama service...'
    Start-Process -FilePath $ollamaExe -ArgumentList 'serve' `
                  -WindowStyle Minimized
    Write-Host 'Waiting for Ollama to start...'
    if (-not (Wait-ForUrl $OLLAMA_URL)) {
        Write-Fail 'Ollama startup timed out.'
    }
    Write-Ok 'Ollama service started.'
} else {
    Write-Ok 'Ollama service already running.'
}

# モデルのチェック / pull
foreach ($model in @($OLLAMA_MODEL_ANALYST, $OLLAMA_MODEL_WRITER, $OLLAMA_MODEL_EMBED)) {
    Write-Host "Checking model '$model'..."
    $listed = & $ollamaExe list 2>$null | Select-String -Pattern ([regex]::Escape($model)) -Quiet
    if (-not $listed) {
        Write-Host "Model '$model' not found. Pulling now - this may take a while..."
        & $ollamaExe pull $model
        if ($LASTEXITCODE -ne 0) {
            Write-Fail "Failed to pull model '$model'."
        }
        Write-Ok "Model '$model' ready."
    } else {
        Write-Ok "Model '$model' OK."
    }
}

# ── [3/6] Python venv ────────────────────────────────────────────
Write-Step '[3/6] Setting up Python venv...'
$venvActivate = Join-Path $VENV 'Scripts\activate.ps1'
if (-not (Test-Path $venvActivate)) {
    Write-Host 'Creating venv...'
    python -m venv $VENV
    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Failed to create venv. Check Python 3.9+ is installed.'
    }
}
Write-Host 'Installing pip packages...'
$pip = Join-Path $VENV 'Scripts\pip.exe'
& $pip install -q -r (Join-Path $ROOT 'backend\requirements.txt')
if ($LASTEXITCODE -ne 0) {
    Write-Fail 'pip install failed.'
}
Write-Ok 'Python venv ready.'

# ── [4/6] Frontend npm install ───────────────────────────────────
Write-Step '[4/6] Checking frontend dependencies...'
$nodeModules = Join-Path $ROOT 'frontend\node_modules'
if (-not (Test-Path $nodeModules)) {
    Write-Host 'Running npm install...'
    Push-Location (Join-Path $ROOT 'frontend')
    npm install --silent
    $npmExit = $LASTEXITCODE
    Pop-Location
    if ($npmExit -ne 0) {
        Write-Fail 'npm install failed.'
    }
}
Write-Ok 'Frontend dependencies OK.'

# ── [5/6] Start backend ──────────────────────────────────────────
Write-Step '[5/6] Starting backend server...'

# 既存のバックエンドプロセスを停止
Get-Process -Name python -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like '*makeReportOllama-Backend*' } |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$pythonExe  = Join-Path $VENV 'Scripts\python.exe'
$backendDir = Join-Path $ROOT 'backend'
$backendArgs = "-m uvicorn main:app --host 0.0.0.0 --port $BACKEND_PORT"

Start-Process -FilePath $pythonExe `
              -ArgumentList $backendArgs `
              -WorkingDirectory $backendDir `
              -WindowStyle Minimized

Write-Host 'Waiting for backend to start...'
if (-not (Wait-ForUrl $BACKEND_URL)) {
    Write-Fail "Backend startup timed out. Check $ROOT\backend\app.log"
}
Write-Ok 'Backend started.'

# ── [6/6] Start frontend ─────────────────────────────────────────
Write-Step '[6/6] Checking frontend server...'
if (-not (Test-Url $FRONTEND_URL)) {
    Write-Host "Starting frontend on port $FRONTEND_PORT..."
    Start-Process -FilePath 'cmd.exe' `
                  -ArgumentList "/c npm run dev" `
                  -WorkingDirectory (Join-Path $ROOT 'frontend') `
                  -WindowStyle Minimized
    Start-Sleep -Seconds 4
    Write-Ok 'Frontend started.'
} else {
    Write-Ok 'Frontend already running.'
}

# ── ブラウザを開く ───────────────────────────────────────────────
Write-Host ''
Write-Host '============================================================' -ForegroundColor Green
Write-Host "  Ready! Opening browser at $FRONTEND_URL"              -ForegroundColor Green
Write-Host '============================================================' -ForegroundColor Green
Start-Sleep -Seconds 2
Start-Process $FRONTEND_URL

Write-Host ''
Write-Host 'Servers are running in background windows.'
Write-Host 'To stop: close the Backend/Frontend/Ollama windows or use Task Manager.'
Read-Host 'Press Enter to exit'
