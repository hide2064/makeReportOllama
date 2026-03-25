@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"
set "BACKEND_PORT=8000"
set "FRONTEND_PORT=5173"
set "BACKEND_URL=http://localhost:%BACKEND_PORT%/health"
set "FRONTEND_URL=http://localhost:%FRONTEND_PORT%"

echo ============================================================
echo  makeReportOllama - 自動起動スクリプト
echo ============================================================
echo.

:: ── 1. Node.js 確認 / インストール ─────────────────────────
echo [1/5] Node.js を確認中...
where node >nul 2>&1
if errorlevel 1 (
    echo Node.js が見つかりません。winget でインストールします...
    winget install -e --id OpenJS.NodeJS --silent
    if errorlevel 1 (
        echo [ERROR] Node.js のインストールに失敗しました。手動でインストールしてください。
        pause & exit /b 1
    )
    :: PATH を再読み込み
    call refreshenv >nul 2>&1
    where node >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] インストール後も node が見つかりません。ターミナルを再起動してください。
        pause & exit /b 1
    )
)
for /f "tokens=*" %%v in ('node --version 2^>nul') do set "NODE_VER=%%v"
echo Node.js %NODE_VER% を確認しました。

:: ── 2. Python venv 構築・有効化・依存インストール ─────────
echo.
echo [2/5] Python 仮想環境を準備中...
if not exist "%VENV%\Scripts\activate.bat" (
    echo venv を作成します...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] venv の作成に失敗しました。Python 3.9+ がインストールされているか確認してください。
        pause & exit /b 1
    )
)
echo pip パッケージをインストール中...
"%VENV%\Scripts\pip" install -q -r "%ROOT%backend\requirements.txt"
if errorlevel 1 (
    echo [ERROR] pip install に失敗しました。
    pause & exit /b 1
)
echo Python 環境の準備完了。

:: ── 3. フロントエンド npm install ─────────────────────────
echo.
echo [3/5] フロントエンド依存パッケージを確認中...
if not exist "%ROOT%frontend\node_modules" (
    echo npm install を実行します...
    pushd "%ROOT%frontend"
    call npm install --silent
    if errorlevel 1 (
        echo [ERROR] npm install に失敗しました。
        popd & pause & exit /b 1
    )
    popd
)
echo フロントエンド依存パッケージ OK。

:: ── 4. バックエンドサーバー起動（未起動時のみ）────────────
echo.
echo [4/5] バックエンドサーバーを確認・起動中...
curl -s -o nul -w "%%{http_code}" "%BACKEND_URL%" 2>nul | findstr "200" >nul
if errorlevel 1 (
    echo バックエンドを起動します (ポート %BACKEND_PORT%)...
    start "makeReportOllama-Backend" /min cmd /c ^
        "cd /d "%ROOT%backend" && "%VENV%\Scripts\python" -m uvicorn main:app --host 0.0.0.0 --port %BACKEND_PORT% 2>> "%ROOT%backend\app.log""
    :: 起動待機（最大30秒）
    set /a "tries=0"
    :wait_backend
    timeout /t 2 /nobreak >nul
    curl -s -o nul -w "%%{http_code}" "%BACKEND_URL%" 2>nul | findstr "200" >nul
    if not errorlevel 1 goto backend_ready
    set /a "tries+=1"
    if !tries! lss 15 goto wait_backend
    echo [ERROR] バックエンドの起動がタイムアウトしました。app.log を確認してください。
    pause & exit /b 1
    :backend_ready
    echo バックエンド起動完了。
) else (
    echo バックエンドはすでに起動済みです。
)

:: ── 5. フロントエンドサーバー起動（未起動時のみ）──────────
echo.
echo [5/5] フロントエンドサーバーを確認・起動中...
curl -s -o nul "%FRONTEND_URL%" 2>nul
if errorlevel 1 (
    echo フロントエンドを起動します (ポート %FRONTEND_PORT%)...
    start "makeReportOllama-Frontend" /min cmd /c ^
        "cd /d "%ROOT%frontend" && npm run dev"
    timeout /t 4 /nobreak >nul
    echo フロントエンド起動完了。
) else (
    echo フロントエンドはすでに起動済みです。
)

:: ── ブラウザで UI を開く ──────────────────────────────────
echo.
echo ============================================================
echo  起動完了！ブラウザで UI を開きます...
echo  URL: %FRONTEND_URL%
echo ============================================================
timeout /t 2 /nobreak >nul
start "" "%FRONTEND_URL%"

echo.
echo このウィンドウを閉じてもサーバーは動作し続けます。
echo サーバーを停止するにはタスクマネージャーからプロセスを終了してください。
pause
endlocal
