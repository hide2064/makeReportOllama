@echo off
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"
set "BACKEND_PORT=8000"
set "FRONTEND_PORT=5173"
set "BACKEND_URL=http://localhost:%BACKEND_PORT%/health"
set "FRONTEND_URL=http://localhost:%FRONTEND_PORT%"

echo ============================================================
echo  makeReportOllama - Auto Start Script
echo ============================================================
echo.

:: [1/5] Check Node.js
echo [1/5] Checking Node.js...
where node >nul 2>&1
if errorlevel 1 (
    echo Node.js not found. Installing via winget...
    winget install -e --id OpenJS.NodeJS --silent
    if errorlevel 1 (
        echo [ERROR] Node.js installation failed. Please install manually.
        pause
        exit /b 1
    )
    call refreshenv >nul 2>&1
    where node >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] node not found after install. Please restart terminal.
        pause
        exit /b 1
    )
)
for /f "tokens=*" %%v in ('node --version 2^>nul') do set "NODE_VER=%%v"
echo Node.js %NODE_VER% OK.

:: [2/5] Python venv
echo.
echo [2/5] Setting up Python venv...
if not exist "%VENV%\Scripts\activate.bat" (
    echo Creating venv...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] Failed to create venv. Check Python 3.9+ is installed.
        pause
        exit /b 1
    )
)
echo Installing pip packages...
"%VENV%\Scripts\pip" install -q -r "%ROOT%backend\requirements.txt"
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)
echo Python venv ready.

:: [3/5] Frontend npm install
echo.
echo [3/5] Checking frontend dependencies...
if not exist "%ROOT%frontend\node_modules" (
    echo Running npm install...
    pushd "%ROOT%frontend"
    call npm install --silent
    if errorlevel 1 (
        echo [ERROR] npm install failed.
        popd
        pause
        exit /b 1
    )
    popd
)
echo Frontend dependencies OK.

:: [4/5] Start backend
echo.
echo [4/5] Checking backend server...
curl -s -o nul -w "%%{http_code}" "%BACKEND_URL%" 2>nul | findstr "200" >nul
if errorlevel 1 (
    echo Starting backend on port %BACKEND_PORT%...
    start "makeReportOllama-Backend" /min cmd /c "cd /d "%ROOT%backend" && "%VENV%\Scripts\python" -m uvicorn main:app --host 0.0.0.0 --port %BACKEND_PORT%"
    set /a "tries=0"
    :wait_backend
    timeout /t 2 /nobreak >nul
    curl -s -o nul -w "%%{http_code}" "%BACKEND_URL%" 2>nul | findstr "200" >nul
    if not errorlevel 1 goto backend_ready
    set /a "tries+=1"
    if !tries! lss 15 goto wait_backend
    echo [ERROR] Backend startup timed out. Check backend\app.log.
    pause
    exit /b 1
    :backend_ready
    echo Backend started.
) else (
    echo Backend already running.
)

:: [5/5] Start frontend
echo.
echo [5/5] Checking frontend server...
curl -s -o nul "%FRONTEND_URL%" 2>nul
if errorlevel 1 (
    echo Starting frontend on port %FRONTEND_PORT%...
    start "makeReportOllama-Frontend" /min cmd /c "cd /d "%ROOT%frontend" && npm run dev"
    timeout /t 4 /nobreak >nul
    echo Frontend started.
) else (
    echo Frontend already running.
)

:: Open browser
echo.
echo ============================================================
echo  Ready! Opening browser at %FRONTEND_URL%
echo ============================================================
timeout /t 2 /nobreak >nul
start "" "%FRONTEND_URL%"

echo.
echo Servers are running in background windows.
echo To stop: close the Backend/Frontend windows or use Task Manager.
pause
endlocal
