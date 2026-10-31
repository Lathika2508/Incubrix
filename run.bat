@echo off
setlocal
title IncuBrix Option 3 Visibility Engine Launcher

:: 1. Navigate to project root directory
cd /d "%~dp0"

echo ===============================================================================
echo                INCUBRIX OPTION 3: AEO, SEO AND GEO VISIBILITY ENGINE
echo ===============================================================================
echo Project Directory: %CD%
echo.

:: 2. Verify Python is available
echo [Step 1/4] Checking Python installation...
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.10+ and add it to PATH.
    pause
    exit /b 1
)
python --version
echo.

:: 3. Database initialization and verification via Python
echo [Step 2/4] Verifying database and dataset integrity...
python import_dataset.py
if errorlevel 1 (
    echo [ERROR] Database initialization failed.
    pause
    exit /b 1
)
echo.

:: 4. Launch FastAPI Backend
echo [Step 3/4] Checking FastAPI backend on port 8000...
netstat -ano | findstr :8000 | findstr LISTENING >nul 2>&1
if errorlevel 1 (
    echo Starting FastAPI backend on http://127.0.0.1:8000 ...
    start "IncuBrix FastAPI Backend" cmd /k "cd /d "%~dp0" && python -m uvicorn backend.app.main:app --host 127.0.0.1 --port 8000"
) else (
    echo [OK] Port 8000 is active.
)
echo.

:: 5. Launch Streamlit Dashboard
echo [Step 4/4] Starting Streamlit dashboard on http://localhost:8501 ...
echo ===============================================================================
echo IncuBrix Engine is running:
echo - Dashboard UI: http://localhost:8501
echo - API Docs:     http://localhost:8000/docs
echo ===============================================================================
echo.

netstat -ano | findstr :8501 | findstr LISTENING >nul 2>&1
if not errorlevel 1 (
    echo [OK] Streamlit is already active on port 8501.
    start http://localhost:8501
    exit /b 0
)

python -m streamlit run streamlit_app.py --server.port 8501 --server.address 127.0.0.1

pause
