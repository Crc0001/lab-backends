@echo off
setlocal
title Lab Backends

set "BASE_DIR=%~dp0"

echo ============================================
echo   Lab Backends - Start All Services
echo ============================================
echo   8001 - chemistry-backend
echo   8002 - protein-backend
echo   8003 - microbial-backend
echo   8004 - gcms-backend
echo   8005 - hplc-backend
echo ============================================
echo.

call :start_backend "Chemistry-8001" "chemistry-backend" "8001"
call :start_backend "Protein-8002" "protein-backend" "8002"
call :start_backend "Microbial-8003" "microbial-backend" "8003"
call :start_backend "GCMS-8004" "gcms-backend" "8004"
call :start_backend "HPLC-8005" "hplc-backend" "8005"

echo.
echo All backend start commands were launched.
echo Close a service window to stop that service.
exit /b 0

:start_backend
set "WINDOW_TITLE=%~1"
set "BACKEND_NAME=%~2"
set "BACKEND_PORT=%~3"
set "BACKEND_DIR=%BASE_DIR%%BACKEND_NAME%"
set "PYTHON_EXE=%BACKEND_DIR%\venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

echo Starting %BACKEND_NAME% on port %BACKEND_PORT%...
start "%WINDOW_TITLE%" /D "%BACKEND_DIR%" cmd.exe /k ""%PYTHON_EXE%" -m uvicorn main:app --host 0.0.0.0 --port %BACKEND_PORT%"
exit /b 0
