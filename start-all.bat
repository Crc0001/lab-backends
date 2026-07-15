@echo off
chcp 65001 >nul
title Lab Backends - 启动中...

echo ============================================
echo   Lab Backends - 批量启动脚本
echo ============================================
echo.
echo   8001 - chemistry-backend  (化学检测)
echo   8002 - protein-backend    (蛋白)
echo   8003 - microbial-backend  (微生物)
echo   8004 - gcms-backend       (气质)
echo   8005 - hplc-backend       (液相)
echo.
echo   即将在5个独立窗口中启动...
echo ============================================

set BASE_DIR=%~dp0

:: 化学检测 - 8001
start "Chemistry-8001" cmd /k "cd /d "%BASE_DIR%chemistry-backend" && venv\Scripts\activate && uvicorn main:app --host 0.0.0.0 --port 8001"
echo [OK] 化学检测 (8001) 启动中...

:: 蛋白 - 8002
start "Protein-8002" cmd /k "cd /d "%BASE_DIR%protein-backend" && venv\Scripts\activate && uvicorn main:app --host 0.0.0.0 --port 8002"
echo [OK] 蛋白 (8002) 启动中...

:: 微生物 - 8003
start "Microbial-8003" cmd /k "cd /d "%BASE_DIR%microbial-backend" && venv\Scripts\activate && uvicorn main:app --host 0.0.0.0 --port 8003"
echo [OK] 微生物 (8003) 启动中...

:: 气质 - 8004
start "GCMS-8004" cmd /k "cd /d "%BASE_DIR%gcms-backend" && venv\Scripts\activate && uvicorn main:app --host 0.0.0.0 --port 8004"
echo [OK] 气质 (8004) 启动中...

:: 液相 - 8005
start "HPLC-8005" cmd /k "cd /d "%BASE_DIR%hplc-backend" && venv\Scripts\activate && uvicorn main:app --host 0.0.0.0 --port 8005"
echo [OK] 液相 (8005) 启动中...

echo.
echo ============================================
echo   全部启动完毕！共 5 个系统
echo ============================================
echo.
echo   关闭本窗口不会影响已启动的服务
echo   关闭对应的 cmd 窗口即可停止对应服务
echo.
pause
