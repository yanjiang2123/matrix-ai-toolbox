@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Matrix AI 数据工具箱
echo ============================================

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] 首次运行，正在创建独立环境...
  goto :create_venv
)
goto :venv_ready

:create_venv
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m venv .venv
  if errorlevel 1 goto :failed
  goto :venv_ready
)
where python >nul 2>nul
if errorlevel 1 goto :python_missing
python -m venv .venv
if errorlevel 1 goto :python_missing

:venv_ready

echo [2/3] 检查依赖...
".venv\Scripts\python.exe" -c "import flask, openpyxl, fpdf, docx" >nul 2>nul
if errorlevel 1 (
  echo 正在安装项目依赖，请保持网络连接...
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :failed
)

echo [3/3] 正在启动，本窗口关闭后程序退出...
".venv\Scripts\python.exe" app.py
exit /b %errorlevel%

:python_missing
echo 未找到可用的 Python。请先安装 Python 3.10 或更高版本，并勾选 Add Python to PATH。
pause
exit /b 1

:failed
echo.
echo 启动准备失败，请查看上方提示。
pause
exit /b 1
