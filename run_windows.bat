@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 未找到 Python。请先安装 Python 3.10 或更新版本，并勾选 Add Python to PATH。
  echo 下载地址: https://www.python.org/downloads/windows/
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 正在创建本地运行环境...
  python -m venv .venv
  if errorlevel 1 (
    echo [错误] 创建虚拟环境失败。
    pause
    exit /b 1
  )
)

echo 正在安装或检查依赖...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo [错误] 依赖安装失败。请检查网络连接或 pip 配置。
  pause
  exit /b 1
)

echo.
echo 工具启动中，请在浏览器打开:
echo http://127.0.0.1:8501
echo.
".venv\Scripts\streamlit.exe" run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless true
pause
