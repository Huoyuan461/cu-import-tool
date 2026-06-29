@echo off
setlocal
cd /d "%~dp0"

if not exist "%SystemRoot%\System32\mshta.exe" (
  echo [错误] 找不到 Windows HTML 应用运行器 mshta.exe。
  pause
  exit /b 1
)

start "" "%SystemRoot%\System32\mshta.exe" "%~dp0excel_mapper_html.hta"
