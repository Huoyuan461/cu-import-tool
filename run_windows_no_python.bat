@echo off
setlocal
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel_mapper_windows.ps1"
if errorlevel 1 (
  echo.
  echo 程序运行失败。请确认这台电脑已安装 Microsoft Excel，且公司策略允许运行 PowerShell。
  pause
)
