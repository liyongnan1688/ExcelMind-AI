@echo off
chcp 65001 >nul
title ExcelMind AI - 一键安装
echo 正在启动 ExcelMind AI 插件安装向导，请稍候...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0core\install_addin.ps1"

echo.
pause
