@echo off
chcp 65001 >nul
title ExcelMind AI - 一键卸载
echo 正在启动 ExcelMind AI 插件卸载向导，请稍候...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0core\uninstall_addin.ps1"

echo.
pause
