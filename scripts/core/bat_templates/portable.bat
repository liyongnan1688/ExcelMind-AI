@echo off
chcp 65001 >nul
title ExcelMind AI - 免安装便携启动
echo 正在检测系统架构并启动 Excel，请稍候...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0core\launch_portable.ps1"

timeout /t 3 >nul
