@echo off
chcp 65001 >nul
title Lee-Excel AI 插件 - 一键安装
echo 正在启动 Lee-Excel 插件安装向导，请稍候...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0core\install_addin.ps1"

echo.
pause
