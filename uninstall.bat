@echo off
chcp 65001 >nul
title Word AI 创作助手 - 统一卸载程序
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
pause
