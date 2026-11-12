@echo off
chcp 65001 >nul
title Word AI 创作助手 - 统一安装程序
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
pause
