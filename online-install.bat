@echo off
chcp 65001 >nul
title Word AI 创作助手 - Windows 一键在线安装
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0online-install.ps1" %*
pause
