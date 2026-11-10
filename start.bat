@echo off
chcp 65001 >nul
title Word AI 创作助手 - 本地开发调试服务

echo ==========================================================
echo       Word AI 创作助手 (Office Web Add-in) 一键启动
echo ==========================================================
echo.

:: 检查 Python 环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到系统安装 Python，请先安装 Python 3.8+ 并勾选 Add to PATH。
    pause
    exit /b
)

echo [*] 正在启动本地 HTTPS 静态服务器 (https://localhost:3000)...
echo [*] 服务启动后将自动调起默认浏览器进行本地功能测试...
echo.

python server.py

pause
