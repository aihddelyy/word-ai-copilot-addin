@echo off
setlocal
cd /d "%~dp0"

echo ==========================================================
echo        Word AI Copilot Add-in - 一键卸载工具
echo ==========================================================
echo.

set "ADDIN_ID=d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
set "CATALOG_GUID={d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
set "SHARE_NAME=WordAIAssistant"

echo [1/3] 正在从注册表清理开发者记录...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1

echo [2/3] 正在清理受信任目录配置...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>&1

echo [3/3] 正在删除本地网络共享...
net share %SHARE_NAME% /delete /y >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "Start-Process cmd -ArgumentList '/c net share %SHARE_NAME% /delete /y' -Verb RunAs -Wait -WindowStyle Hidden" >nul 2>&1
)

echo.
echo ==========================================================
echo                 卸载完成！
echo ==========================================================
echo 重新启动 Microsoft Word 后生效。
echo.
pause
