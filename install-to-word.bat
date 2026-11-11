@echo off
chcp 65001 >nul
title Word AI Copilot 插件一键安装与修复工具
setlocal

echo ==========================================================
echo        Word AI Copilot Add-in - 一键安装与修复
echo ==========================================================
echo.

set "MANIFEST=%~dp0manifest.xml"
set "ID=d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
set "CATALOG_GUID={d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
set "SHARE_NAME=WordAIAssistant"

:: 检查 manifest.xml 是否存在
if not exist "%MANIFEST%" (
    echo [错误] 未在当前目录下找到 manifest.xml！
    echo 请确认当前脚本位于项目根目录。
    pause
    exit /b 1
)

:: 检查 Word 是否正在运行，如果在运行则提示关闭或自动关闭
tasklist /fi "imagename eq winword.exe" 2>nul | find /i "winword.exe" >nul
if %errorlevel% equ 0 (
    echo [提示] 检测到 Microsoft Word 正在运行。
    echo 为确保 Office 插件缓存清理生效，正在关闭 Word...
    taskkill /f /im winword.exe >nul 2>&1
    timeout /t 1 >nul
)

echo [1/5] 正在深度清理 Office WEF 插件与 WebView 缓存...
:: 清理 Office 16.0 WEF 缓存目录
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
    rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)
:: 清理可能存在的 Edge WebView2 Office 缓存
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache" (
    rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache" >nul 2>&1
)

echo [2/5] 正在向 Windows 注册表挂载 Developer 旁加载清单...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ID%" /t REG_SZ /d "%MANIFEST%" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo       ^> 开发者清单挂载成功
) else (
    echo       ^> 警告: 写入 Developer 注册表失败，请检查权限
)

echo [3/5] 正在配置 Office 受信任外接程序目录 (Trusted Catalogs)...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "\\%COMPUTERNAME%\%SHARE_NAME%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul 2>&1

echo [4/5] 正在创建本地网络共享目录映射 (作为受信任共享双保险)...
net share %SHARE_NAME%="%~dp0." /grant:everyone,READ >nul 2>&1
if %errorlevel% neq 0 (
    :: 若非管理员权限导致失败，尝试提权静默执行
    powershell -NoProfile -Command "Start-Process cmd -ArgumentList '/c net share %SHARE_NAME%=""%~dp0."" /grant:everyone,READ' -Verb RunAs -Wait -WindowStyle Hidden" >nul 2>&1
)

echo [5/5] 检查本地开发证书 (若存在本地证书则自动信任)...
if exist "%~dp0certs\localhost.crt" (
    certutil -addstore -user -f "Root" "%~dp0certs\localhost.crt" >nul 2>&1
    echo       ^> 已自动信任本地证书
) else (
    echo       ^> 云端托管模式 (无需本地证书)
)

echo.
echo ==========================================================
echo                   安装与刷新完成！
echo ==========================================================
echo.
echo 【如何使用】：
echo 1. 打开 Microsoft Word
echo 2. 在上方功能区【主页 / 开始】标签最右侧，即可看到【AI 创作助手】按钮
echo 3. 或者在【插入】->【我的外接程序】->【共享文件夹】中点击加载
echo.
echo (若界面有任何更新未显示，右键点击插件侧边栏空白处选择"重新加载"即可)
echo.
pause
