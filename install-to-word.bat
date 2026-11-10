@echo off
title Word AI Copilot Installer
echo ==========================================================
echo        Word AI Copilot Add-in - Installer
echo ==========================================================
echo.

set "MANIFEST=%~dp0manifest.xml"
set "ID=d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"

echo [1/5] Cleaning Office WEF Cache...
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1

echo [2/5] Registering Add-in Manifest...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ID%" /t REG_SZ /d "%MANIFEST%" /f

echo [3/5] Adding Trusted Catalog...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{%ID%}" /v "Id" /t REG_SZ /d "{%ID%}" /f
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{%ID%}" /v "Url" /t REG_SZ /d "\\%COMPUTERNAME%\WordAIAssistant" /f
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{%ID%}" /v "Flags" /t REG_DWORD /d 1 /f

echo [4/5] Trusting Local SSL Certificate (if exists)...
if exist "%~dp0certs\localhost.crt" (
    certutil -addstore -user -f "Root" "%~dp0certs\localhost.crt" >nul 2>&1
)

echo [5/5] Sharing Folder (Run as Admin if needed)...
net share WordAIAssistant="%~dp0." /grant:everyone,READ

echo.
echo ==========================================================
echo        Installation Process Finished!
echo ==========================================================
echo.
pause
