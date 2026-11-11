@echo off
title Word AI Copilot Add-in Installer
setlocal

echo ==========================================================
echo        Word AI Copilot Add-in - Installer ^& Cache Cleaner
echo ==========================================================
echo.

set "MANIFEST=%~dp0manifest.xml"
set "ID=d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
set "CATALOG_GUID={d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
set "SHARE_NAME=WordAIAssistant"

if not exist "%MANIFEST%" (
    echo [ERROR] manifest.xml not found in current folder!
    pause
    exit /b 1
)

echo [*] Closing running Word process if open...
taskkill /f /im winword.exe >nul 2>&1

echo [1/5] Cleaning Office WEF Cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
    rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache" (
    rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache" >nul 2>&1
)

echo [2/5] Registering Developer Sideload Manifest in Registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ID%" /t REG_SZ /d "%MANIFEST%" /f >nul 2>&1

echo [3/5] Configuring Trusted Catalogs...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "\\%COMPUTERNAME%\%SHARE_NAME%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul 2>&1

echo [4/5] Sharing Local Folder as Network Share...
net share %SHARE_NAME%="%~dp0." /grant:everyone,READ >nul 2>&1

echo [5/5] Checking Local Certificate...
if exist "%~dp0certs\localhost.crt" (
    certutil -addstore -user -f "Root" "%~dp0certs\localhost.crt" >nul 2>&1
    echo       ^> Local certificate trusted
) else (
    echo       ^> Cloud-hosted mode (no local certificate required)
)

echo.
echo ==========================================================
echo        Installation and Cache Refresh Finished!
echo ==========================================================
echo.
echo [Usage]:
echo 1. Open Microsoft Word.
echo 2. Find [AI Copilot] on the far right of the Home tab.
echo    Or go to [Insert] -^> [My Add-ins] -^> [Shared Folder].
echo 3. To reload changes anytime, right-click inside the add-in pane
echo    and choose 'Reload'.
echo.
pause
