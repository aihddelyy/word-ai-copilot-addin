@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo        Word AI Copilot Add-in - 一键信任本地 SSL 证书
echo ==========================================================
echo.
echo 正在将本地证书导入 Windows 受信任证书库...
echo (若弹出 Windows 安全警告框提示"是否安装该证书"，请务必点击【是(Y)】)...
echo.
certutil -addstore -user -f "Root" "%~dp0certs\localhost.crt"
echo.
if %errorlevel% equ 0 (
    echo [成功] 本地证书已成功受信任！
    echo 现在回到 Word 侧边栏点击【重新启动】，即可正常显示！
) else (
    echo [提示] 导入未完成。
)
echo.
pause
