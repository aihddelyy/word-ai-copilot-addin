# ==========================================================================
# Word AI 创作助手 - Windows 一键在线安装脚本 (使用 GitHub 镜像加速下载)
# 快速在线运行 (PowerShell 一行命令):
# irm https://ghfast.top/https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main/online-install.ps1 | iex
# ==========================================================================
param(
    [string]$App = "",
    [string]$Mirror = "auto"
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(65001) } catch {}
$Host.UI.RawUI.WindowTitle = "Word AI 创作助手 - GitHub 镜像极速在线安装"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   Word AI 创作助手 - Windows 在线极速安装程序" -ForegroundColor Cyan
Write-Host "   (支持 Microsoft Word / WPS 文字，内置 GitHub 镜像加速)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# 1. GitHub 镜像加速源配置
$mirrorSources = @(
    @{ Name = "GHFast 镜像加速站 (推荐)"; BaseUrl = "https://ghfast.top/https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main"; Probe = "https://ghfast.top/https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main/manifest.xml" },
    @{ Name = "GH-Proxy 镜像加速站"; BaseUrl = "https://gh-proxy.com/https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main"; Probe = "https://gh-proxy.com/https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main/manifest.xml" },
    @{ Name = "Cloudflare Pages 镜像备用源"; BaseUrl = "https://word-ai-copilot-addin.pages.dev"; Probe = "https://word-ai-copilot-addin.pages.dev/assets/icon-32.png" },
    @{ Name = "GitHub 官方直连源"; BaseUrl = "https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main"; Probe = "https://raw.githubusercontent.com/aihddelyy/word-ai-copilot-addin/main/manifest.xml" }
)

function Test-MirrorLatency {
    param([string]$Url, [int]$TimeoutSec = 3)
    try {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Method = "HEAD"
        $req.Timeout = $TimeoutSec * 1000
        $req.AllowAutoRedirect = $true
        $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WordAICopilot/1.0"
        $resp = $req.GetResponse()
        $status = [int]$resp.StatusCode
        $resp.Close()
        $sw.Stop()
        if ($status -ge 200 -and $status -lt 400) {
            return $sw.ElapsedMilliseconds
        }
    }
    catch {
        return -1
    }
    return -1
}

Write-Host "[1/4] 正在探测最优 GitHub 镜像加速下载节点..." -ForegroundColor Gray
$selectedBaseUrl = ""
$selectedName = ""

foreach ($m in $mirrorSources) {
    Write-Host "  -> 探测节点: $($m.Name)..." -NoNewline -ForegroundColor Gray
    $latency = Test-MirrorLatency -Url $m.Probe -TimeoutSec 3
    if ($latency -gt 0) {
        Write-Host " [√ 畅通: ${latency}ms]" -ForegroundColor Green
        if ($selectedBaseUrl -eq "") {
            $selectedBaseUrl = $m.BaseUrl
            $selectedName = $m.Name
            break
        }
    }
    else {
        Write-Host " [超时/跳过]" -ForegroundColor Yellow
    }
}

if ($selectedBaseUrl -eq "") {
    Write-Host "  -> 未测出最优节点，默认选用 GHFast 镜像站" -ForegroundColor Yellow
    $selectedBaseUrl = $mirrorSources[0].BaseUrl
    $selectedName = $mirrorSources[0].Name
}

Write-Host "  => 选定下载镜像: $selectedName" -ForegroundColor Green
Write-Host ""

# 2. 准备本地安装目录
$installDir = Join-Path $env:LOCALAPPDATA "WordAICopilot"
Write-Host "[2/4] 初始化本地安装目录: $installDir" -ForegroundColor Gray
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

$filesToDownload = @(
    "manifest.xml",
    "install.bat",
    "install.ps1",
    "uninstall.bat",
    "uninstall.ps1",
    "assets/icon-16.png",
    "assets/icon-32.png",
    "assets/icon-64.png",
    "assets/icon-80.png",
    "wps-addon/manifest.json",
    "wps-addon/ribbon.xml",
    "wps-addon/index.html",
    "wps-addon/main.js",
    "wps-addon/icon-16.png",
    "wps-addon/icon-32.png"
)

function Download-RemoteFile {
    param([string]$Url, [string]$Dest)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.AllowAutoRedirect = $true
    $req.Timeout = 10000
    $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WordAICopilot/1.0"
    $resp = $req.GetResponse()
    $stream = $resp.GetResponseStream()
    $fs = [System.IO.File]::Create($Dest)
    $stream.CopyTo($fs)
    $fs.Close()
    $stream.Close()
    $resp.Close()
}

# 3. 通过 GitHub 镜像站下载全部文件
Write-Host "[3/4] 正在通过 GitHub 镜像高速下载插件核心组件..." -ForegroundColor Gray

foreach ($file in $filesToDownload) {
    $remoteUrl = "$selectedBaseUrl/$file"
    $localFile = Join-Path $installDir $file
    $localFileDir = Split-Path -Parent $localFile
    if (-not (Test-Path $localFileDir)) {
        New-Item -ItemType Directory -Path $localFileDir -Force | Out-Null
    }

    $downloadOk = $false
    try {
        Download-RemoteFile -Url $remoteUrl -Dest $localFile
        $downloadOk = $true
    }
    catch {
        # 尝试第二备用镜像源
        try {
            $fallbackUrl = "$($mirrorSources[1].BaseUrl)/$file"
            Download-RemoteFile -Url $fallbackUrl -Dest $localFile
            $downloadOk = $true
        }
        catch {
            $downloadOk = $false
        }
    }

    if ($downloadOk) {
        Write-Host "  [OK] $file" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [失败] $file" -ForegroundColor Red
    }
}

# 确保脚本文件为 UTF-8 带 BOM，彻底杜绝 Windows PowerShell 5.1 下的中文解析乱码
$scriptFiles = @("install.ps1", "uninstall.ps1", "install.bat", "uninstall.bat")
foreach ($sf in $scriptFiles) {
    $fullSf = Join-Path $installDir $sf
    if (Test-Path $fullSf) {
        try {
            $rawBytes = [System.IO.File]::ReadAllBytes($fullSf)
            $text = [System.Text.Encoding]::UTF8.GetString($rawBytes)
            [System.IO.File]::WriteAllText($fullSf, $text, [System.Text.Encoding]::UTF8)
        }
        catch {}
    }
}

Write-Host "  -> 所有文件下载校验完成！" -ForegroundColor Green
Write-Host ""

# 4. 调用本地安装逻辑注册到 Word / WPS
Write-Host "[4/4] 启动安装注册向导..." -ForegroundColor Gray
$localInstallPs1 = Join-Path $installDir "install.ps1"

if (Test-Path $localInstallPs1) {
    if ($App -ne "") {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $localInstallPs1 $App
    }
    else {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $localInstallPs1
    }
}
else {
    Write-Host "[错误] 未找到安装脚本: $localInstallPs1" -ForegroundColor Red
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "               一键在线安装全流程执行完毕！" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "插件文件已就绪于: $installDir" -ForegroundColor Cyan
Write-Host "后续如需重新配置或卸载，可随时进入该目录运行 install.bat 或 uninstall.bat。" -ForegroundColor Gray
Write-Host ""
