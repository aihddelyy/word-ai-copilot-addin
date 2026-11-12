# Word AI 创作助手 - 智能线路测速与一键安装脚本
param(
    [string]$Route = "auto"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Word AI 创作助手 - 智能线路安装"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "       Word AI Copilot Add-in - 智能线路选择与安装" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ManifestPath = Join-Path $ScriptDir "manifest.xml"
$Id = "d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
$CatalogGuid = "{d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
$ShareName = "WordAIAssistant"

if (-not (Test-Path $ManifestPath)) {
    Write-Host "[错误] 未在当前目录下找到 manifest.xml 清单文件！" -ForegroundColor Red
    pause
    exit 1
}

# 1. 确定目标线路
$TargetRoute = ""
if ($Route -match "^(cf|cloudflare)$") {
    $TargetRoute = "cloudflare"
    Write-Host "[*] 已手动指定生效线路: [Cloudflare Pages 极速加速线路]" -ForegroundColor Yellow
} elseif ($Route -match "^(gh|github)$") {
    $TargetRoute = "github"
    Write-Host "[*] 已手动指定生效线路: [GitHub Pages 官方线路]" -ForegroundColor Yellow
} else {
    Write-Host "[*] 正在检测主选线路 (GitHub Pages) 连通性..." -ForegroundColor Gray
    $ghOk = $false
    try {
        $req = [System.Net.WebRequest]::Create("https://aihddelyy.github.io/word-ai-copilot-addin/src/taskpane.html")
        $req.Timeout = 2500
        $res = $req.GetResponse()
        if ($res.StatusCode -eq 200) { $ghOk = $true }
        $res.Close()
    } catch {
        $ghOk = $false
    }

    if ($ghOk) {
        Write-Host "[√] GitHub Pages 连接畅通，优先采用 GitHub Pages 线路！" -ForegroundColor Green
        $TargetRoute = "github"
    } else {
        Write-Host "[!] GitHub Pages 访问较慢或不可达，正在测速 Cloudflare Pages 备用极速线路..." -ForegroundColor Yellow
        $cfOk = $false
        try {
            $req2 = [System.Net.WebRequest]::Create("https://word-ai-copilot-addin.pages.dev/assets/icon-32.png")
            $req2.Timeout = 2500
            $res2 = $req2.GetResponse()
            if ($res2.StatusCode -eq 200) { $cfOk = $true }
            $res2.Close()
        } catch {
            $cfOk = $false
        }

        if ($cfOk) {
            Write-Host "[√] Cloudflare Pages 连接极速畅通，已自动为您选用 Cloudflare 线路！" -ForegroundColor Green
            $TargetRoute = "cloudflare"
        } else {
            Write-Host "[!] 两条线路均受网络波动影响，默认选用 Cloudflare Pages 备用线路。" -ForegroundColor Yellow
            $TargetRoute = "cloudflare"
        }
    }
}

# 2. 动态更新 manifest.xml
Write-Host ""
Write-Host "[*] 正在为 manifest.xml 动态应用线路配置: [$TargetRoute]..." -ForegroundColor Gray
$content = [System.IO.File]::ReadAllText($ManifestPath, [System.Text.Encoding]::UTF8)
if ($TargetRoute -eq "cloudflare") {
    $content = $content -replace "https://aihddelyy\.github\.io/word-ai-copilot-addin", "https://word-ai-copilot-addin.pages.dev"
    Write-Host "    -> 目标基准地址: https://word-ai-copilot-addin.pages.dev" -ForegroundColor Cyan
} else {
    $content = $content -replace "https://word-ai-copilot-addin\.pages\.dev", "https://aihddelyy.github.io/word-ai-copilot-addin"
    Write-Host "    -> 目标基准地址: https://aihddelyy.github.io/word-ai-copilot-addin" -ForegroundColor Cyan
}
[System.IO.File]::WriteAllText($ManifestPath, $content, [System.Text.Encoding]::UTF8)

# 3. 关闭运行中的 Word 进程
Write-Host ""
Write-Host "[*] 正在关闭已打开的 Word 进程以刷新插件缓存..." -ForegroundColor Gray
Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue

# 4. 清理 Office 插件缓存目录
Write-Host "[1/4] 清理 Office 插件缓存目录 (Wef & WebServiceCache)..." -ForegroundColor Gray
$wefPath = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
$webCachePath = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\WebServiceCache"
if (Test-Path $wefPath) { Remove-Item -Path $wefPath -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $webCachePath) { Remove-Item -Path $webCachePath -Recurse -Force -ErrorAction SilentlyContinue }

# 5. 注册开发者旁加载清单到注册表
Write-Host "[2/4] 注册开发者旁加载清单到注册表..." -ForegroundColor Gray
$devRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $devRegPath)) { New-Item -Path $devRegPath -Force | Out-Null }
Set-ItemProperty -Path $devRegPath -Name $Id -Value $ManifestPath -Force | Out-Null

# 6. 配置 Office 信任目录 (Trusted Catalogs)
Write-Host "[3/4] 配置 Office 信任目录 (Trusted Catalogs)..." -ForegroundColor Gray
$catRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"
if (-not (Test-Path $catRegPath)) { New-Item -Path $catRegPath -Force | Out-Null }
Set-ItemProperty -Path $catRegPath -Name "Id" -Value $CatalogGuid -Force | Out-Null
Set-ItemProperty -Path $catRegPath -Name "Url" -Value "\\$env:COMPUTERNAME\$ShareName" -Force | Out-Null
Set-ItemProperty -Path $catRegPath -Name "Flags" -Value 1 -Type DWord -Force | Out-Null

# 7. 共享本地目录以建立信任源
Write-Host "[4/4] 共享本地目录以建立网络信任源..." -ForegroundColor Gray
cmd.exe /c "net share $ShareName=`"$ScriptDir\.`" /grant:everyone,READ >nul 2>&1"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "       安装与配置成功！当前生效线路: [$TargetRoute]" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "使用提示:"
Write-Host "1. 打开 Microsoft Word；"
Write-Host "2. 在【开始】功能区最右侧找到【AI 写作助手】图标点击打开；"
Write-Host "   （或在【插入】->【我的加载项】->【共享文件夹】中点击添加）；"
Write-Host "3. 后续若需要切换线路，可直接双击运行："
Write-Host "   - install-to-word.bat cf   (强制切换至 Cloudflare 极速线路)"
Write-Host "   - install-to-word.bat gh   (强制切换至 GitHub Pages 线路)"
Write-Host "   - install-to-word.bat      (默认智能自动测速与优选)"
Write-Host ""
