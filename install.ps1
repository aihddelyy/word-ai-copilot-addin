# ==========================================================================
# Word AI 创作助手 - 统一智能安装与配置脚本 (支持 Microsoft Word 与 WPS 文字)
# ==========================================================================
param(
    [string]$App = "",
    [string]$Route = ""
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(65001) } catch {}
$Host.UI.RawUI.WindowTitle = "Word AI 创作助手 - 统一安装配置"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "    Word AI 创作助手 - 一键安装配置 (Word / WPS 通用)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ManifestPath = Join-Path $ScriptDir "manifest.xml"
$Id = "d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
$CatalogGuid = "{d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
$ShareName = "WordAIAssistant"

# 1. 解析目标软件参数
$targetApp = $App.ToLower().Trim()
$targetRoute = $Route.ToLower().Trim()

# 若第一个参数传的是线路（如 cf 或 gh），则智能调整
if ($targetApp -match "^(cf|cloudflare|gh|github)$" -and $targetRoute -eq "") {
    $targetRoute = $targetApp
    $targetApp = ""
}

if ($targetApp -eq "") {
    Write-Host "请选择要安装的目标办公软件：" -ForegroundColor Yellow
    Write-Host "  [1] Microsoft Word (Office 365 / 2021 / 2019 / 2016)"
    Write-Host "  [2] WPS 文字 (WPS Office / 个人版 / 专业版 / 信创版)"
    Write-Host "  [3] 同时安装到 Word 和 WPS 两者 (推荐)" -ForegroundColor Green
    Write-Host "  [0] 退出"
    Write-Host ""
    $choice = Read-Host "请输入数字选项 [默认 3]"
    if ($choice -eq "") { $choice = "3" }

    switch ($choice) {
        "1" { $targetApp = "word" }
        "2" { $targetApp = "wps" }
        "3" { $targetApp = "all" }
        "0" { Write-Host "已取消安装。"; exit 0 }
        default { $targetApp = "all" }
    }
} elseif ($targetApp -match "^(1|w|word|office)$") {
    $targetApp = "word"
} elseif ($targetApp -match "^(2|wps|kingsoft)$") {
    $targetApp = "wps"
} elseif ($targetApp -match "^(3|all|both)$") {
    $targetApp = "all"
} else {
    $targetApp = "all"
}

Write-Host ""
Write-Host "已选择安装目标: " -NoNewline -ForegroundColor Gray
if ($targetApp -eq "word") { Write-Host "Microsoft Word" -ForegroundColor Green }
elseif ($targetApp -eq "wps") { Write-Host "WPS 文字 (WPS Office)" -ForegroundColor Green }
else { Write-Host "Microsoft Word + WPS 文字 (双平台)" -ForegroundColor Green }
Write-Host ""

# 2. 线路智能测速与选择
$ghProbeUrl = "https://aihddelyy.github.io/word-ai-copilot-addin/assets/icon-32.png"
$ghTaskpaneUrl = "https://aihddelyy.github.io/word-ai-copilot-addin/src/taskpane.html?host=wps"

$cfProbeUrl = "https://word-ai-copilot-addin.pages.dev/assets/icon-32.png"
$cfTaskpaneUrl = "https://word-ai-copilot-addin.pages.dev/src/taskpane.html?host=wps"

function Test-EndpointLatency {
    param([string]$Url, [int]$TimeoutSec = 3)
    try {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Method = "HEAD"
        $req.Timeout = $TimeoutSec * 1000
        $req.UserAgent = "WordAICopilot-Installer/1.0"
        $resp = $req.GetResponse()
        $status = [int]$resp.StatusCode
        $resp.Close()
        $sw.Stop()
        if ($status -ge 200 -and $status -lt 400) {
            return $sw.ElapsedMilliseconds
        }
    } catch {
        return -1
    }
    return -1
}

$chosenRoute = ""
if ($targetRoute -match "^(cf|cloudflare)$") {
    $chosenRoute = "cloudflare"
    Write-Host "[*] 已手动指定生效线路: [Cloudflare Pages 极速加速线路]" -ForegroundColor Yellow
} elseif ($targetRoute -match "^(gh|github)$") {
    $chosenRoute = "github"
    Write-Host "[*] 已手动指定生效线路: [GitHub Pages 官方线路]" -ForegroundColor Yellow
} else {
    Write-Host "[*] 正在探测网络连通性与最优线路延迟..." -ForegroundColor Gray
    $ghMs = Test-EndpointLatency -Url $ghProbeUrl -TimeoutSec 3
    if ($ghMs -gt 0) {
        Write-Host "  -> [√] GitHub Pages 连接畅通 (延迟: ${ghMs}ms)" -ForegroundColor Green
        $chosenRoute = "github"
    } else {
        Write-Host "  -> [!] GitHub Pages 超时或无法直连，测速 Cloudflare 极速备用线路..." -ForegroundColor Yellow
        $cfMs = Test-EndpointLatency -Url $cfProbeUrl -TimeoutSec 3
        if ($cfMs -gt 0) {
            Write-Host "  -> [√] Cloudflare Pages 连接通畅 (延迟: ${cfMs}ms)" -ForegroundColor Green
            $chosenRoute = "cloudflare"
        } else {
            Write-Host "  -> [!] 网络受限，默认采用 Cloudflare Pages 备用容灾线路。" -ForegroundColor Yellow
            $chosenRoute = "cloudflare"
        }
    }
}

# 3. 执行 Microsoft Word 安装
function Install-ForWord {
    Write-Host ""
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "  正在配置 Microsoft Word 加载项..." -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkCyan

    if (-not (Test-Path $ManifestPath)) {
        Write-Host "  [错误] 未找到 manifest.xml 清单文件！" -ForegroundColor Red
        return
    }

    # 更新 manifest.xml 中的线上 URL
    $content = [System.IO.File]::ReadAllText($ManifestPath, [System.Text.Encoding]::UTF8)
    if ($chosenRoute -eq "cloudflare") {
        $content = $content -replace "https://aihddelyy\.github\.io/word-ai-copilot-addin", "https://word-ai-copilot-addin.pages.dev"
        Write-Host "  -> 清单基准地址指向: https://word-ai-copilot-addin.pages.dev" -ForegroundColor Gray
    } else {
        $content = $content -replace "https://word-ai-copilot-addin\.pages\.dev", "https://aihddelyy.github.io/word-ai-copilot-addin"
        Write-Host "  -> 清单基准地址指向: https://aihddelyy.github.io/word-ai-copilot-addin" -ForegroundColor Gray
    }
    [System.IO.File]::WriteAllText($ManifestPath, $content, [System.Text.Encoding]::UTF8)

    # 关闭正在运行的 Word
    $wordProc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    if ($wordProc) {
        Write-Host "  -> 检测到 Word 正在运行，正在自动保存并关闭..." -ForegroundColor Yellow
        Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 600
    }

    # 清除 WEF 开发者缓存
    Write-Host "  -> 刷新 Word 加载项运行缓存..." -ForegroundColor Gray
    $wefDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
    if (Test-Path $wefDir) {
        Remove-Item -Path "$wefDir\*" -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 注册 Developer 注册表
    Write-Host "  -> 写入 Office 开发者旁加载清单注册表..." -ForegroundColor Gray
    $devRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
    if (-not (Test-Path $devRegPath)) {
        New-Item -Path $devRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $devRegPath -Name $Id -Value $ManifestPath -Type String -Force

    # 注册受信任目录
    Write-Host "  -> 注册受信任网络与本地目录..." -ForegroundColor Gray
    $catalogRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"
    if (-not (Test-Path $catalogRegPath)) {
        New-Item -Path $catalogRegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $catalogRegPath -Name "Id" -Value $CatalogGuid -Type String -Force
    Set-ItemProperty -Path $catalogRegPath -Name "Url" -Value "\\localhost\$ShareName" -Type String -Force
    Set-ItemProperty -Path $catalogRegPath -Name "Flags" -Value 1 -Type DWord -Force

    Write-Host "  [√] Microsoft Word 加载项配置成功！" -ForegroundColor Green
}

# 4. 执行 WPS 文字安装
function Install-ForWps {
    Write-Host ""
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "  正在配置 WPS 文字 (WPS Writer) 加载项..." -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkCyan

    $appData = [Environment]::GetFolderPath('ApplicationData')
    $jsaddonsDir = Join-Path $appData "kingsoft\wps\jsaddons"
    $targetAddonDir = Join-Path $jsaddonsDir "WordAICopilot"
    $sourceAddonDir = Join-Path $ScriptDir "wps-addon"

    if (!(Test-Path $jsaddonsDir)) {
        New-Item -ItemType Directory -Path $jsaddonsDir -Force | Out-Null
    }
    if (!(Test-Path $targetAddonDir)) {
        New-Item -ItemType Directory -Path $targetAddonDir -Force | Out-Null
    }

    # 更新 wps-addon/main.js 中的入口地址
    $wpsTaskpaneUrl = if ($chosenRoute -eq "cloudflare") { $cfTaskpaneUrl } else { $ghTaskpaneUrl }
    $mainJsPath = Join-Path $sourceAddonDir "main.js"
    if (Test-Path $mainJsPath) {
        $mainContent = Get-Content $mainJsPath -Raw -Encoding UTF8
        $escapedUrl = $wpsTaskpaneUrl.Replace("\", "\\")
        $mainContent = [regex]::Replace($mainContent, 'return "https?://[^"]+\?host=wps";', "return `"$escapedUrl`";")
        Set-Content -Path $mainJsPath -Value $mainContent -Encoding UTF8
    }

    # 复制插件资源
    Write-Host "  -> 同步插件资源至 WPS 加载项目录..." -ForegroundColor Gray
    Copy-Item -Path (Join-Path $sourceAddonDir "*") -Destination $targetAddonDir -Recurse -Force

    # 注册 publish.xml
    Write-Host "  -> 更新 WPS publish.xml 清单..." -ForegroundColor Gray
    $publishXmlPath = Join-Path $jsaddonsDir "publish.xml"
    $entryHtmlPath = (Join-Path $targetAddonDir "index.html").Replace("\", "/")

    [xml]$publishDoc = $null
    if (Test-Path $publishXmlPath) {
        try { $publishDoc = [xml](Get-Content $publishXmlPath -Raw -Encoding UTF8) } catch { $publishDoc = $null }
    }
    if ($null -eq $publishDoc -or $null -eq $publishDoc.DocumentElement) {
        $publishDoc = [xml]"<?xml version=`"1.0`" encoding=`"utf-8`"?><jsaddons></jsaddons>"
    }

    $existingNode = $publishDoc.SelectSingleNode("//addon[@name='WordAICopilot']")
    if ($null -ne $existingNode) {
        $existingNode.SetAttribute("title", "Word AI 创作助手")
        $existingNode.SetAttribute("version", "1.0.0")
        $existingNode.SetAttribute("online", "false")
        $existingNode.SetAttribute("addonType", "wps")
        $existingNode.SetAttribute("url", $entryHtmlPath)
    } else {
        $newAddon = $publishDoc.CreateElement("addon")
        $newAddon.SetAttribute("name", "WordAICopilot")
        $newAddon.SetAttribute("title", "Word AI 创作助手")
        $newAddon.SetAttribute("version", "1.0.0")
        $newAddon.SetAttribute("online", "false")
        $newAddon.SetAttribute("addonType", "wps")
        $newAddon.SetAttribute("url", $entryHtmlPath)
        $publishDoc.DocumentElement.AppendChild($newAddon) | Out-Null
    }
    $publishDoc.Save($publishXmlPath)

    Write-Host "  [√] WPS 文字加载项配置成功！" -ForegroundColor Green
}

# 5. 根据选择执行具体安装任务
if ($targetApp -eq "word" -or $targetApp -eq "all") {
    Install-ForWord
}
if ($targetApp -eq "wps" -or $targetApp -eq "all") {
    Install-ForWps
}

# 6. 安装完成提示
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "                   安装配置全部完成！" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "👉 启动与使用指南：" -ForegroundColor Cyan

if ($targetApp -eq "word" -or $targetApp -eq "all") {
    Write-Host "  【Microsoft Word】:" -ForegroundColor Yellow
    Write-Host "    - 启动 Word -> 顶部【开始】选项卡最右侧可看到【AI 创作助手】；"
    Write-Host "    - 或在【插入】->【我的加载项】->【开发人员】中点击启动。"
}
if ($targetApp -eq "wps" -or $targetApp -eq "all") {
    Write-Host "  【WPS 文字 (WPS Office)】:" -ForegroundColor Yellow
    Write-Host "    - 启动或重启 WPS 文字 -> 顶部功能区将出现专属【AI 创作助手】选项卡；"
    Write-Host "    - 点击大图标【打开创作助手】即可在右侧唤出智能侧边栏！"
}

Write-Host ""
Write-Host "提示：若需要卸载，随时双击运行 uninstall.bat 即可一键移除。" -ForegroundColor Gray
Write-Host ""
