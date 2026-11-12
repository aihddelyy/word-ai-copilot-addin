# ==========================================================================
# Word AI 创作助手 - 统一卸载脚本 (支持 Microsoft Word 与 WPS 文字)
# ==========================================================================
param(
    [string]$App = ""
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(65001) } catch {}
$Host.UI.RawUI.WindowTitle = "Word AI 创作助手 - 卸载程序"

Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host "    Word AI 创作助手 - 一键卸载工具 (Word / WPS 通用)" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host ""

$targetApp = $App.ToLower().Trim()

if ($targetApp -eq "") {
    Write-Host "请选择要卸载的目标办公软件：" -ForegroundColor Yellow
    Write-Host "  [1] 仅从 Microsoft Word 卸载"
    Write-Host "  [2] 仅从 WPS 文字卸载"
    Write-Host "  [3] 同时从 Word 和 WPS 卸载 (完全清理)" -ForegroundColor Green
    Write-Host "  [0] 取消退出"
    Write-Host ""
    $choice = Read-Host "请输入数字选项 [默认 3]"
    if ($choice -eq "") { $choice = "3" }

    switch ($choice) {
        "1" { $targetApp = "word" }
        "2" { $targetApp = "wps" }
        "3" { $targetApp = "all" }
        "0" { Write-Host "已取消卸载。"; exit 0 }
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
Write-Host "准备卸载的目标: " -NoNewline -ForegroundColor Gray
if ($targetApp -eq "word") { Write-Host "Microsoft Word" -ForegroundColor Yellow }
elseif ($targetApp -eq "wps") { Write-Host "WPS 文字 (WPS Office)" -ForegroundColor Yellow }
else { Write-Host "Microsoft Word + WPS 文字 (完全清理)" -ForegroundColor Yellow }
Write-Host ""

# 1. 卸载 Word 配置
function Uninstall-FromWord {
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  正在清理 Microsoft Word 加载项记录..." -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray

    $Id = "d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1"
    $CatalogGuid = "{d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1}"
    $ShareName = "WordAIAssistant"

    # 清理注册表
    Write-Host "  -> 清理 Office 开发者旁加载记录..." -ForegroundColor Gray
    $devRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
    if (Test-Path $devRegPath) {
        Remove-ItemProperty -Path $devRegPath -Name $Id -ErrorAction SilentlyContinue
    }

    Write-Host "  -> 清理受信任目录配置..." -ForegroundColor Gray
    $catalogRegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"
    if (Test-Path $catalogRegPath) {
        Remove-Item -Path $catalogRegPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 清理本地共享
    Write-Host "  -> 删除本地网络共享..." -ForegroundColor Gray
    Start-Process cmd -ArgumentList "/c net share $ShareName /delete /y" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue

    Write-Host "  [√] Microsoft Word 相关加载项已成功清除！" -ForegroundColor Green
}

# 2. 卸载 WPS 配置
function Uninstall-FromWps {
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray
    Write-Host "  正在清理 WPS 文字加载项配置..." -ForegroundColor Cyan
    Write-Host "----------------------------------------------------------" -ForegroundColor Gray

    $appData = [Environment]::GetFolderPath('ApplicationData')
    $jsaddonsDir = Join-Path $appData "kingsoft\wps\jsaddons"
    $targetAddonDir = Join-Path $jsaddonsDir "WordAICopilot"
    $publishXmlPath = Join-Path $jsaddonsDir "publish.xml"

    if (Test-Path $targetAddonDir) {
        Write-Host "  -> 删除 WPS 插件目录: $targetAddonDir" -ForegroundColor Gray
        Remove-Item -Path $targetAddonDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $publishXmlPath) {
        try {
            [xml]$publishDoc = Get-Content $publishXmlPath -Raw -Encoding UTF8
            $node = $publishDoc.SelectSingleNode("//addon[@name='WordAICopilot']")
            if ($null -ne $node) {
                $node.ParentNode.RemoveChild($node) | Out-Null
                $publishDoc.Save($publishXmlPath)
                Write-Host "  -> 已从 WPS publish.xml 移除注册记录" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  -> 清理 publish.xml 异常: $_" -ForegroundColor Yellow
        }
    }

    Write-Host "  [√] WPS 文字相关加载项已成功清除！" -ForegroundColor Green
}

if ($targetApp -eq "word" -or $targetApp -eq "all") {
    Uninstall-FromWord
}
if ($targetApp -eq "wps" -or $targetApp -eq "all") {
    Uninstall-FromWps
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "                     卸载清理完成！" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "重新启动对应办公软件后配置将完全复原生效。" -ForegroundColor Gray
Write-Host ""

