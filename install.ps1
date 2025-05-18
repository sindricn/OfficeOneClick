# install.ps1
chcp 65001 > $null
$ErrorActionPreference = "Stop"

function Show-Menu {
    Clear-Host
    Write-Host "========== OfficeOneClick 安装工具 ==========" -ForegroundColor Cyan
    Write-Host "1. 一键安装 Office"
    Write-Host "2. 下载 Office 安装工具"
    Write-Host "3. 下载 Office 安装包"
    Write-Host "4. 执行 Office 安装"
    Write-Host "5. 激活 Office"
    Write-Host "6. 卸载脚本及缓存文件"
    Write-Host "0. 退出"
    Write-Host "=============================================`n"
}

function Download-SetupTool {
    $url = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe"
    $output = "setup.exe"
    if (-Not (Test-Path $output)) {
        Write-Host "[*] 正在下载 Office 安装工具..."
        Invoke-WebRequest -Uri $url -OutFile $output
        Write-Host "[✓] 下载完成：$output" -ForegroundColor Green
    } else {
        Write-Host "[!] 安装工具已存在，跳过下载。" -ForegroundColor Yellow
    }
}

function Download-OfficePackage {
    if (-Not (Test-Path "Office")) { New-Item -ItemType Directory -Path "Office" > $null }
    Write-Host "[*] 正在下载 Office 安装包，请稍候..."
    Start-Process -Wait -FilePath "setup.exe" -ArgumentList "/download config.xml"
    if (Test-Path "Office\Data") {
        Write-Host "[✓] Office 安装包下载完成。" -ForegroundColor Green
    } else {
        Write-Host "[×] 下载失败，请检查 config.xml 或网络。" -ForegroundColor Red
    }
}

function Install-Office {
    Write-Host "[*] 正在安装 Office..."
    Start-Process -Wait -FilePath "setup.exe" -ArgumentList "/configure config.xml"
    Write-Host "[✓] 安装完成。" -ForegroundColor Green
}

function Activate-Office {
    $officePath64 = "C:\Program Files\Microsoft Office\Office16"
    $officePath32 = "C:\Program Files (x86)\Microsoft Office\Office16"
    $kmsHost = "kms.03k.org"
    if (Test-Path $officePath64) {
        Set-Location $officePath64
    } elseif (Test-Path $officePath32) {
        Set-Location $officePath32
    } else {
        Write-Host "[×] 未找到 Office 安装路径，激活失败。" -ForegroundColor Red
        return
    }
    Write-Host "[*] 正在激活 Office..."
    cscript ospp.vbs /sethst:$kmsHost
    cscript ospp.vbs /act
    Write-Host "[✓] 激活完成。" -ForegroundColor Green
}

function Full-Install {
    Download-SetupTool
    Download-OfficePackage
    Install-Office
    Activate-Office
}

function Cleanup-Script {
    Write-Host "[*] 正在清理脚本和缓存文件..."
    $items = @("setup.exe", "Office", "logs", "install.ps1", "config.xml")
    foreach ($item in $items) {
        if (Test-Path $item) {
            Remove-Item $item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "已删除：$item" -ForegroundColor Gray
        }
    }
    Write-Host "[✓] 清理完成，不影响 Office 正常使用。" -ForegroundColor Yellow
    Pause
    Exit
}

# 权限检测
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "[×] 请右键以管理员身份运行此脚本。" -ForegroundColor Red
    Pause
    Exit
}

# 主程序循环
do {
    Show-Menu
    $choice = Read-Host "请输入操作编号"
    switch ($choice) {
        "1" { Full-Install }
        "2" { Download-SetupTool }
        "3" { Download-OfficePackage }
        "4" { Install-Office }
        "5" { Activate-Office }
        "6" { Cleanup-Script }
        "0" { break }
        default { Write-Host "[!] 请输入有效编号。" -ForegroundColor Red }
    }
    Pause
} while ($true)
