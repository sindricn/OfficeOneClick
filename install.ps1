# install.ps1
chcp 65001 > $null
$ErrorActionPreference = "Stop"

function Show-Menu {
    Clear-Host
    $line = "=" * 50
    Write-Host "`n$line" -ForegroundColor Cyan
    Write-Host "            OfficeOneClick 安装工具" -ForegroundColor Cyan
    Write-Host "$line" -ForegroundColor Cyan
    Write-Host " 1. 一键安装 Office"
    Write-Host " 2. 下载 Office 安装工具"
    Write-Host " 3. 下载 Office 安装包"
    Write-Host " 4. 执行 Office 安装"
    Write-Host " 5. 激活 Office"
    Write-Host " 6. 卸载脚本及缓存文件"
    Write-Host " 0. 退出"
    Write-Host "$line`n" -ForegroundColor Cyan
}

function Download-File($url, $output) {
    try {
        if (-Not (Test-Path $output)) {
            Write-Host "`n[*] 正在下载 $output ..." -ForegroundColor Gray
            $client = New-Object System.Net.WebClient
            $client.DownloadFile($url, $output)
            Write-Host "[✓] 下载完成：$output" -ForegroundColor Green
        } else {
            Write-Host "[!] $output 已存在，跳过下载。" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[×] 下载失败：$($_.Exception.Message)" -ForegroundColor Red
        Exit
    }
}

function Download-SetupTool {
    Download-File "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe" "setup.exe"
    Download-File "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml" "config.xml"
}

function Download-OfficePackage {
    if (-Not (Test-Path "Office")) { New-Item -ItemType Directory -Path "Office" > $null }

    Write-Host "`n[*] 正在下载 Office 安装包（耗时可能较长）..." -ForegroundColor Gray
    $proc = Start-Process -FilePath "setup.exe" -ArgumentList "/download config.xml" -Wait -PassThru -NoNewWindow

    if ($proc.ExitCode -ne 0 -or -Not (Test-Path "Office\Data")) {
        Write-Host "[×] 安装包下载失败，可能是网络或配置问题。" -ForegroundColor Red
        Exit
    }

    Write-Host "[✓] Office 安装包下载完成。" -ForegroundColor Green
}

function Install-Office {
    Write-Host "`n[*] 正在安装 Office ..." -ForegroundColor Gray
    $proc = Start-Process -FilePath "setup.exe" -ArgumentList "/configure config.xml" -Wait -PassThru -NoNewWindow

    if ($proc.ExitCode -ne 0) {
        Write-Host "[×] Office 安装失败。" -ForegroundColor Red
        Exit
    }

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
        Exit
    }

    Write-Host "`n[*] 正在激活 Office ..." -ForegroundColor Gray
    cscript ospp.vbs /sethst:$kmsHost | Out-Null
    cscript ospp.vbs /act | Out-Null

    Write-Host "[✓] 激活完成。" -ForegroundColor Green
}

function Full-Install {
    Download-SetupTool
    Download-OfficePackage
    Install-Office
    Activate-Office
}

function Cleanup-Script {
    Write-Host "`n[*] 正在清理脚本和缓存文件 ..." -ForegroundColor Gray
    $items = @("setup.exe", "Office", "logs", "install.ps1", "config.xml")

    foreach ($item in $items) {
        if (Test-Path $item) {
            Remove-Item $item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "已删除：$item" -ForegroundColor DarkGray
        }
    }

    Write-Host "[✓] 清理完成，不影响已安装的 Office 使用。" -ForegroundColor Yellow
    Pause
    Exit
}

# 权限检测
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "[×] 请以管理员身份运行此脚本。" -ForegroundColor Red
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
        default {
            Write-Host "[!] 无效输入，请输入 0-6 的数字。" -ForegroundColor Red
        }
    }
    Write-Host ""  # 空行缓冲
    Pause
} while ($true)
