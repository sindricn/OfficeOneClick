# install.ps1 - Office 一键安装脚本（PowerShell 版）
# 作者：二进制（sindri）
# 项目：https://github.com/sindricn/OfficeOneClick

$ErrorActionPreference = "Stop"

# 设置变量
$SETUP_URL   = "https://download.microsoft.com/download/1/2/3/12345678-abcd-1234-abcd-12345678abcd/setup.exe"
$SETUP_EXE   = "setup.exe"
$CONFIG_FILE = "config.xml"
$OFFICE_DIR  = Join-Path $PSScriptRoot "Office"
$LOG_DIR     = Join-Path $PSScriptRoot "logs"
$KMS_HOST    = "kms.03k.org"
$OFFICE_PATH64 = "C:\Program Files\Microsoft Office\Office16"
$OFFICE_PATH32 = "C:\Program Files (x86)\Microsoft Office\Office16"

# 检查管理员权限
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n[错误] 请使用管理员权限运行本脚本。" -ForegroundColor Red
    Pause
    Exit 1
}

# 创建必要目录
If (!(Test-Path $LOG_DIR))   { New-Item -ItemType Directory -Path $LOG_DIR | Out-Null }
If (!(Test-Path $OFFICE_DIR)) { New-Item -ItemType Directory -Path $OFFICE_DIR | Out-Null }

function Download-SetupTool {
    If (!(Test-Path $SETUP_EXE)) {
        Write-Host "`n[1/4] 正在下载 Office 安装工具..."
        Invoke-WebRequest -Uri $SETUP_URL -OutFile $SETUP_EXE
        If (!(Test-Path $SETUP_EXE)) {
            Write-Host "[错误] 安装器下载失败，请检查网络或代理设置。" -ForegroundColor Red
            Pause
            Exit 1
        }
    }
    Write-Host "[✓] 安装器准备完毕。" -ForegroundColor Green
}

function Download-OfficeFiles {
    Write-Host "`n[2/4] 正在下载 Office 安装包..."
    & .\$SETUP_EXE /download $CONFIG_FILE
    If (!(Test-Path "$OFFICE_DIR\Data")) {
        Write-Host "[错误] 安装包下载失败，请检查网络或配置文件。" -ForegroundColor Red
        Pause
        Exit 1
    }
    Write-Host "[✓] 安装包下载完成。" -ForegroundColor Green
}

function Install-Office {
    Write-Host "`n[3/4] 正在安装 Office..."
    & .\$SETUP_EXE /configure $CONFIG_FILE
    Write-Host "[✓] 安装完成。" -ForegroundColor Green
}

function Activate-Office {
    If (Test-Path $OFFICE_PATH64) {
        Set-Location $OFFICE_PATH64
    } elseif (Test-Path $OFFICE_PATH32) {
        Set-Location $OFFICE_PATH32
    } else {
        Write-Host "[错误] 找不到 Office 安装目录，跳过激活步骤。" -ForegroundColor Red
        Pause
        Exit 1
    }

    Write-Host "`n[4/4] 正在激活 Office..."
    cscript ospp.vbs /sethst:$KMS_HOST
    cscript ospp.vbs /act
    Write-Host "[✓] Office 激活完成。" -ForegroundColor Green
}

function Show-Menu {
    Clear-Host
    Write-Host "=== Office 安装工具（PowerShell 版）==="
    Write-Host "作者：二进制（sindri）"
    Write-Host "博客：https://blog.nbvil.com"
    Write-Host "项目：https://github.com/sindricn/OfficeOneClick"
    Write-Host ""
    Write-Host "请选择一个操作："
    Write-Host "1. 下载 Office 安装工具"
    Write-Host "2. 下载 Office 安装包"
    Write-Host "3. 安装 Office"
    Write-Host "4. 激活 Office"
    Write-Host "5. 一键安装全部步骤"
    Write-Host "6. 退出"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "请输入选项数字"

    switch ($choice) {
        "1" { Download-SetupTool }
        "2" { Download-SetupTool; Download-OfficeFiles }
        "3" { Download-SetupTool; Install-Office }
        "4" { Activate-Office }
        "5" {
            Download-SetupTool
            Download-OfficeFiles
            Install-Office
            Activate-Office
        }
        "6" {
            Write-Host "`n感谢使用本工具，再见！" -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "[提示] 无效选项，请重新输入。" -ForegroundColor Yellow
        }
    }

    Write-Host "`n按任意键返回菜单..." -NoNewline
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} while ($true)
