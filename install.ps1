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

function Download-ConfigFile {
    $url = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml"
    $output = "config.xml"
    if (-Not (Test-Path $output)) {
        Write-Host "[*] 正在下载配置文件 config.xml..."
        Invoke-WebRequest -Uri $url -OutFile $output
        Write-Host "[✓] 配置文件下载完成。" -ForegroundColor Green
    } else {
        Write-Host "[!] config.xml 已存在，跳过下载。" -ForegroundColor Yellow
    }
}

function Download-OfficePackage {
    try {
        if (-Not (Test-Path "Office")) { New-Item -ItemType Directory -Path "Office" > $null }
        if (-Not (Test-Path "setup.exe")) {
            Write-Host "[×] 未找到 setup.exe，请先下载安装工具。" -ForegroundColor Red
            return
        }
        if (-Not (Test-Path "config.xml")) {
            Write-Host "[×] 未找到 config.xml，尝试重新下载。" -ForegroundColor Red
            Download-ConfigFile
        }
        Write-Host "[*] 正在下载 Office 安装包，请稍候..."
        Start-Process -Wait -FilePath "setup.exe" -ArgumentList "/download config.xml"
        if (Test-Path "Office\Data") {
            Write-Host "[✓] 安装包下载成功。" -ForegroundColor Green
        } else {
            throw "安装包目录缺失，下载失败。"
        }
    } catch {
        Write-Host "[×] $_" -ForegroundColor Red
    }
}

function Install-Office {
    try {
        if (-Not (Test-Path "setup.exe") -or -Not (Test-Path "config.xml")) {
            Write-Host "[×] 缺少必要文件，无法安装。" -ForegroundColor Red
            return
        }
        Write-Host "[*] 正在安装 Office..."
        Start-Process -Wait -FilePath "setup.exe" -ArgumentList "/configure config.xml"
        Write-Host "[✓] 安装完成。" -ForegroundColor Green
    } catch {
        Write-Host "[×] 安装失败：$_" -ForegroundColor Red
    }
}

function Activate-Office {
    $officePath64 = "C:\Program Files\Microsoft Office\Office16"
    $officePath32 = "C:\Program Files (x86)\Microsoft Office\Office16"
    $kmsHost = "kms.03k.org"
    try {
        if (Test-Path $officePath64) {
            Set-Location $officePath64
        } elseif (Test-Path $officePath32) {
            Set-Location $officePath32
        } else {
            throw "未找到 Office 安装路径"
        }
        Write-Host "[*] 正在激活 Office..."
        cscript ospp.vbs /sethst:$kmsHost
        cscript ospp.vbs /act
        Write-Host "[✓] 激活完成。" -ForegroundColor Green
    } catch {
        Write-Host "[×] 激活失败：$_" -ForegroundColor Red
    }
}

function Full-Install {
    try {
        Download-SetupTool
        Download-ConfigFile
        Download-OfficePackage
        Install-Office
        Activate-Office
    } catch {
        Write-Host "[×] 执行失败：$_" -ForegroundColor Red
    }
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
        "3" { Download-ConfigFile; Download-OfficePackage }
        "4" { Install-Office }
        "5" { Activate-Office }
        "6" { Cleanup-Script }
        "0" { break }
        default { Write-Host "[!] 请输入有效编号。" -ForegroundColor Red }
    }
    Write-Host ""
    Pause
} while ($true)
