# install.ps1
chcp 65001 > $null
$ErrorActionPreference = "Stop"

function Show-Menu {
    Clear-Host
    Start-Sleep -Milliseconds 200
    Write-Host ""
    Write-Host "========== OfficeOneClick 安装工具 ==========" -ForegroundColor Cyan
    Write-Host "1. 一键安装 Office"
    Write-Host "2. 下载 Office 安装工具"
    Write-Host "3. 下载 Office 安装包"
    Write-Host "4. 执行 Office 安装"
    Write-Host "5. 激活 Office"
    Write-Host "6. 卸载脚本及缓存文件"
    Write-Host "0. 退出"
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
}

function Download-SetupTool {
    try {
        $url = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe"
        $output = "setup.exe"
        if (-Not (Test-Path $output)) {
            Write-Host "[*] 正在下载 Office 安装工具..."
            Invoke-WebRequest -Uri $url -OutFile $output
            Write-Host "[✓] 下载完成：$output" -ForegroundColor Green
        } else {
            Write-Host "[!] 安装工具已存在，跳过下载。" -ForegroundColor Yellow
        }
        return $true
    } catch {
        Write-Host "[×] 安装工具下载失败：$_" -ForegroundColor Red
        return $false
    }
}

function Download-ConfigFile {
    try {
        $url = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml"
        $output = "config.xml"
        Write-Host "[*] 正在下载配置文件..."
        Invoke-WebRequest -Uri $url -OutFile $output
        Write-Host "[✓] 下载完成：$output" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "[×] 配置文件下载失败：$_" -ForegroundColor Red
        return $false
    }
}

function Download-OfficePackage {
    try {
        if (-Not (Test-Path "setup.exe")) {
            Write-Host "[×] 缺少安装工具 setup.exe" -ForegroundColor Red
            return $false
        }
        if (-Not (Test-Path "config.xml")) {
            if (-not (Download-ConfigFile)) { return $false }
        }

        Write-Host "[*] 正在启动 Office 安装包下载..."
        $proc = Start-Process -FilePath "setup.exe" -ArgumentList "/download config.xml" -PassThru
        Wait-Process -Id $proc.Id
        Start-Sleep -Seconds 2

        if (Test-Path "Office\Data") {
            Write-Host "[✓] 安装包下载成功。" -ForegroundColor Green
            return $true
        } else {
            Write-Host "[×] 未检测到安装包目录，下载失败。" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "[×] 安装包下载异常：$_" -ForegroundColor Red
        return $false
    }
}

function Install-Office {
    try {
        if (-Not (Test-Path "setup.exe") -or -Not (Test-Path "config.xml")) {
            Write-Host "[×] 缺少 setup.exe 或 config.xml，无法安装。" -ForegroundColor Red
            return $false
        }
        Write-Host "[*] 正在安装 Office..."
        $proc = Start-Process -FilePath "setup.exe" -ArgumentList "/configure config.xml" -PassThru
        Wait-Process -Id $proc.Id
        Write-Host "[✓] 安装完成。" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "[×] 安装失败：$_" -ForegroundColor Red
        return $false
    }
}

function Activate-Office {
    try {
        $officePath64 = "C:\Program Files\Microsoft Office\Office16"
        $officePath32 = "C:\Program Files (x86)\Microsoft Office\Office16"
        $kmsHost = "kms.03k.org"

        if (Test-Path $officePath64) {
            Set-Location $officePath64
        } elseif (Test-Path $officePath32) {
            Set-Location $officePath32
        } else {
            Write-Host "[×] 未找到 Office 安装路径，激活失败。" -ForegroundColor Red
            return $false
        }

        Write-Host "[*] 正在激活 Office..."
        cscript ospp.vbs /sethst:$kmsHost
        cscript ospp.vbs /act
        Write-Host "[✓] 激活完成。" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "[×] 激活失败：$_" -ForegroundColor Red
        return $false
    }
}

function Full-Install {
    if (-not (Download-SetupTool)) { return }
    if (-not (Download-ConfigFile)) { return }
    if (-not (Download-OfficePackage)) { return }
    if (-not (Install-Office)) { return }
    if (-not (Activate-Office)) { return }
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
        "2" { Download-SetupTool | Out-Null }
        "3" { Download-OfficePackage | Out-Null }
        "4" { Install-Office | Out-Null }
        "5" { Activate-Office | Out-Null }
        "6" { Cleanup-Script }
        "0" { break }
        default { Write-Host "[!] 请输入有效编号。" -ForegroundColor Red }
    }
    Pause
} while ($true)
