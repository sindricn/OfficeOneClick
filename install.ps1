# Office-Install-Deploy
# 自动化安装和激活 Microsoft Office 的 PowerShell 脚本
# GitHub: https://github.com/sindricn/OfficeOneClick

# 管理员权限检查
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "请以管理员身份运行此脚本！"
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 设置工作目录
$workDir = "$env:TEMP\OfficeInstall"
if (!(Test-Path $workDir)) {
    New-Item -Path $workDir -ItemType Directory -Force | Out-Null
}
Set-Location $workDir
Write-Host "工作目录: $workDir" -ForegroundColor Cyan

# 从 GitHub 仓库下载 setup.exe 和 config.xml
Write-Host "正在从 GitHub 仓库下载安装文件..." -ForegroundColor Cyan
$setupUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe"
$configUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml"

try {
    Invoke-WebRequest -Uri $setupUrl -OutFile "$workDir\setup.exe"
    Invoke-WebRequest -Uri $configUrl -OutFile "$workDir\config.xml"
    Write-Host "安装文件下载成功。" -ForegroundColor Green
} catch {
    Write-Host "下载安装文件失败。请检查网络连接或访问 https://github.com/sindricn/OfficeOneClick 手动下载" -ForegroundColor Red
    Write-Host "错误详情: $_" -ForegroundColor Red
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 下载 Office
Write-Host "正在下载 Office（这可能需要一些时间）..." -ForegroundColor Cyan
Start-Process -FilePath "$workDir\setup.exe" -ArgumentList "/download", "$workDir\config.xml" -Wait
Write-Host "Office 安装包下载成功。" -ForegroundColor Green

# 安装 Office
Write-Host "正在安装 Office（这可能需要一些时间）..." -ForegroundColor Cyan
Start-Process -FilePath "$workDir\setup.exe" -ArgumentList "/configure", "$workDir\config.xml" -Wait
Write-Host "Office 安装成功。" -ForegroundColor Green

# 确定 Office 安装路径
$officePath = ""
if (Test-Path "C:\Program Files\Microsoft Office\Office16") {
    $officePath = "C:\Program Files\Microsoft Office\Office16"
} elseif (Test-Path "C:\Program Files (x86)\Microsoft Office\Office16") {
    $officePath = "C:\Program Files (x86)\Microsoft Office\Office16"
} else {
    Write-Host "无法找到 Office 安装路径。" -ForegroundColor Red
    Write-Host "请手动激活 Office。" -ForegroundColor Red
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 激活 Office
Write-Host "正在通过 KMS 激活 Office..." -ForegroundColor Cyan
Set-Location $officePath
& cscript ospp.vbs /sethst:kms.03k.org
& cscript ospp.vbs /act
Write-Host "Office 激活过程完成。" -ForegroundColor Green

# 清理
Set-Location $env:USERPROFILE
Write-Host "正在清理临时文件..." -ForegroundColor Cyan
Remove-Item -Path $workDir -Recurse -Force
Write-Host "清理完成。" -ForegroundColor Green

# 完成
Write-Host "`nOffice 安装和激活已成功完成！" -ForegroundColor Green
Write-Host "您现在可以使用 Microsoft Office 产品了。" -ForegroundColor Green
Write-Host "按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
