@echo off
chcp 65001 >nul

:: Office 一键安装部署脚本
:: 作者: 二进制 (sindri)
:: 博客: https://blog.nbvil.com

setlocal enabledelayedexpansion

:: 检查管理员权限
>nul 2>&1 net session || (
  echo [错误] 请使用管理员身份运行此脚本.
  pause
  exit /b
)

:: 设置变量
set SETUP_URL=https://download.microsoft.com/download/1/2/3/12345678-abcd-1234-abcd-12345678abcd/setup.exe
set SETUP_EXE=setup.exe
set CONFIG=config.xml
set LOGFOLDER=%~dp0logs
set OFFICE_DIR=%~dp0Office
set KMS_HOST=kms.03k.org
set OFFICE_PATH32="C:\Program Files (x86)\Microsoft Office\Office16"
set OFFICE_PATH64="C:\Program Files\Microsoft Office\Office16"

if not exist %LOGFOLDER% mkdir %LOGFOLDER%

:: 步骤 1: 下载 Office 部署工具
if not exist %SETUP_EXE% (
    echo [1/4] 正在下载 Office 安装器...
    powershell -Command "Invoke-WebRequest -Uri %SETUP_URL% -OutFile '%SETUP_EXE%'"
    if not exist %SETUP_EXE% (
        echo [错误] 安装器下载失败，请检查网络或代理设置.
        pause
        exit /b
    )
)
echo [√] 安装器准备完毕.

:: 步骤 2: 下载 Office 安装包
if not exist %OFFICE_DIR% mkdir %OFFICE_DIR%

echo [2/4] 正在下载 Office 安装包，请耐心等待 (首次下载约需几分钟)...
start "" /wait %SETUP_EXE% /download %CONFIG%

:: 检查下载是否完成
if not exist "%OFFICE_DIR%\Data" (
    echo [错误] 下载未成功，请检查网络或代理.
    pause
    exit /b
)
echo [√] 安装包下载完成.

:: 步骤 3: 安装 Office
echo [3/4] 正在安装 Office...
start "" /wait %SETUP_EXE% /configure %CONFIG%

:: 检查安装结果（此处可扩展日志检查）
echo [√] 安装完成.

:: 步骤 4: 激活 Office
if exist %OFFICE_PATH64% (
    cd /d %OFFICE_PATH64%
) else if exist %OFFICE_PATH32% (
    cd /d %OFFICE_PATH32%
) else (
    echo [错误] 找不到 Office 安装目录，跳过激活步骤.
    pause
    exit /b
)

echo [4/4] 正在激活 Office...
cscript ospp.vbs /sethst:%KMS_HOST%
cscript ospp.vbs /act

echo [√] Office 激活完成.
echo.
echo 感谢使用由 二进制 (sindri) 开发的 Office 安装工具
pause
exit /b
