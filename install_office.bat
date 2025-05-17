@echo off
chcp 65001 >nul

:: Office һ����װ����ű�
:: ����: ������ (sindri)
:: ����: https://blog.nbvil.com

setlocal enabledelayedexpansion

:: ������ԱȨ��
>nul 2>&1 net session || (
  echo [����] ��ʹ�ù���Ա������д˽ű�.
  pause
  exit /b
)

:: ���ñ���
set SETUP_URL=https://download.microsoft.com/download/1/2/3/12345678-abcd-1234-abcd-12345678abcd/setup.exe
set SETUP_EXE=setup.exe
set CONFIG=config.xml
set LOGFOLDER=%~dp0logs
set OFFICE_DIR=%~dp0Office
set KMS_HOST=kms.03k.org
set OFFICE_PATH32="C:\Program Files (x86)\Microsoft Office\Office16"
set OFFICE_PATH64="C:\Program Files\Microsoft Office\Office16"

if not exist %LOGFOLDER% mkdir %LOGFOLDER%

:: ���� 1: ���� Office ���𹤾�
if not exist %SETUP_EXE% (
    echo [1/4] �������� Office ��װ��...
    powershell -Command "Invoke-WebRequest -Uri %SETUP_URL% -OutFile '%SETUP_EXE%'"
    if not exist %SETUP_EXE% (
        echo [����] ��װ������ʧ�ܣ�����������������.
        pause
        exit /b
    )
)
echo [��] ��װ��׼�����.

:: ���� 2: ���� Office ��װ��
if not exist %OFFICE_DIR% mkdir %OFFICE_DIR%

echo [2/4] �������� Office ��װ���������ĵȴ� (�״�����Լ�輸����)...
start "" /wait %SETUP_EXE% /download %CONFIG%

:: ��������Ƿ����
if not exist "%OFFICE_DIR%\Data" (
    echo [����] ����δ�ɹ���������������.
    pause
    exit /b
)
echo [��] ��װ���������.

:: ���� 3: ��װ Office
echo [3/4] ���ڰ�װ Office...
start "" /wait %SETUP_EXE% /configure %CONFIG%

:: ��鰲װ������˴�����չ��־��飩
echo [��] ��װ���.

:: ���� 4: ���� Office
if exist %OFFICE_PATH64% (
    cd /d %OFFICE_PATH64%
) else if exist %OFFICE_PATH32% (
    cd /d %OFFICE_PATH32%
) else (
    echo [����] �Ҳ��� Office ��װĿ¼�����������.
    pause
    exit /b
)

echo [4/4] ���ڼ��� Office...
cscript ospp.vbs /sethst:%KMS_HOST%
cscript ospp.vbs /act

echo [��] Office �������.
echo.
echo ��лʹ���� ������ (sindri) ������ Office ��װ����
pause
exit /b
