@echo off
set LOGFILE=%~dp0script_log.txt
set PYTHON_PATH="C:\Program Files (x86)\Python37-32\python.exe"
set SCRIPT_PATH="%~dp0�Զ��ճ����.py"

echo ======================================== >> %LOGFILE%
echo [%date% %time%] ��ʼִ��Python�ű� >> %LOGFILE%
echo Python·��: %PYTHON_PATH% >> %LOGFILE%
echo �ű�·��: %SCRIPT_PATH% >> %LOGFILE%

%PYTHON_PATH% %SCRIPT_PATH%
set EXIT_CODE=%errorlevel%

if %EXIT_CODE% equ 0 (
    echo [%date% %time%] �ű�ִ�гɹ� >> %LOGFILE%
) else (
    echo [%date% %time%] �ű�ִ��ʧ�ܣ��������: %EXIT_CODE% >> %LOGFILE%
)

echo ִ����ɣ��˳�����: %EXIT_CODE% >> %LOGFILE%