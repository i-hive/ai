@echo off
set LOGFILE=%~dp0script_log.txt
set PYTHON_PATH="C:\Program Files (x86)\Python37-32\python.exe"
set SCRIPT_PATH="%~dp0自动日常检查.py"

echo ======================================== >> %LOGFILE%
echo [%date% %time%] 开始执行Python脚本 >> %LOGFILE%
echo Python路径: %PYTHON_PATH% >> %LOGFILE%
echo 脚本路径: %SCRIPT_PATH% >> %LOGFILE%

%PYTHON_PATH% %SCRIPT_PATH%
set EXIT_CODE=%errorlevel%

if %EXIT_CODE% equ 0 (
    echo [%date% %time%] 脚本执行成功 >> %LOGFILE%
) else (
    echo [%date% %time%] 脚本执行失败，错误代码: %EXIT_CODE% >> %LOGFILE%
)

echo 执行完成，退出代码: %EXIT_CODE% >> %LOGFILE%