@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
call "%~dp0activate_test_env.bat"
if errorlevel 1 goto :fail
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail
python webui.py
if errorlevel 1 goto :fail
exit /b 0
:fail
echo Web UI startup failed. Check the test environment and the error above.
pause
exit /b 1
