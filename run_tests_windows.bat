@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
call "%~dp0activate_test_env.bat"
if errorlevel 1 goto :fail
if "%DOCSTUDIO_SAMPLE_DIR%"=="" set "DOCSTUDIO_SAMPLE_DIR=%~dp0test_samples"
python tests\test_sample_documents.py
if errorlevel 1 goto :fail
python tests\test_webui_api.py
if errorlevel 1 goto :fail
exit /b 0
:fail
echo Checks did not complete. Check the test environment and the error above.
pause
exit /b 1
