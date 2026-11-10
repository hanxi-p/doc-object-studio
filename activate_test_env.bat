@echo off
call "%USERPROFILE%\miniconda3\condabin\conda.bat" activate test
if errorlevel 1 goto :activate_failed
if /I not "%CONDA_DEFAULT_ENV%"=="test" goto :activate_failed
exit /b 0
:activate_failed
echo [ERROR] Unable to activate the Conda environment named test.
exit /b 1