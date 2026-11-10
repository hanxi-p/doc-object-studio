@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   文档对象工坊 - Windows EXE 构建
echo ============================================

call "%~dp0activate_test_env.bat"
if errorlevel 1 goto :fail

python --version
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "文档对象工坊_全篇勾选版" ^
  --icon "web\app-icon.ico" ^
  --add-data "web;web" ^
  --exclude-module=PIL.ImageQt ^
  --collect-all fitz ^
  --collect-all docx ^
  --collect-all pptx ^
  --collect-all lxml ^
  --collect-all win32com ^
  --hidden-import=win32timezone ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --hidden-import=win32com.client ^
  webui.py
if errorlevel 1 goto :fail

echo.
echo [完成] EXE 已生成：dist\文档对象工坊_全篇勾选版.exe
echo 提示：PPTX/DOCX 的原生复制仍需本机安装 Microsoft Office。
pause
exit /b 0

:fail
echo.
echo [失败] 构建未完成。请查看上方错误信息。
pause
exit /b 1
