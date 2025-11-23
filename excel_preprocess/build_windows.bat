@echo off
echo ============================================
echo       Building GEpmTool (Windows .exe)
echo ============================================

REM ������}���Ҧb�ؿ� (excel_preprocess)
cd /d %~dp0

REM �]�w��X��Ƨ�
set OUTPUT_DIR=..\Tool_Pack_win

REM �إ߿�X��Ƨ�
if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

echo Cleaning previous builds...

REM �R���� build/dist/spec
rmdir /s /q "%OUTPUT_DIR%\build" 2>nul
rmdir /s /q "%OUTPUT_DIR%\GEpmTool" 2>nul
del "%OUTPUT_DIR%\GEpmTool.spec" 2>nul

echo Running PyInstaller...

pyinstaller ^
    --onefile ^
    --distpath "%OUTPUT_DIR%" ^
    --workpath "%OUTPUT_DIR%\build" ^
    --specpath "%OUTPUT_DIR%" ^
    --windowed ^
    --name GEpmTool ^
    --add-data "%~dp0ui_GEpmToolUI.py;." ^
    --add-data "%~dp0..\Doc\report_demo.xlsx;Doc" ^
    --add-data "%~dp0..\Doc\logo.png;Doc" ^
    GUI_Tool.py

REM �T�{ EXE �O�_�ͦ�
if not exist "%OUTPUT_DIR%\GEpmTool.exe" (
    echo ============================================
    echo  ? Build failed! EXE NOT generated!
    echo ============================================
    pause
    exit /b
)

echo ============================================
echo  ?? Build Success!
echo  EXE Output:
echo     %OUTPUT_DIR%\GEpmTool\GEpmTool.exe
echo ============================================
pause