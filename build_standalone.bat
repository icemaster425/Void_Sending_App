@echo off
title V.O.I.D. Portable Builder
echo ======================================================
echo   V.O.I.D. SYSTEM - PORTABLE BUILDER
echo   Verified On-boarding Institutional Dispatcher
echo ======================================================
echo.

REM 1. Clean up previous build folders to ensure a fresh start [cite: 2]
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

echo Step 1: Compiling Executable with PyInstaller... [cite: 3]
echo [This may take a minute. Please wait...] [cite: 3]

REM --noconsole: Hides the command prompt when the app runs. [cite: 4]
REM --onefile: Bundles the app into a single EXE. [cite: 4]
REM --icon: Attaches your custom void_icon.ico. [cite: 4]
REM --hidden-import: Ensures background libraries are included in the build. [cite: 4, 5]

pyinstaller --noconsole --onefile --icon="void_icon.ico" --name "VOID_System" ^
--hidden-import "tkcalendar" ^
--hidden-import "babel.numbers" ^
--hidden-import "pyminizip" ^
--hidden-import "win32com" ^
--hidden-import "pythoncom" ^
--hidden-import "watchdog" ^
main.py 

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed. Please ensure 'void_icon.ico' is in this folder. 
    pause
    exit /b 1
)

echo.
echo Step 2: Creating Portable Folder Structure... [cite: 6]
if not exist "Portable-VOID" mkdir "Portable-VOID"
if not exist "Portable-VOID\To Send" mkdir "Portable-VOID\To Send" [cite: 6]

echo.
echo Step 3: Moving Files to Portable Folder... [cite: 7]
copy "dist\VOID_System.exe" "Portable-VOID\VOID_System.exe" /Y [cite: 7]
copy "config.ini" "Portable-VOID\config.ini" /Y [cite: 7]

REM Only copy the database if it exists, to avoid errors [cite: 7]
if exist "file_monitor.db" (
    copy "file_monitor.db" "Portable-VOID\file_monitor.db" /Y
)

echo.
echo ======================================================
echo   SUCCESS! V.O.I.D. is ready in 'Portable-VOID' [cite: 8]
echo ======================================================
pause