@echo off
echo ========================================================
echo   VOID- Sending App Local Compiler
echo ========================================================
echo.

:: 1. Clean up previous build artifacts to prevent errors
echo [1/4] Cleaning old build files...
if exist dist del /q dist\*.*
if exist build rd /s /q build
if exist "VOID- Sending App.spec" del /q "VOID- Sending App.spec"

:: 2. Ensure all required libraries are installed locally
echo [2/4] Verifying dependencies...
pip install -r requirements.txt
pip install pyinstaller

:: 3. Run the PyInstaller command
echo [3/4] Compiling VOID- Sending App.exe...
pyinstaller --noconsole --onefile --icon="New_void_icon.ico" --name "VOID- Sending App" ^
--collect-submodules xlwt ^
--hidden-import "xlwt" ^
--hidden-import "pandas" ^
--hidden-import "openpyxl" ^
--hidden-import "xlrd" ^
--hidden-import "PyPDF2" ^
--hidden-import "win32com" ^
--hidden-import "pythoncom" ^
--hidden-import "fitz" ^
--hidden-import "PIL" ^
main.py

:: 4. Finalizing
echo.
echo [4/4] Build Complete!
echo Your new executable is located in the 'dist' folder.
echo.
pause