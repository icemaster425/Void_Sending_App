@echo off
echo ========================================================
echo   VOID- Sending App Workspace Architect (AWS Edition)
echo ========================================================
echo.

:: Define the absolute source paths
set "target_exe=%~dp0VOID- Sending App.exe"
set "target_folder=%~dp0To Send"

:: 1. Build EXE Shortcut
echo [1/3] Hunting Desktop path and forging App shortcut...
powershell -NoProfile -Command "$desktop = [Environment]::GetFolderPath('Desktop'); $s = (New-Object -COM WScript.Shell).CreateShortcut($desktop + '\VOID- Sending App.lnk'); $s.TargetPath = '%target_exe%'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = '%target_exe%'; $s.Save()"

:: 2. Build Folder Shortcut
echo [2/3] Forging 'To Send' folder shortcut...
if not exist "%target_folder%" mkdir "%target_folder%"
powershell -NoProfile -Command "$desktop = [Environment]::GetFolderPath('Desktop'); $s = (New-Object -COM WScript.Shell).CreateShortcut($desktop + '\VOID- To Send.lnk'); $s.TargetPath = '%target_folder%'; $s.Save()"

:: 3. The Taskbar Pin
echo [3/3] Executing Taskbar Pin...
powershell -NoProfile -Command "$shell = New-Object -ComObject Shell.Application; $folder = $shell.Namespace((Get-Item '%target_exe%').DirectoryName); $item = $folder.ParseName((Get-Item '%target_exe%').Name); $item.InvokeVerb('taskbarpin')"

echo.
echo Workspace primed. 
echo.
pause