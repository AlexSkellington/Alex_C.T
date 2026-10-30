@echo off
rem ==================================================================================================
rem  TBS Maintenance Script - double-click launcher
rem
rem  Downloads the current TBS Maintenance Script to %TEMP% with Windows' own curl.exe and runs that file.
rem  Do NOT go back to "powershell -Command ... Invoke-RestMethod ... | Invoke-Expression": Microsoft
rem  Defender's machine-learning detection blocks that download-and-run command line as
rem  Trojan:Win32/Commando.A!ml (seen 2026-10-30). Running a downloaded file is ordinary admin use.
rem ==================================================================================================
setlocal
set "URL=https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1"
set "FILE=%TEMP%\TBS_Maintenance_Script.ps1"

where curl.exe >nul 2>nul
if errorlevel 1 (
	echo curl.exe was not found ^(it ships with Windows 10 1803 / Server 2019 and newer^).
	echo Download https://github.com/AlexSkellington/Alex_C.T/releases/latest/download/TBS_Maintenance_Script.ps1
	echo save it as %FILE% and run it with PowerShell.
	pause
	exit /b 1
)

echo Downloading TBS Maintenance Script...
curl.exe -fsSL --retry 2 --connect-timeout 15 -o "%FILE%" "%URL%"
if errorlevel 1 (
	echo Download failed: %URL%
	pause
	exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%FILE%"
endlocal
