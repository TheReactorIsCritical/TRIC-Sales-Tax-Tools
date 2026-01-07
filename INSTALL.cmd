@echo off
setlocal

REM Run the build script with ExecutionPolicy bypass (no global policy changes)
REM -NoProfile makes it faster and more predictable
REM -Force overwrites the existing add-in

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-addin.ps1" -Force

echo.
echo If you saw "Built add-in:", you're good.
echo Next: Excel -> File -> Options -> Add-ins -> Excel Add-ins -> Go... -> Browse -> select the .xlam
echo.
pause
