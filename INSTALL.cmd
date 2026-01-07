@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%build-addin.ps1"


echo --------------------------------------------
echo INSTALLING THE ADDIN
echo --------------------------------------------
echo.
echo Loading source files...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Force
set "EXITCODE=%ERRORLEVEL%"

echo.
if %EXITCODE% neq 0 goto :fail

echo --------------------------------------------
echo INSTALL COMPLETE
echo --------------------------------------------
echo.
pause
exit /b 0

:fail
echo --------------------------------------------
echo INSTALL FAILED - exit code %EXITCODE%
echo --------------------------------------------
echo.
pause
exit /b %EXITCODE%
