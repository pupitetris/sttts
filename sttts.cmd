@echo off
rem This is a launcher batch file geared primariliy to work around Windows' ExecutionPolicy

setlocal

title STTTS

set "script_dir=%~dp0"
set "ESC=["

where pwsh 2>nul >nul
if %ERRORLEVEL% neq 0 goto :not_installed
pwsh -Version | findstr "Powershell 7" >nul
if %ERRORLEVEL% neq 0 goto :not_installed

rem "CALL" statement is a hack that avoids interrupt confirmation
pwsh -NoProfile -ExecutionPolicy Bypass -Command %script_dir%sttts.ps1 %* || call if EnsureError
set "err=%ERRORLEVEL%"
goto :pause

:not_installed
set "err=%ERRORLEVEL%"
echo:
echo %ESC%31mPowershell 7 was not found. For installation directions, go to:%ESC%m
echo %ESC%1;4;93mhttps://learn.microsoft.com/powershell/scripting/install/install-powershell-on-windows%ESC%m

rem Pause but only if not running from an interactive prompt:
:pause

rem No pause if this is a PowerShell session:
echo %PSModulePath% | findstr /L %USERPROFILE% >NUL
if %ERRORLEVEL% equ 0 goto :finally

rem Remove quotes for reliable IF statement comparison:
set "close_check=%CMDCMDLINE:""=%"
set "close_check=%close_check:cmd.exe /c =close_flag%"
set "close_check=%close_check:CMD.exe /C =close_flag%"

rem Check if the flag indicating a non-console session is present:
echo "%close_check%" | find "close_flag" >nul
if %ERRORLEVEL% neq 0 goto :finally
echo %ESC%m
pause

:finally
if not defined err set "err=%ERRORLEVEL%"
rem Reset text style before exiting:
<nul set /p="%ESC%m"
exit /b %err%
