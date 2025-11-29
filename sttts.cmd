@echo off
rem This is a launcher batch file geared primariliy to check if PowerShell Core
rem is installed and to work around Windows' ExecutionPolicy scheme.

rem Secondarily, the window title is set while sttts is running, and if
rem something goes wrong on while running on a non-interactive shell session
rem (i.e. if script was double-clicked from Explorer/Start menu), a pause is
rem issued so the resulting errors can be examined.


setlocal

set "script_dir=%~dp0"
set "ESC=["

rem Identify what kind of shell session we are in:
set shell_is_interactive=1
set shell_is_pwsh=1
echo %PSModulePath% | findstr /L %USERPROFILE% >NUL
if %ERRORLEVEL% equ 0 goto :start

set shell_is_pwsh=0
rem Remove quotes for reliable IF statement comparison:
set "close_check=%CMDCMDLINE:""=%"
set "close_check=%close_check:cmd.exe /c =close_flag%"
set "close_check=%close_check:CMD.exe /C =close_flag%"

rem Check if the flag indicating a non-console session is present:
echo "%close_check%" | find "close_flag" >nul
if %ERRORLEVEL% neq 0 goto :start

set shell_is_interactive=0

:start
where pwsh 2>nul >nul
if %ERRORLEVEL% neq 0 goto :not_installed
pwsh -Version | findstr "Powershell 7" >nul
if %ERRORLEVEL% neq 0 goto :not_installed

if %shell_is_pwsh% equ 1 (
   for /f "tokens=*" %%f in ("pwsh -Command [console]::Title") do (
      set "saved_title=%%f"
   )
) else (
   if defined WT_SESSION (
      set "saved_title=Command Prompt"
   ) else (
      set "saved_title=%comspec%"
   )
)
title STTTS

rem "CALL" statement is a hack that avoids interrupt confirmation
pwsh -NoProfile -ExecutionPolicy Bypass -Command %script_dir%sttts.ps1 %* || call if EnsureError
set "err=%ERRORLEVEL%"
goto :pause

:not_installed
set "err=%ERRORLEVEL%"
echo:
echo %ESC%31mPowershell 7 was not found. For installation directions, go to:%ESC%m
echo %ESC%1;4;93mhttps://learn.microsoft.com/powershell/scripting/install/install-powershell-on-windows%ESC%m

rem Pause but only if not running from an interactive session:
:pause
if %shell_is_interactive% equ 1 goto :finally
echo %ESC%m
pause

:finally
if not defined err set "err=%ERRORLEVEL%"
if defined saved_title title %saved_title%
rem Reset text style before exiting:
<nul set /p="%ESC%m"
exit /b %err%
