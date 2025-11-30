@echo off
rem This is a launcher batch file geared primariliy to check if PowerShell Core
rem is installed and to work around Windows' ExecutionPolicy scheme.

rem Secondarily, the window title is set while sttts is running, and if
rem something goes wrong on while running on a non-interactive shell session
rem (i.e. if script was double-clicked from Explorer/Start menu), a pause is
rem issued so the resulting errors can be examined.


setlocal

if not defined STTTS_ARGS (
rem Add desired arguments to sttts.ps1 here:
   set "STTTS_ARGS="
)


set "script_dir=%~dp0"
set "ESC=["


rem Argument pass-through and for shortcut:
if "%STTTS_ARGS%" equ "" (
   if [%1] neq [] set "STTTS_ARGS=%*"
) else (
   if [%1] neq [] set "STTTS_ARGS=%STTTS_ARGS% %*"
)


rem Identify what kind of shell session we are in:
set shell_is_interactive=1
set shell_is_pwsh=1
echo %PSModulePath% | findstr /L %USERPROFILE% >NUL
if %ERRORLEVEL% equ 0 goto :check_installed

set shell_is_pwsh=0
rem Remove quotes for reliable IF statement comparison:
set "close_check=%CMDCMDLINE:""=%"
set "close_check=%close_check:cmd.exe /c =close_flag%"
set "close_check=%close_check:CMD.exe /C =close_flag%"

rem Check if the flag indicating a non-console session is present:
echo "%close_check%" | find "close_flag" >nul
if %ERRORLEVEL% neq 0 goto :check_installed

set shell_is_interactive=0


:check_installed
where pwsh 2>nul >nul
if %ERRORLEVEL% neq 0 goto :not_installed
pwsh -Version | findstr "Powershell 7" >nul
if %ERRORLEVEL% neq 0 goto :not_installed


rem Window title setting
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


rem Create shortcut if it doesn't exist:
if not exist %script_dir%STTTS.lnk (
   (
      echo $sh = New-Object -ComObject WScript.Shell
      echo $lnk = $sh.CreateShortcut('%script_dir%STTTS.lnk'^^^)
      echo $lnk.TargetPath = '%script_dir%sttts.cmd'
      echo $lnk.Arguments = '%STTTS_ARGS%'
      echo $lnk.WorkingDirectory = '%script_dir%'
      echo $lnk.IconLocation = '%script_dir%res\sttts.ico,0'
      echo $lnk.Save(^^^)
   ) | pwsh -NoProfile -ExecutionPolicy Bypass >nul
)


rem At last call main Powershell script
rem "CALL" statement is a hack that avoids interrupt confirmation:
pwsh -NoProfile -ExecutionPolicy Bypass -Command %script_dir%sttts.ps1 %STTTS_ARGS% || call if EnsureError
set "err=%ERRORLEVEL%"
goto :pause


:not_installed
set "err=%ERRORLEVEL%"
echo:
echo %ESC%31mPowershell 7 was not found. For installation directions, go to:%ESC%m
echo %ESC%1;4;93mhttps://learn.microsoft.com/powershell/scripting/install/install-powershell-on-windows%ESC%m


rem Pause but only if not running from an interactive session:
:pause
if %shell_is_interactive% neq 1 (
   echo %ESC%m
   pause
)

if not defined err set "err=%ERRORLEVEL%"
if defined saved_title title %saved_title%
rem Reset text style before exiting:
<nul set /p="%ESC%m"
exit /b %err%
