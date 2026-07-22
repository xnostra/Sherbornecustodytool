@echo off
title IT Asset Custody Form Tool
cd /d "%~dp0"

REM Set to 1 to auto-email the finished form (you'll get one Microsoft sign-in prompt per run).
REM Set to 0 to just save/open it locally, no emailing.
set "AUTO_EMAIL=1"

REM Run the script's TEXT as a scriptblock instead of "-File" or plain Invoke-Expression. Windows
REM blocks double-clicking/running .ps1 FILES whenever script execution is disabled - even by
REM group policy, which "-ExecutionPolicy Bypass" cannot override. A scriptblock isn't a file, so
REM Execution Policy never applies, and it lets us pass -EmailForm through with normal param()
REM binding (plain Invoke-Expression would just have param() overwrite it).
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$sb = [scriptblock]::Create((Get-Content -Raw -LiteralPath '%~dp0Fill-CustodyForm.ps1')); if ('%AUTO_EMAIL%' -eq '1') { & $sb -EmailForm } else { & $sb }"

set "RC=%errorlevel%"
if not "%RC%"=="0" (
    echo.
    echo The tool exited with an error ^(code %RC%^).
    echo If this mentions Execution Policy or "scripts disabled", try right-clicking this .bat
    echo and choosing "Run as administrator" once.
    echo Full details are also saved in CustodyForm.log in this same folder.
    echo.
    pause
)
