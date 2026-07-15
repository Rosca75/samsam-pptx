@echo off
setlocal
rem ============================================================================
rem  SamSam Tools build script
rem
rem  Expects: you saved the VBA project from PowerPoint as
rem      %AppData%\Microsoft\AddIns\SamSamTools_src.ppam
rem  (PowerPoint ALWAYS writes .ppam files there, silently ignoring the folder
rem   you picked in the Save As dialog — that is normal.)
rem
rem  Produces: %AppData%\Microsoft\AddIns\SamSamTools.ppam  (ribbon injected)
rem
rem  Requirements: Python 3 on PATH. Standard library only — no pip, no network.
rem  IMPORTANT: close ALL PowerPoint windows before running this script; the
rem  registered add-in file is locked while PowerPoint runs, and the ribbon is
rem  cached until a full restart anyway.
rem ============================================================================

set "ADDINDIR=%AppData%\Microsoft\AddIns"
set "SRC_PPAM=%ADDINDIR%\SamSamTools_src.ppam"
set "OUT_PPAM=%ADDINDIR%\SamSamTools.ppam"
set "RIBBON=%~dp0..\ribbon\customUI14.xml"

if not exist "%SRC_PPAM%" (
    echo [build] ERROR: %SRC_PPAM% not found.
    echo [build] In PowerPoint: File - Save As - PowerPoint Add-in ^(*.ppam^),
    echo [build] filename SamSamTools_src . PowerPoint saves it to %ADDINDIR%
    echo [build] automatically, whatever folder the dialog shows.
    exit /b 1
)

if not exist "%RIBBON%" (
    echo [build] ERROR: ribbon XML not found at %RIBBON%
    exit /b 1
)

if exist "%OUT_PPAM%" (
    echo [build] Deleting old %OUT_PPAM%
    del /f "%OUT_PPAM%"
    if exist "%OUT_PPAM%" (
        echo [build] ERROR: could not delete the old add-in. Close ALL
        echo [build] PowerPoint windows and run this script again.
        exit /b 1
    )
)

where python >nul 2>nul
if errorlevel 1 (
    echo [build] ERROR: python not found on PATH. Install Python 3 ^(no admin
    echo [build] rights needed for the per-user installer^) and retry.
    exit /b 1
)

python "%~dp0inject_ribbon.py" "%RIBBON%" "%SRC_PPAM%" "%OUT_PPAM%"
if errorlevel 1 (
    echo [build] ERROR: ribbon injection failed. Nothing was installed.
    exit /b 1
)

echo.
echo [build] DONE: %OUT_PPAM%
echo [build] ============================================================
echo [build]  RESTART REQUIRED: close and reopen ALL PowerPoint windows.
echo [build]  PowerPoint caches the ribbon - the new tab will not appear
echo [build]  (or the old one will linger) until a full restart.
echo [build] ============================================================
endlocal
