@echo off
:: ============================================================
:: PMT Dashboard — Watch & Auto-Push
::
:: Watches your Downloads folder for a new App.jsx
:: When detected, automatically copies + commits + pushes
::
:: Usage: Just double-click this file and leave it running.
::        Whenever Claude gives you a new App.jsx and you save
::        it to Downloads, it will auto-push within 5 seconds.
::
:: Place this file in the ROOT of your repo.
:: ============================================================

setlocal

set BRANCH=claude/fix-attachment-link-routing-5kCBB
set WATCH_FILE=%USERPROFILE%\Downloads\App.jsx
set DEST_FILE=client\src\App.jsx
set LAST_HASH=

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║     PMT Dashboard — Auto Watch ^& Push                   ║
echo ║                                                          ║
echo ║  Watching: %USERPROFILE%\Downloads\App.jsx
echo ║  Branch:   %BRANCH%
echo ║                                                          ║
echo ║  Drop a new App.jsx into Downloads and it will          ║
echo ║  auto-commit and push within 5 seconds.                 ║
echo ║                                                          ║
echo ║  Press Ctrl+C to stop watching.                         ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

:: Make sure we're on the right branch
git checkout %BRANCH% >nul 2>&1

:loop
    :: Check if App.jsx exists in Downloads
    if not exist "%WATCH_FILE%" (
        goto :wait
    )

    :: Get hash of the Downloads file
    for /f %%h in ('certutil -hashfile "%WATCH_FILE%" MD5 ^| findstr /v ":"') do set CURR_HASH=%%h

    :: Compare with last known hash
    if "%CURR_HASH%"=="%LAST_HASH%" goto :wait

    :: Hash changed — new file detected!
    set LAST_HASH=%CURR_HASH%

    echo [%TIME%] New App.jsx detected — deploying...

    :: Copy to repo
    copy /Y "%WATCH_FILE%" "%DEST_FILE%" >nul

    :: Pull first to avoid conflicts
    git pull origin %BRANCH% --quiet >nul 2>&1

    :: Stage
    git add client\src\App.jsx

    :: Check if there are actual changes
    git diff --cached --quiet
    if errorlevel 1 (
        :: Build commit message with timestamp
        for /f "tokens=1-3 delims=/ " %%a in ("%DATE%") do set DS=%%c-%%a-%%b
        git commit -m "feat: dashboard update %DS% %TIME:~0,5%" >nul
        git push origin %BRANCH% >nul 2>&1
        if errorlevel 1 (
            echo [%TIME%] ❌ Push failed — check credentials
        ) else (
            echo [%TIME%] ✅ Pushed! Now run: git pull origin %BRANCH%
        )
    ) else (
        echo [%TIME%] No changes detected in App.jsx
    )

:wait
    timeout /t 5 /nobreak >nul
    goto :loop
