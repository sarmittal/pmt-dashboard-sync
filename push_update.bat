@echo off
:: ============================================================
:: PMT Dashboard — Auto Push Update
:: 
:: Usage:
::   push_update.bat "your commit message"
::   push_update.bat  (uses default message with timestamp)
::
:: Place this file in the ROOT of your repo:
::   C:\path\to\pmt-dashboard-sync\push_update.bat
:: ============================================================

setlocal

:: ── Config ───────────────────────────────────────────────────
set BRANCH=claude/fix-attachment-link-routing-5kCBB
set SRC_DIR=%USERPROFILE%\Downloads
set DEST_FILE=client\src\App.jsx
:: ─────────────────────────────────────────────────────────────

:: Get commit message — use arg if provided, else use timestamp
if "%~1"=="" (
    for /f "tokens=1-3 delims=/ " %%a in ("%DATE%") do set DATESTAMP=%%c-%%a-%%b
    for /f "tokens=1-2 delims=: " %%a in ("%TIME%") do set TIMESTAMP=%%a%%b
    set MSG=feat: dashboard update %DATESTAMP% %TIMESTAMP%
) else (
    set MSG=%~1
)

echo.
echo ╔══════════════════════════════════════════════╗
echo ║       PMT Dashboard — Push Update            ║
echo ╚══════════════════════════════════════════════╝
echo.

:: ── Step 1: Look for App.jsx in Downloads ────────────────────
if exist "%SRC_DIR%\App.jsx" (
    echo [1/5] Found App.jsx in Downloads folder
    echo       Copying to %DEST_FILE%...
    copy /Y "%SRC_DIR%\App.jsx" "%DEST_FILE%" >nul
    if errorlevel 1 (
        echo ERROR: Could not copy App.jsx
        pause
        exit /b 1
    )
    echo       Done.
) else (
    echo [1/5] No App.jsx found in Downloads — using existing file
)

:: ── Step 2: Make sure we're on the right branch ──────────────
echo [2/5] Switching to branch: %BRANCH%
git checkout %BRANCH% 2>nul
if errorlevel 1 (
    echo ERROR: Could not switch to branch %BRANCH%
    echo        Make sure you have pulled this branch first:
    echo        git pull origin %BRANCH%
    pause
    exit /b 1
)

:: ── Step 3: Pull latest ──────────────────────────────────────
echo [3/5] Pulling latest from remote...
git pull origin %BRANCH% --quiet
if errorlevel 1 (
    echo WARNING: Pull had issues — continuing anyway
)

:: ── Step 4: Stage and commit ─────────────────────────────────
echo [4/5] Staging changes...
git add client\src\App.jsx

git diff --cached --quiet
if errorlevel 1 (
    echo       Committing: %MSG%
    git commit -m "%MSG%"
) else (
    echo       No changes to commit — file is already up to date.
    goto :done
)

:: ── Step 5: Push ─────────────────────────────────────────────
echo [5/5] Pushing to GitHub...
git push origin %BRANCH%
if errorlevel 1 (
    echo ERROR: Push failed. Check your credentials or network.
    pause
    exit /b 1
)

:done
echo.
echo ✅ Done! Changes pushed to %BRANCH%
echo.
echo Next steps:
echo   1. git pull origin %BRANCH%   (on localhost to test)
echo   2. npm run dev                 (test locally)
echo   3. Deploy to BTP when ready
echo.
pause
