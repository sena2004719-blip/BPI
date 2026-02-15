@echo off
setlocal EnableExtensions

REM ==============================
REM  BPI (BOAT PRO INDEX) START
REM  - Run update now
REM  - Open viewer
REM  - (Optional) auto-create daily task
REM  IMPORTANT: Always run from an extracted folder (not inside ZIP)
REM ==============================

REM Move to this folder
cd /d "%~dp0"

REM Guard: if running inside ZIP path
echo %cd% | findstr /i "\.zip\\" >nul
if %errorlevel%==0 (
  echo [ERROR] You are running inside a ZIP. Please Right-click the ZIP ^> "Extract All" then run START.bat.
  pause
  exit /b 1
)

REM Ensure folders
if not exist "data" mkdir "data" >nul 2>nul
if not exist "logs" mkdir "logs" >nul 2>nul

REM Run updater (always writes data\status.json)
echo.
echo === BPI updater start ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0updater.ps1"

REM Start viewer server (keeps window open; closing stops viewer)
echo.
echo === Opening viewer ===
start "BPI_VIEWER" /min powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve.ps1"

REM Auto-create daily scheduled task (best effort)
REM Runs updater only (no viewer) every day at 02:10
set TASKNAME=BPI_DailyUpdate
schtasks /Query /TN "%TASKNAME%" >nul 2>nul
if not %errorlevel%==0 (
  echo.
  echo === Setting up daily auto-update (02:10) ===
  schtasks /Create /F /SC DAILY /ST 02:10 /TN "%TASKNAME%" /TR "powershell -NoProfile -ExecutionPolicy Bypass -File \"%~dp0updater.ps1\"" >nul 2>nul
  if %errorlevel%==0 (
    echo [OK] Daily task created: %TASKNAME%
  ) else (
    echo [WARN] Could not create scheduled task (you can still run START.bat anytime).
  )
)

echo.
echo Done. You can close this window anytime.
echo (Auto-update task name: BPI_DailyUpdate)
pause
