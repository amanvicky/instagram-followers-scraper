@echo off
rem One-click launcher for the Instagram Follower Comparator (Windows)
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found on PATH. Install Python 3.9+ from python.org first.
  pause
  exit /b 1
)
rem A venv folder can exist but be broken (e.g. after a Python update).
rem Verify its interpreter actually runs; rebuild when it does not.
if exist ".venv\Scripts\python.exe" (
  .venv\Scripts\python.exe --version >nul 2>nul
  if errorlevel 1 (
    echo The existing .venv is broken - rebuilding it...
    rmdir /s /q .venv
  )
)
if not exist ".venv\Scripts\python.exe" (
  echo First run - setting things up, this takes a couple of minutes...
  python -m venv .venv || goto :fail
  call .venv\Scripts\python.exe -m pip install --upgrade pip || goto :fail
  call .venv\Scripts\python.exe -m pip install -r requirements.txt || goto :fail
  call .venv\Scripts\python.exe -m playwright install chromium || goto :fail
  echo Setup complete.
)
echo Starting the app - your browser will open in a moment.
echo Keep this window open while you use the tool.
call .venv\Scripts\python.exe app.py
rem The app exiting (crash or Ctrl+C) lands here - keep the window open so
rem any error message stays readable instead of flashing and disappearing.
echo.
echo The app has stopped. If an error is printed above, paste it for support.
pause
goto :eof

:fail
echo.
echo Setup failed. Check the error above, then run this file again.
pause
exit /b 1
