@echo off
setlocal

REM === startup_boot_email_scheduled.bat ===
REM Quiet run for Task Scheduler (logs to ProgramData)

set "LOG=C:\ProgramData\uptime_logger\boot_email.log"
if not exist "C:\ProgramData\uptime_logger" mkdir "C:\ProgramData\uptime_logger"

> "%LOG%" echo -------- BOOT EMAIL RUN %DATE% %TIME% --------

set "PYTHON_EXE=C:\Users\Malachi Clifton\AppData\Local\Programs\Python\Python313\python.exe"
set "SCRIPT_DIR=C:\Users\Malachi Clifton\Documents\startup_boot_email"

REM Optional: if you keep pythonEmailNotify.py in a central place, set this.
REM If not needed, leave empty.
set "PYEMAILHELPER="

REM Give Windows time to bring up networking/event log
timeout /t 30 /nobreak >> "%LOG%" 2>&1

cd /d "%SCRIPT_DIR%"
set PYTHONIOENCODING=utf-8

if defined PYEMAILHELPER (
  "%PYTHON_EXE%" -c "import sys; sys.path.insert(0, r'%PYEMAILHELPER%'); import startup_boot_email as sbe; sbe.main()" >> "%LOG%" 2>&1
) else (
  "%PYTHON_EXE%" -c "import startup_boot_email as sbe; sbe.main()" >> "%LOG%" 2>&1
)

echo EXITCODE=%ERRORLEVEL% >> "%LOG%"
endlocal


