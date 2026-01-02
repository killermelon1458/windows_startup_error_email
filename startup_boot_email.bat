@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM =====================================================
REM Portable startup_boot_email_scheduled.bat
REM - No internal delays
REM - Relies on Task Scheduler for timing/network
REM - Console + log output
REM =====================================================

REM --- Detect if we have a visible console
set "HAS_CONSOLE=1"
echo. >nul 2>&1 || set "HAS_CONSOLE=0"

REM --- Script directory
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "SCRIPT_NAME=startup_boot_email.py"
set "LOG=%SCRIPT_DIR%\boot_email.log"

REM --- Logging helper
call :log "===== BOOT EMAIL RUN ====="
call :log "Date: %DATE%"
call :log "Time: %TIME%"
call :log "Script directory: %SCRIPT_DIR%"

REM =====================================================
REM STEP 1: Verify Python script exists
REM =====================================================
if not exist "%SCRIPT_DIR%\%SCRIPT_NAME%" (
    call :log "[ERROR] Python script not found: %SCRIPT_NAME%"
    call :log "[ERROR] Expected location: %SCRIPT_DIR%"
    goto :fatal
)

call :log "[OK] Found %SCRIPT_NAME%"

REM =====================================================
REM STEP 2: Locate Python
REM =====================================================
set "PYTHON_EXE="

where python >nul 2>&1
if %ERRORLEVEL%==0 (
    for /f "delims=" %%P in ('where python') do (
        set "PYTHON_EXE=%%P"
        goto :python_found
    )
)

for %%P in (
    "%LocalAppData%\Programs\Python\Python*\python.exe"
    "%ProgramFiles%\Python*\python.exe"
    "%ProgramFiles(x86)%\Python*\python.exe"
) do (
    for %%F in (%%P) do (
        if exist "%%F" (
            set "PYTHON_EXE=%%F"
            goto :python_found
        )
    )
)

call :log "[ERROR] Python not found."
call :log "Install Python from:"
call :log "https://www.python.org/downloads/"
call :log "IMPORTANT: Check 'Add Python to PATH' during install."
goto :fatal

:python_found
call :log "[OK] Python detected: %PYTHON_EXE%"

REM =====================================================
REM STEP 3: Execute Python script
REM =====================================================
cd /d "%SCRIPT_DIR%"
set PYTHONIOENCODING=utf-8

call :log "[INFO] Running Python script"
"%PYTHON_EXE%" "%SCRIPT_DIR%\%SCRIPT_NAME%" >> "%LOG%" 2>&1

set "EXITCODE=%ERRORLEVEL%"
call :log "[INFO] Python exit code: %EXITCODE%"

if not "%EXITCODE%"=="0" goto :fatal

call :log "===== END RUN ====="
exit /b 0

REM =====================================================
REM ERROR HANDLING
REM =====================================================
:fatal
call :log "===== SCRIPT FAILED ====="
if %HAS_CONSOLE%==1 (
    echo.
    echo ERROR occurred. See log:
    echo   %LOG%
    echo.
    pause
)
exit /b 1

REM =====================================================
REM Logging function
REM =====================================================
:log
echo %~1
>>"%LOG%" echo %~1
exit /b
