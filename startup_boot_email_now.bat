@echo off
REM === startup_boot_email_now.bat ===
REM Run the Python boot email script right now (visible, pauses at end)

set "PYTHON_EXE=C:\Users\Malachi Clifton\AppData\Local\Programs\Python\Python313\python.exe"
set "SCRIPT=C:\Users\Malachi Clifton\Documents\startup_boot_email\startup_boot_email.py"

echo [%DATE% %TIME%] Running %SCRIPT% with %PYTHON_EXE% ...
"%PYTHON_EXE%" -X dev "%SCRIPT%"
echo.
echo Done. Press any key to close.
pause
