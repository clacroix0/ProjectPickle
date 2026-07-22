@echo off
setlocal

set "LAUNCHER=%~dp0Project Pickle App Files\Run-ProjectPickle-32bit.bat"

if not exist "%LAUNCHER%" (
    echo Could not find the Project Pickle troubleshooting launcher.
    pause
    exit /b 1
)

call "%LAUNCHER%" %*
set "APP_EXIT=%ERRORLEVEL%"

endlocal
exit /b %APP_EXIT%
