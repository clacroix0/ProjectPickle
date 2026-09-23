@echo off
setlocal

set "SCRIPT=%~dp0ProjectPickle.ps1"
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
for %%I in ("%~dp0..") do set "APPROOT=%%~fI"
set "CRASHLOG=%APPROOT%\ProjectPickle.LastError.txt"
set "DIAGLOG=%APPROOT%\ProjectPickle.Diagnostics.txt"

if not exist "%PS32%" (
    set "PS32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

if not exist "%SCRIPT%" (
    echo Could not find the Project Pickle app file.
    pause
    exit /b 1
)

if exist "%CRASHLOG%" del /q "%CRASHLOG%" >nul 2>nul
if exist "%DIAGLOG%" del /q "%DIAGLOG%" >nul 2>nul

"%PS32%" -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "APP_EXIT=%ERRORLEVEL%"

if not "%APP_EXIT%"=="0" (
    echo.
    echo Project Pickle closed because of an error.
    echo The error message above should explain what happened.
    if exist "%CRASHLOG%" (
        echo.
        echo Crash log:
        echo %CRASHLOG%
        echo.
        type "%CRASHLOG%"
    )
    if not exist "%CRASHLOG%" if exist "%DIAGLOG%" (
        echo.
        echo Diagnostic log:
        echo %DIAGLOG%
        echo.
        type "%DIAGLOG%"
    )
    echo.
    pause
)

endlocal
exit /b %APP_EXIT%
