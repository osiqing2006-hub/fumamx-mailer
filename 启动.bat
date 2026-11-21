@echo off
chcp 65001 >nul
setlocal
set "DIR=%~dp0"

rem Prefer the py launcher, then pythonw / python on PATH.
where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw -3 "%DIR%app.py"
    goto :eof
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%DIR%app.py"
    goto :eof
)
where python >nul 2>nul
if %errorlevel%==0 (
    python "%DIR%app.py"
    goto :eof
)

echo.
echo   Python was not found.
echo   Install Python 3.10 or newer from https://www.python.org/downloads/
echo   During setup, tick "Add python.exe to PATH".
echo.
pause
