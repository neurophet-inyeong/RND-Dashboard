@echo off
cd /d "%~dp0"

:: Check if port 5500 is already in use
netstat -an 2>nul | find "5500" | find "LISTENING" > nul
if %errorlevel% == 0 (
    echo [RnD Dashboard] Port 5500 is already open. Opening browser.
    start "" "http://localhost:5500/dashboard.html"
    goto :end
)

:: Check that the py launcher is available
where py > nul 2>&1
if not %errorlevel% == 0 (
    echo [RnD Dashboard] ERROR: Python launcher py.exe was not found.
    echo Please check that Python is installed and registered in PATH.
    pause
    goto :end
)

:: Start the HTTP server in a new window
start "RnD Dashboard Server" /min py -m http.server 5500

:: Wait up to 5 seconds for the server to come up
set /a tries=0
:waitloop
timeout /t 1 /nobreak > nul
set /a tries+=1
netstat -an 2>nul | find "5500" | find "LISTENING" > nul
if %errorlevel% == 0 goto :serverready
if %tries% lss 5 goto :waitloop

echo [RnD Dashboard] ERROR: server did not start within 5 seconds.
echo Please check that running "py -m http.server 5500" works.
pause
goto :end

:serverready
:: Open the browser
start "" "http://localhost:5500/dashboard.html"

echo.
echo [RnD Dashboard] Server is running.
echo Closing this window (RnD Dashboard Server) will stop the server.

:end
