@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "APP_URL=http://127.0.0.1:8502"
set "APP_PORT=8502"

if /I "%~1"=="--minimized" goto :main
start "Market Report" /min "%~f0" --minimized
exit /b 0

:main
call :is_ready
if not errorlevel 1 goto :open_browser

start "Market Report" /min /d "%~dp0" python server.py

set /a _n=0
:wait_ready
call :is_ready
if not errorlevel 1 goto :open_browser
set /a _n+=1
if %_n% GEQ 45 goto :open_browser
timeout /t 1 /nobreak >nul
goto :wait_ready

:open_browser
start "" "%APP_URL%"
exit /b 0

:is_listening
netstat -ano | findstr /R /C:"127.0.0.1:%APP_PORT% .*LISTENING" >nul
if not errorlevel 1 exit /b 0
netstat -ano | findstr /C:":%APP_PORT% " | findstr /C:"LISTENING" >nul
exit /b %ERRORLEVEL%

:is_ready
call :is_listening
if errorlevel 1 exit /b 1
curl.exe -s -o nul -m 3 "%APP_URL%/health"
exit /b %ERRORLEVEL%
