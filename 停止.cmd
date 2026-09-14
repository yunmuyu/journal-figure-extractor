@echo off
setlocal EnableExtensions
cd /d "%~dp0"
if not exist "runtime\server.pid" (
  echo No server.pid found for v0.9. It may already be stopped.
  pause
  exit /b 0
)
set /p PID=<"runtime\server.pid"
if not defined PID (
  echo server.pid is empty.
  pause
  exit /b 1
)
echo Stopping v0.9 server PID %PID% ...
taskkill /PID %PID% /T /F
if errorlevel 1 (
  echo Could not stop that PID. It may have already exited.
) else (
  del /q "runtime\server.pid" >nul 2>&1
  del /q "runtime\server.json" >nul 2>&1
  echo Stopped.
)
pause
