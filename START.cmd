@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Journal Figure Extractor v1.4

echo ============================================================
echo Journal Figure Extractor v1.4
echo ============================================================
echo.

set "PYEXE="

rem 1) Prefer an existing venv if it can run bootstrap.
if exist ".venv\Scripts\python.exe" set "PYEXE=%CD%\.venv\Scripts\python.exe"

rem 2) Known WorkBuddy runtime on this machine.
if not defined PYEXE if exist "%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe" set "PYEXE=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"

rem 3) Normal Python on PATH.
if not defined PYEXE for /f "delims=" %%P in ('where python.exe 2^>nul') do if not defined PYEXE set "PYEXE=%%P"

rem 4) Python launcher.
if not defined PYEXE for /f "delims=" %%P in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do if not defined PYEXE set "PYEXE=%%P"

rem 5) Last-resort scan under WorkBuddy.
if not defined PYEXE if exist "%USERPROFILE%\.workbuddy\binaries" (
  for /r "%USERPROFILE%\.workbuddy\binaries" %%P in (python.exe) do if not defined PYEXE set "PYEXE=%%P"
)

if not defined PYEXE (
  echo ERROR: No Python runtime was found.
  echo Checked PATH, py launcher, and WorkBuddy runtimes.
  echo.
  pause
  exit /b 1
)

echo Python candidate: %PYEXE%
"%PYEXE%" "%CD%\bootstrap_progress.py"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo START FAILED. Error code: %RC%
  echo Send logs\startup.log to me.
  echo.
  pause
)
exit /b %RC%