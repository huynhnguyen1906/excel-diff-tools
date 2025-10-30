@echo off
rem Build script for excel-sheet-diff (Windows)
rem Usage: double-click build.bat or run from cmd

setlocal enabledelayedexpansion

rem Resolve script directory
set ROOT_DIR=%~dp0
cd /d "%ROOT_DIR%"

echo ================================
echo Building excel-sheet-diff
echo Directory: %ROOT_DIR%
echo ================================

rem Optional: pass "clean" to remove previous build/dist folders
if /I "%1"=="clean" (
    echo Cleaning previous build artifacts...
    if exist "%ROOT_DIR%build" rmdir /s /q "%ROOT_DIR%build"
    if exist "%ROOT_DIR%dist" rmdir /s /q "%ROOT_DIR%dist"
    if exist "%ROOT_DIR%excel_diff.spec" (
        rem leave spec file, remove *.spec-related files in root if any
    )
)

rem Use virtualenv python if available, otherwise fallback to system python
set VENV_PY=%ROOT_DIR%venv\Scripts\python.exe
if exist "%VENV_PY%" (
    set "PYEXEC=%VENV_PY%"
) else (
    set "PYEXEC=python"
)

echo Using python: %PYEXEC%

rem Run PyInstaller with the spec file
%PYEXEC% -m PyInstaller "%ROOT_DIR%excel_diff.spec"
if errorlevel 1 (
    echo.
    echo Build failed with exit code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Build finished successfully.
for /f "delims=" %%I in ('dir /b /ad "%ROOT_DIR%dist" 2^>nul') do (
    rem show the dist folder contents
)
echo Output dir: %ROOT_DIR%dist
echo.
pause
