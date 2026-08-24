@echo off
rem ResumeForge launcher: prefer windowless pythonw, fall back to python.
cd /d "%~dp0"
set "SCRIPT=%~dp0resume_converter.py"
set "LAUNCH="
where pythonw >nul 2>nul && set "LAUNCH=pythonw"
if not defined LAUNCH where python >nul 2>nul && set "LAUNCH=python"
if defined LAUNCH (
    start "ResumeForge" %LAUNCH% "%SCRIPT%"
) else (
    echo [ERROR] Python not found. Install Python and enable "Add to PATH", then retry.
    pause
)
