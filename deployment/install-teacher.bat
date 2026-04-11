@echo off
setlocal

REM ============================================================================
REM  Install exam shortcuts - TEACHER PC
REM  Creates 4 shortcuts: Examiner, Teacher, Examinee, Student (practice)
REM ============================================================================

echo.
echo ================================================================
echo    Exam System Shortcuts - TEACHER PC Installation
echo ================================================================
echo.
echo Shortcuts to create: Examiner, Teacher, Examinee, Student
echo.

REM Check if running as admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo NOTE: Not running as Administrator.
    echo Shortcuts will be created for the current user only.
    echo To create them for ALL users, right-click this file
    echo and choose "Run as administrator".
    echo.
    timeout /t 3 >nul
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-ExamShortcuts.ps1" -Role Teacher

if %errorLevel% neq 0 (
    echo.
    echo An error occurred during installation. See output above.
    pause
    exit /b 1
)

echo.
echo Installation complete. You can close this window.
pause
exit /b 0
