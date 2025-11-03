@echo off
setlocal
set "BASE=%~1"
if not defined BASE set "BASE=%CD%"

echo Starting Support Case Manager (Improved Version)...
echo Base Path: %BASE%

REM PowerShellウィンドウを非表示で実行
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" ^
  -WindowStyle Hidden ^
  -NoProfile ^
  -ExecutionPolicy Bypass ^
  -STA ^
  -File "%~dp0SupportCaseManager.ps1" -BasePath "%BASE%"

REM エラーレベルの確認（PowerShellウィンドウが非表示のため）
if %ERRORLEVEL% neq 0 (
    echo.
    echo Error occurred while starting the application.
    echo Please check the log file for details.
    echo Press any key to continue...
    pause >nul
)

