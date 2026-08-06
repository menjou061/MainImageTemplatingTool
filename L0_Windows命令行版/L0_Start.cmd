@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Main Image Templating Tool 1.4

set "BASE=%~dp0"
set "RUNNER=%BASE%_internal\L0_Run.bat"
if defined LOCALAPPDATA (
  set "STARTUP_DIR=%LOCALAPPDATA%\MainImageTemplatingTool\startup"
) else (
  set "STARTUP_DIR=%TEMP%\MainImageTemplatingTool\startup"
)
set "LOG=%STARTUP_DIR%\start.log"
set "RUNNER_CONSOLE=%STARTUP_DIR%\runner_stdout_stderr.log"

if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%" >nul 2>nul
if exist "%STARTUP_DIR%\failure_dialog_shown.marker" del /q "%STARTUP_DIR%\failure_dialog_shown.marker" >nul 2>nul
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=entered file=L0_Start.cmd
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=path_resolved base="%BASE%" runner="%RUNNER%"

if not exist "%RUNNER%" (
  >> "%LOG%" echo [%DATE% %TIME%] error=missing_runner path="%RUNNER%"
  > "%STARTUP_DIR%\failure.txt" echo Main Image Templating Tool startup failed
  >> "%STARTUP_DIR%\failure.txt" echo stage=entry
  >> "%STARTUP_DIR%\failure.txt" echo reason=missing _internal\L0_Run.bat
  call :SHOW_STARTUP_FAILURE
  exit /b 1
)

rem Hide the launcher console after the entrypoint has been validated. Set
rem L0_SHOW_CONSOLE=1 before starting the tool to keep it visible for support.
if /I not "%L0_SHOW_CONSOLE%"=="1" if exist "%BASE%_internal\hide_console.ps1" (
  set "HIDE_PS_EXE="
  if defined PROCESSOR_ARCHITEW6432 if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "HIDE_PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
  if not defined HIDE_PS_EXE if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set "HIDE_PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
  if defined HIDE_PS_EXE "%HIDE_PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%BASE%_internal\hide_console.ps1" >nul 2>&1
)

set "L0_ENTRY_OK=1"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=before_runner archive="%RUNNER_CONSOLE%"
call "%RUNNER%" >> "%RUNNER_CONSOLE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=runner_returned exitCode=%EXIT_CODE%
if not "%EXIT_CODE%"=="0" if not exist "%STARTUP_DIR%\failure_dialog_shown.marker" call :SHOW_STARTUP_FAILURE

exit /b %EXIT_CODE%

:SHOW_STARTUP_FAILURE
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" goto :EOF
"%PS_EXE%" -STA -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show('The tool could not start. Open the MainImageTemplatingTool startup log for details.','Main Image Templating Tool',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)" >nul 2>&1
goto :EOF
