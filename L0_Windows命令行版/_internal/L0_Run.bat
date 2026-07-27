@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Main Image Templating Tool runner

set "INTERNAL_DIR=%~dp0"
if defined LOCALAPPDATA (
  set "STARTUP_DIR=%LOCALAPPDATA%\MainImageTemplatingTool\startup"
) else (
  set "STARTUP_DIR=%TEMP%\MainImageTemplatingTool\startup"
)
set "LOG=%STARTUP_DIR%\start.log"
set "PS_SCRIPT=%INTERNAL_DIR%L0_Run.ps1"
set "PS_CONSOLE=%STARTUP_DIR%\powershell_stdout_stderr.log"
set "PS_EXE="

if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%" >nul 2>nul
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=entered file=_internal\L0_Run.bat
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=path_resolved internal="%INTERNAL_DIR%" ps1="%PS_SCRIPT%"

if not "%L0_ENTRY_OK%"=="1" (
  call :WRITE_FAILURE "entry" "run L0_Start.cmd from the parent folder"
  exit /b 1
)

if not exist "%PS_SCRIPT%" (
  call :WRITE_FAILURE "runner" "missing _internal\L0_Run.ps1"
  exit /b 1
)

call :RESOLVE_POWERSHELL
if not defined PS_EXE (
  call :WRITE_FAILURE "powershell" "Windows PowerShell 5.1 executable was not found"
  exit /b 1
)

>> "%LOG%" echo [%DATE% %TIME%] checkpoint=before_powershell exe="%PS_EXE%" archive="%PS_CONSOLE%"
"%PS_EXE%" -STA -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" > "%PS_CONSOLE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=powershell_returned exitCode=%EXIT_CODE%
exit /b %EXIT_CODE%

:RESOLVE_POWERSHELL
if defined PROCESSOR_ARCHITEW6432 if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE if exist "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE for /f "usebackq delims=" %%P in (`where powershell.exe 2^>nul`) do if not defined PS_EXE set "PS_EXE=%%P"
goto :EOF

:WRITE_FAILURE
set "FAIL_STAGE=%~1"
set "FAIL_REASON=%~2"
if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%" >nul 2>nul
> "%STARTUP_DIR%\failure.txt" echo Main Image Templating Tool startup failed
>> "%STARTUP_DIR%\failure.txt" echo stage=%FAIL_STAGE%
>> "%STARTUP_DIR%\failure.txt" echo reason=%FAIL_REASON%
>> "%STARTUP_DIR%\failure.txt" echo psExe=%PS_EXE%
>> "%STARTUP_DIR%\failure.txt" echo psScript=%PS_SCRIPT%
>> "%LOG%" echo [%DATE% %TIME%] startup_failure stage="%FAIL_STAGE%" reason="%FAIL_REASON%"
goto :EOF
