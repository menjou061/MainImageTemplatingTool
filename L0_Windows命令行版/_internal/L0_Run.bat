@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul 2>nul
title L0 internal runner

set "INTERNAL_DIR=%~dp0"
for %%I in ("%INTERNAL_DIR%..") do set "BASE=%%~fI\"
set "DIAG=%BASE%_diagnostics"
set "LOG=%DIAG%\start.log"
set "PS_SCRIPT=%INTERNAL_DIR%L0_Run.ps1"
set "PS1_MARKER=%DIAG%\ps1_started.marker"
set "STATUS_JSON=%DIAG%\status.json"
set "FAILURE_TXT=%DIAG%\failure.txt"
set "FAILURE_CSV=%DIAG%\failure.csv"
set "PS_CONSOLE=%DIAG%\powershell_stdout_stderr.log"
set "PS_EXE="

if not exist "%DIAG%" mkdir "%DIAG%" >nul 2>nul
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=entered file=_internal\L0_Run.bat
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=path resolved internal="%INTERNAL_DIR%" base="%BASE%" ps1="%PS_SCRIPT%" diag="%DIAG%"

if not "%L0_ENTRY_OK%"=="1" (
  >> "%LOG%" echo [%DATE% %TIME%] error=direct_internal_launch L0_ENTRY_OK="%L0_ENTRY_OK%"
  call :WRITE_FALLBACK "entry" "run L0_Start.cmd from the parent folder instead of _internal\L0_Run.bat" "run L0_Start.cmd from the parent folder instead of _internal\L0_Run.bat" "1"
  exit /b 1
)

if not exist "%PS_SCRIPT%" (
  >> "%LOG%" echo [%DATE% %TIME%] error=missing_ps1 path="%PS_SCRIPT%"
  call :WRITE_FALLBACK "bat" "missing _internal\L0_Run.ps1" "missing _internal\L0_Run.ps1" "1"
  exit /b 1
)

call :RESOLVE_POWERSHELL
if not defined PS_EXE (
  >> "%LOG%" echo [%DATE% %TIME%] error=powershell_not_found SystemRoot="%SystemRoot%" PATH="%PATH%"
  call :WRITE_FALLBACK "powershell" "Windows PowerShell 5.1 executable was not found" "Windows PowerShell 5.1 executable was not found" "1"
  exit /b 1
)

>> "%LOG%" echo [%DATE% %TIME%] checkpoint=before powershell exe="%PS_EXE%" ps1="%PS_SCRIPT%" archive="%PS_CONSOLE%"
"%PS_EXE%" -STA -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" > "%PS_CONSOLE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=powershell returned exitCode=%EXIT_CODE%

if not "%EXIT_CODE%"=="0" (
  if not exist "%PS1_MARKER%" (
    call :WRITE_FALLBACK "powershell" "PowerShell exited before PS1 startup marker; exitCode=%EXIT_CODE%" "PowerShell exited before PS1 startup marker; exitCode=%EXIT_CODE%" "%EXIT_CODE%"
  ) else (
    if not exist "%STATUS_JSON%" if not exist "%FAILURE_TXT%" if not exist "%FAILURE_CSV%" (
      call :WRITE_FALLBACK "ps1" "PS1 exited without diagnostics; exitCode=%EXIT_CODE%" "PS1 exited without diagnostics; exitCode=%EXIT_CODE%" "%EXIT_CODE%"
    ) else (
      >> "%LOG%" echo [%DATE% %TIME%] info=ps1_diagnostics_present
    )
  )
)

exit /b %EXIT_CODE%

:RESOLVE_POWERSHELL
if defined PROCESSOR_ARCHITEW6432 if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE if exist "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" set "PS_EXE=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not defined PS_EXE for /f "usebackq delims=" %%P in (`where powershell.exe 2^>nul`) do if not defined PS_EXE set "PS_EXE=%%P"
if defined PS_EXE (
  >> "%LOG%" echo [%DATE% %TIME%] checkpoint=powershell path resolved exe="%PS_EXE%"
) else (
  >> "%LOG%" echo [%DATE% %TIME%] checkpoint=powershell path resolved exe=
)
goto :EOF

:WRITE_FALLBACK
set "FAIL_STAGE=%~1"
set "FAIL_REASON_CN=%~2"
set "FAIL_REASON=%~3"
set "FAIL_EXIT=%~4"
if not defined FAIL_EXIT set "FAIL_EXIT=1"
if not exist "%DIAG%" mkdir "%DIAG%" >nul 2>nul
echo Startup failed: %FAIL_REASON_CN%
if not exist "%STATUS_JSON%" call :WRITE_STATUS_JSON
if not exist "%FAILURE_TXT%" (
  > "%FAILURE_TXT%" echo L0 startup failed
  >> "%FAILURE_TXT%" echo stage=%FAIL_STAGE%
  >> "%FAILURE_TXT%" echo reason_cn=%FAIL_REASON_CN%
  >> "%FAILURE_TXT%" echo reason=%FAIL_REASON%
  >> "%FAILURE_TXT%" echo exitCode=%FAIL_EXIT%
  >> "%FAILURE_TXT%" echo psExe=%PS_EXE%
  >> "%FAILURE_TXT%" echo psScript=%PS_SCRIPT%
)
if not exist "%FAILURE_CSV%" (
  > "%FAILURE_CSV%" echo "field","value"
  >> "%FAILURE_CSV%" echo "status","failed"
  >> "%FAILURE_CSV%" echo "stage","%FAIL_STAGE%"
  >> "%FAILURE_CSV%" echo "reason_cn","%FAIL_REASON_CN%"
  >> "%FAILURE_CSV%" echo "reason","%FAIL_REASON%"
  >> "%FAILURE_CSV%" echo "exitCode","%FAIL_EXIT%"
  >> "%FAILURE_CSV%" echo "psExe","%PS_EXE%"
  >> "%FAILURE_CSV%" echo "psScript","%PS_SCRIPT%"
)
>> "%LOG%" echo [%DATE% %TIME%] fallback_written stage="%FAIL_STAGE%" exitCode=%FAIL_EXIT% reason="%FAIL_REASON%"
goto :EOF

:WRITE_STATUS_JSON
setlocal EnableDelayedExpansion
set "JSON_STAGE=%FAIL_STAGE%"
set "JSON_MESSAGE=%FAIL_REASON_CN%"
set "JSON_MESSAGE_EN=%FAIL_REASON%"
set "JSON_STAGE=!JSON_STAGE:\=\\!"
set "JSON_MESSAGE=!JSON_MESSAGE:\=\\!"
set "JSON_MESSAGE_EN=!JSON_MESSAGE_EN:\=\\!"
set "JSON_STAGE=!JSON_STAGE:"=\"!"
set "JSON_MESSAGE=!JSON_MESSAGE:"=\"!"
set "JSON_MESSAGE_EN=!JSON_MESSAGE_EN:"=\"!"
> "%STATUS_JSON%" echo {"status":"failed","stage":"!JSON_STAGE!","message":"!JSON_MESSAGE!","message_en":"!JSON_MESSAGE_EN!","exitCode":%FAIL_EXIT%}
endlocal
goto :EOF
