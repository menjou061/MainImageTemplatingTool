@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul 2>nul
title L0 entry

set "BASE=%~dp0"
if defined LOCALAPPDATA (
  set "DIAG=%LOCALAPPDATA%\电商主图套版工具\任务记录"
) else (
  set "DIAG=%BASE%_diagnostics"
)
set "LOG=%DIAG%\start.log"
set "RUNNER=%BASE%_internal\L0_Run.bat"
set "RUNNER_CONSOLE=%DIAG%\runner_stdout_stderr.log"

if not exist "%DIAG%" mkdir "%DIAG%" >nul 2>nul
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=entered file=L0_Start.cmd
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=path resolved base="%BASE%" runner="%RUNNER%" diag="%DIAG%"

if not exist "%RUNNER%" (
  >> "%LOG%" echo [%DATE% %TIME%] error=missing_runner path="%RUNNER%"
  > "%DIAG%\failure.txt" echo L0 startup failed
  >> "%DIAG%\failure.txt" echo stage=entry
  >> "%DIAG%\failure.txt" echo reason_cn=runner batch file was not found
  >> "%DIAG%\failure.txt" echo reason=missing _internal\L0_Run.bat
  > "%DIAG%\failure.csv" echo "field","value"
  >> "%DIAG%\failure.csv" echo "status","failed"
  >> "%DIAG%\failure.csv" echo "stage","entry"
  >> "%DIAG%\failure.csv" echo "reason_cn","runner batch file was not found"
  >> "%DIAG%\failure.csv" echo "reason","missing _internal\L0_Run.bat"
  > "%DIAG%\status.json" echo {"status":"failed","stage":"entry","message":"runner batch file was not found","message_en":"missing _internal\\L0_Run.bat","exitCode":1}
  set "EXIT_CODE=1"
  goto DONE
)

set "L0_ENTRY_OK=1"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=before_runner archive="%RUNNER_CONSOLE%"
call "%RUNNER%" >> "%RUNNER_CONSOLE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=runner_returned exitCode=%EXIT_CODE%

:DONE
>> "%LOG%" echo [%DATE% %TIME%] checkpoint=L0_Start_end exitCode=%EXIT_CODE%
exit /b %EXIT_CODE%
