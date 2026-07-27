@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul 2>nul
title 电商主图套版工具 1.0

set "BASE=%~dp0"
if defined LOCALAPPDATA (
  set "DIAG=%LOCALAPPDATA%\电商主图套版工具\任务记录"
) else (
  set "DIAG=%BASE%_diagnostics"
)
set "RUNNER=%BASE%_internal\L0_Run.ps1"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%DIAG%" mkdir "%DIAG%" >nul 2>nul

if not exist "%RUNNER%" (
  > "%DIAG%\failure.txt" echo 电商主图套版工具启动失败
  >> "%DIAG%\failure.txt" echo 原因：安装包不完整，缺少 _internal\L0_Run.ps1
  > "%DIAG%\status.json" echo {"status":"failed","stage":"entry","message":"安装包不完整","exitCode":1}
  exit /b 1
)

if not exist "%PS_EXE%" (
  > "%DIAG%\failure.txt" echo 电商主图套版工具启动失败
  >> "%DIAG%\failure.txt" echo 原因：未找到 Windows PowerShell
  exit /b 1
)

start "" /min "%PS_EXE%" -STA -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%"
exit /b 0
