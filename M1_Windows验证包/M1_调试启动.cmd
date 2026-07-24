@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
title M1 Windows Photoshop COM 诊断启动

set "SCRIPT_DIR=%~dp0"
set "MAIN_BAT=%SCRIPT_DIR%M1_一键验证.bat"
set "OUTPUT_DIR=%SCRIPT_DIR%输出"
set "LOG_FILE=%OUTPUT_DIR%\M1启动日志.txt"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%" >nul 2>nul

call :LOG "========================================"
call :LOG "M1 Windows Photoshop COM 诊断启动"
call :LOG "启动时间：%DATE% %TIME%"
call :LOG "脚本目录：%SCRIPT_DIR%"
call :LOG "主bat路径：%MAIN_BAT%"

echo ========================================
echo M1 Windows Photoshop COM 诊断启动
echo ========================================
echo.
echo 这是用于排查双击窗口闪退的诊断启动。
echo 请不要复制或粘贴命令，只查看本窗口和自动打开的输出文件夹。
echo 如果一键验证闪退，请把整个“输出”文件夹发回开发同事。
echo.

if exist "%MAIN_BAT%" (
  call :LOG "是否发现主bat：是"
) else (
  call :LOG "是否发现主bat：否"
  echo 未找到主验证脚本：
  echo %MAIN_BAT%
  echo.
  echo 请确认本文件和 M1_一键验证.bat 放在同一个文件夹。
  echo.
  goto :KEEP_OPEN
)

call :LOG "执行主bat前：%DATE% %TIME%"
echo 正在调用 M1_一键验证.bat，请稍候...
echo.

set "M1_DEBUG_WRAPPER=1"
cmd /d /c call "%MAIN_BAT%"
set "M1_EXIT_CODE=%ERRORLEVEL%"
set "M1_DEBUG_WRAPPER="

echo.
call :LOG "执行主bat后：%DATE% %TIME%"
call :LOG "返回码：%M1_EXIT_CODE%"
echo ========================================
echo 诊断启动已结束
echo 返回码：%M1_EXIT_CODE%
echo 启动日志：%LOG_FILE%
echo ========================================
echo.
echo 请保持这个窗口可见；不要复制命令。
echo 请把整个“输出”文件夹发回开发同事。
echo.

:KEEP_OPEN
call :LOG "诊断窗口保持打开：%DATE% %TIME%"
echo 按 Ctrl+C 或关闭窗口可结束诊断窗口。
cmd /k
goto :EOF

:LOG
>> "%LOG_FILE%" echo [%DATE% %TIME%] %~1
goto :EOF
