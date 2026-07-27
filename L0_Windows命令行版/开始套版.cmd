@echo off
setlocal EnableExtensions DisableDelayedExpansion

call "%~dp0L0_Start.cmd"
exit /b %ERRORLEVEL%
