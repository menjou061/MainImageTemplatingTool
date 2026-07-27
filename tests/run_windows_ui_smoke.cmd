@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0windows_ui_smoke.ps1" -ToolRoot "%~dp0tool" -ExcelPath "%~dp0test.xlsx" -PsdPath "%~dp0template.psd" -OutputRoot "%~dp0ui-output" -ArtifactDir "%~dp0ui-artifacts"
exit /b %ERRORLEVEL%
