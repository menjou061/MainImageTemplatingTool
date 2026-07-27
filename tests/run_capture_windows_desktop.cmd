@echo off
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0capture_windows_desktop.ps1" -OutputPath "%~dp0ui-artifacts\current-desktop.png"
