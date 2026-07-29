@echo off
:: Run as Administrator: right-click > Run as administrator
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0switch-to-google-drive.ps1"
pause
