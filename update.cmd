@echo off
rem solidworks-mcp update: pulls the latest code, then reinstalls and rechecks.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -Update %*
echo.
pause
