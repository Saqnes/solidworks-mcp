@echo off
rem solidworks-mcp install: double-click this file (it runs install.ps1).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
echo.
pause
