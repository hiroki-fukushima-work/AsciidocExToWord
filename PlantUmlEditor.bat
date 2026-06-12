@echo off

setlocal

ECHO PlantUMLエディタ

cd %~d0%~p0
echo %~d0%~p0

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PlantUmlEditor.ps1"
pause
endlocal
