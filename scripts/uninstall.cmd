@echo off
rem Desinstalle le complement Word "Tableaux Word vers Excel".
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
pause
