@echo off
rem Desinstalle le complement Word "Tableaux Word vers Excel" (confirmation administrateur demandee).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
pause
