@echo off
rem Installe le complement Word "Tableaux Word vers Excel" pour l utilisateur courant (sans droits administrateur).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if errorlevel 1 (
  echo.
  echo L installation a echoue.
) 
pause
