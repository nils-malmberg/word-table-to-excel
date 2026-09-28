@echo off
rem Installe le complement Word "Tableaux Word vers Excel" pour tous les utilisateurs du poste.
rem Windows demande une confirmation administrateur (UAC).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if errorlevel 1 (
  echo.
  echo L installation a echoue ou a ete annulee.
)
pause
