@echo off
rem ===========================================================================
rem  Desinstallation du complement Word "Tableaux Word vers Excel"
rem  (utilisateur courant, sans droits administrateur).
rem ===========================================================================
setlocal EnableExtensions

set "BASE=%LOCALAPPDATA%"
if not defined BASE set "BASE=%APPDATA%"
set "TARGET=%BASE%\WordTableToExcel"
set "CLSID={03F63233-F2FE-4A75-AF7A-F99CBEDC8030}"
set "PROGID=WordTableToExcel.Connect"

echo.
echo  Desinstallation de "Tableaux Word vers Excel"
echo  -----------------------------------------------------------------------
echo.

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo  Word est ouvert. Fermez Word, puis appuyez sur une touche pour continuer...
  if not defined WTTE_NO_PAUSE pause >nul
  echo.
)

for %%V in ("" "/reg:32" "/reg:64") do (
  reg delete "HKCU\Software\Classes\CLSID\%CLSID%" /f %%~V >nul 2>&1
  reg delete "HKCU\Software\Classes\%PROGID%" /f %%~V >nul 2>&1
)
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\%PROGID%" /f >nul 2>&1
for %%V in (15.0 16.0) do reg delete "HKCU\Software\Microsoft\Office\%%V\Word\Resiliency\DoNotDisableAddinList" /v %PROGID% /f >nul 2>&1
reg delete "HKCU\Software\WordTableToExcel" /f >nul 2>&1
echo  Enregistrements supprimes.

set "RESULT=0"
if exist "%TARGET%" rmdir /s /q "%TARGET%" >nul 2>&1
if exist "%TARGET%" (
  echo  ATTENTION : le dossier "%TARGET%" n'a pas pu etre supprime
  echo  ^(Word est-il encore ouvert ?^). Supprimez-le apres avoir ferme Word.
  set "RESULT=1"
) else (
  echo  Fichiers supprimes.
)

echo.
echo  Desinstallation terminee.
echo.
if not defined WTTE_NO_PAUSE pause
exit /b %RESULT%
