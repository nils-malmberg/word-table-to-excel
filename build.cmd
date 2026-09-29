@echo off
rem Maintenance du projet uniquement (les utilisateurs n ont rien a compiler).
rem Execute les tests, compile WordTableToExcel.dll (dossier Installation) et TableauxWordExcel.exe (dossier Application).
rem Les chargeurs natifs (Shim32/Shim64) se recompilent avec build\build-shim.sh (MinGW-w64).
rem Necessite le SDK .NET 8 : https://dotnet.microsoft.com/download
setlocal
cd /d "%~dp0"
where dotnet >nul 2>&1
if errorlevel 1 (
  echo Le SDK .NET est introuvable. Installez-le depuis https://dotnet.microsoft.com/download
  exit /b 1
)
dotnet test tests\WordTableToExcel.Tests -c Release
if errorlevel 1 exit /b 1
dotnet build src\WordTableToExcel -c Release
if errorlevel 1 exit /b 1
copy /y src\WordTableToExcel\bin\Release\net40\WordTableToExcel.dll Installation\ >nul
dotnet build src\TableauxWordExcel -c Release
if errorlevel 1 exit /b 1
copy /y src\TableauxWordExcel\bin\Release\net40\TableauxWordExcel.exe Application\ >nul
echo.
echo Dossiers Installation et Application mis a jour.
