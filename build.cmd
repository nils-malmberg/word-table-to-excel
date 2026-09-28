@echo off
rem Compile le complement, execute les tests et prepare le dossier dist\WordTableToExcel
rem (DLL + scripts d installation). Necessite le SDK .NET 8 ou plus : https://dotnet.microsoft.com/download
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

if exist dist rmdir /s /q dist
mkdir dist\WordTableToExcel
copy /y src\WordTableToExcel\bin\Release\net40\WordTableToExcel.dll dist\WordTableToExcel\ >nul
copy /y scripts\*.* dist\WordTableToExcel\ >nul
copy /y README.md dist\WordTableToExcel\LISEZMOI.md >nul

echo.
echo Paquet pret : %CD%\dist\WordTableToExcel
echo Pour installer : double-cliquez sur dist\WordTableToExcel\install.cmd (Word ferme).
