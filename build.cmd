@echo off
rem ===========================================================================
rem  Compilation SANS SDK et SANS INSTALLATION : compilateur C# fourni avec
rem  Windows (.NET Framework 4 : C:\Windows\Microsoft.NET\Framework*\v4.0.30319).
rem
rem  Produit, apres verification par l'autotest de l'application (sans Word) :
rem    Installation\WordTableToExcel.dll   (complement Word, signe par nom fort)
rem    Application\TableauxWordExcel.exe   (application autonome)
rem  En cas d'erreur, rien n'est remplace. Version : src\Version.cs.
rem  Non recompiles ici : chargeurs natifs du complement (Shim32/Shim64, en C).
rem  Avec le SDK .NET (tests automatiques) : voir DEVELOPPEMENT.md.
rem ===========================================================================
setlocal
cd /d "%~dp0"

rem Console en UTF-8 le temps du script (messages du compilateur et de l'autotest).
for /f "tokens=2 delims=:." %%c in ('chcp') do set "OLDCP=%%c"
chcp 65001 >nul

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo Compilateur introuvable : le .NET Framework 4 n'est pas installe sur ce poste.
  goto echec
)
echo Compilateur : %CSC%

set "OUT=build\sortie"
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%" || goto echec

rem Options communes : C# 5 (syntaxe de ce compilateur), sources en UTF-8, bibliotheques du .NET Framework 4.
set COMMON=/nologo /noconfig /langversion:5 /codepage:65001 /utf8output /optimize+ /debug- /platform:anycpu /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll
rem Code partage par le complement et l'application.
set SHARED=/recurse:src\WordTableToExcel\Core\*.cs /recurse:src\WordTableToExcel\Word\*.cs /recurse:src\WordTableToExcel\Export\*.cs /recurse:src\WordTableToExcel\Import\*.cs /recurse:src\WordTableToExcel\UI\*.cs /recurse:src\WordTableToExcel\Infrastructure\*.cs

echo.
echo [1/3] Complement Word : WordTableToExcel.dll
"%CSC%" %COMMON% /target:library /out:"%OUT%\WordTableToExcel.dll" /keyfile:src\WordTableToExcel\WordTableToExcel.snk /resource:src\WordTableToExcel\AddIn\Ribbon.xml,WordTableToExcel.Ribbon.xml src\Version.cs /recurse:src\WordTableToExcel\Properties\*.cs /recurse:src\WordTableToExcel\AddIn\*.cs %SHARED%
if errorlevel 1 goto echec

echo.
echo [2/3] Application : TableauxWordExcel.exe
"%CSC%" %COMMON% /target:winexe /out:"%OUT%\TableauxWordExcel.exe" /win32icon:src\TableauxWordExcel\TableauxWordExcel.ico /win32manifest:src\TableauxWordExcel\app.manifest /resource:src\TableauxWordExcel\TableauxWordExcel.ico,TableauxWordExcel.ico src\Version.cs src\TableauxWordExcel\Program.cs /recurse:src\TableauxWordExcel\Properties\*.cs /recurse:src\TableauxWordExcel\App\*.cs %SHARED%
if errorlevel 1 goto echec

echo.
echo [3/3] Autotest de l'application (sans Word, quelques secondes)
rem Classeur de test du depot (lecture et boite d'import), s'il est present.
set "FIXTURE=%CD%\tests\fixtures\import_complexe.xlsx"
if exist "%FIXTURE%" (
  start "" /wait "%OUT%\TableauxWordExcel.exe" --selftest "%CD%\%OUT%\autotest" "%FIXTURE%"
) else (
  start "" /wait "%OUT%\TableauxWordExcel.exe" --selftest "%CD%\%OUT%\autotest"
)
set "RESULT=%ERRORLEVEL%"
if exist "%OUT%\autotest\selftest.txt" type "%OUT%\autotest\selftest.txt"
if not "%RESULT%"=="0" (
  echo Autotest en echec, code %RESULT%.
  goto echec
)

copy /y "%OUT%\WordTableToExcel.dll" Installation\ >nul || goto echec
copy /y "%OUT%\TableauxWordExcel.exe" Application\ >nul || goto echec
echo.
echo Termine : Installation\WordTableToExcel.dll et Application\TableauxWordExcel.exe sont a jour.
if defined OLDCP chcp %OLDCP% >nul
exit /b 0

:echec
echo.
echo ECHEC : rien n'a ete remplace dans Installation\ ni dans Application\.
if defined OLDCP chcp %OLDCP% >nul
exit /b 1
