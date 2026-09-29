@echo off
rem ===========================================================================
rem  Installation du complement Word "Tableaux Word vers Excel"
rem  - pour l'utilisateur courant, SANS droits administrateur ;
rem  - rien a compiler : tous les fichiers necessaires sont dans ce dossier.
rem  Double-cliquez sur ce fichier (Word ferme de preference).
rem ===========================================================================
setlocal EnableExtensions

set "SOURCE=%~dp0"
set "BASE=%LOCALAPPDATA%"
if not defined BASE set "BASE=%APPDATA%"
set "TARGET=%BASE%\WordTableToExcel"
set "CLSID={03F63233-F2FE-4A75-AF7A-F99CBEDC8030}"
set "PROGID=WordTableToExcel.Connect"
set "NAME=Tableaux Word vers Excel"
set "FILES=WordTableToExcel.dll WordTableToExcel.Shim32.dll WordTableToExcel.Shim64.dll"

echo.
echo  Installation de "%NAME%"
echo  pour l'utilisateur %USERNAME% (aucun droit administrateur necessaire)
echo  -----------------------------------------------------------------------
echo.

rem --- 1. Verifications ------------------------------------------------------
for %%F in (%FILES%) do (
  if not exist "%SOURCE%%%F" (
    echo  ERREUR : fichier manquant : %%F
    echo.
    echo  Extrayez d'abord TOUT le fichier .zip ^(clic droit ^> Extraire tout^),
    echo  puis lancez install.cmd depuis le dossier "Installation" extrait.
    goto :failed
  )
)

reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" >nul 2>&1
if errorlevel 1 (
  echo  ATTENTION : .NET Framework 4 n'a pas ete detecte sur ce poste.
  echo  Il est inclus dans Windows 8, 10 et 11. Sous Windows 7, installez
  echo  .NET Framework 4.8 ^(Microsoft^), sinon le complement ne pourra pas demarrer.
  echo.
)

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo  Word est ouvert. Fermez Word, puis appuyez sur une touche pour continuer...
  if not defined WTTE_NO_PAUSE pause >nul
  echo.
)

rem --- 2. Copie des fichiers dans %LOCALAPPDATA%\WordTableToExcel ----------
if not exist "%TARGET%" mkdir "%TARGET%"
if not exist "%TARGET%" (
  echo  ERREUR : impossible de creer le dossier "%TARGET%".
  goto :failed
)
for %%F in (%FILES%) do (
  copy /Y "%SOURCE%%%F" "%TARGET%\%%F" >nul
  if errorlevel 1 (
    echo  ERREUR : impossible de copier %%F.
    echo  Word est probablement encore ouvert : fermez-le et relancez install.cmd.
    goto :failed
  )
  rem Fichiers telecharges depuis Internet : on retire la marque de provenance,
  rem sinon .NET refuse de charger WordTableToExcel.dll.
  (echo [ZoneTransfer]& echo ZoneId=0) > "%TARGET%\%%F:Zone.Identifier" 2>nul
)
echo  Fichiers copies dans "%TARGET%".

rem --- 3. Enregistrement COM pour l'utilisateur courant (HKCU) ---------------
rem Chargeur 32 bits pour Office 32 bits, chargeur 64 bits pour Office 64 bits.
set "IS64="
if /I "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "IS64=1"
if /I "%PROCESSOR_ARCHITEW6432%"=="AMD64" set "IS64=1"
if /I "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "IS64=1"
if /I "%PROCESSOR_ARCHITEW6432%"=="ARM64" set "IS64=1"

if defined IS64 (
  call :register /reg:32 "%TARGET%\WordTableToExcel.Shim32.dll" || goto :failed
  call :register /reg:64 "%TARGET%\WordTableToExcel.Shim64.dll" || goto :failed
) else (
  call :register "" "%TARGET%\WordTableToExcel.Shim32.dll" || goto :failed
)

rem --- 4. Declaration du complement aupres de Word (toutes versions) ---------
set "ADDIN=HKCU\Software\Microsoft\Office\Word\Addins\%PROGID%"
reg add "%ADDIN%" /v FriendlyName /t REG_SZ /d "%NAME%" /f >nul || goto :failed
reg add "%ADDIN%" /v Description /t REG_SZ /d "Exporte les tableaux du document vers Excel et importe des tableaux Excel dans le document, avec leur mise en forme." /f >nul || goto :failed
reg add "%ADDIN%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul || goto :failed
rem Office 2013 et suivants : ne pas desactiver le complement pour lenteur au demarrage.
for %%V in (15.0 16.0) do reg add "HKCU\Software\Microsoft\Office\%%V\Word\Resiliency\DoNotDisableAddinList" /v %PROGID% /t REG_DWORD /d 1 /f >nul 2>&1

echo  Complement enregistre pour Word.

rem --- 5. Regles de securite de Word qui empecheraient le chargement ---------
rem "Exiger que les complements soient signes par un editeur approuve" (requireaddinsig)
rem ou "Desactiver tous les complements d'application" (disablealladdins) : Word refuse
rem alors le complement et indique "L'utilisateur a choisi de desactiver les macros".
set "POLICY_BLOCK="
set "USER_BLOCK="
for %%V in (12.0 14.0 15.0 16.0) do (
  for %%N in (requireaddinsig disablealladdins) do (
    call :isset "HKCU\Software\Policies\Microsoft\Office\%%V\Word\Security" %%N && set "POLICY_BLOCK=1"
    call :isset "HKLM\Software\Policies\Microsoft\Office\%%V\Word\Security" %%N && set "POLICY_BLOCK=1"
    call :isset "HKCU\Software\Microsoft\Office\%%V\Word\Security" %%N && set "USER_BLOCK=1"
  )
)
if defined POLICY_BLOCK goto :policyblock
if defined USER_BLOCK goto :userblock

echo.
echo  =======================================================================
echo   Installation terminee.
echo   Ouvrez Word : les boutons "Tableaux vers Excel" et "Importer depuis
echo   Excel" se trouvent a droite des onglets Accueil et References, groupe
echo   "Excel" (Word 2000-2003 : barre d'outils Standard).
echo  =======================================================================
echo.
if not defined WTTE_NO_PAUSE pause
exit /b 0

:policyblock
echo.
echo  =======================================================================
echo   Installation terminee, MAIS votre organisation impose a Word une regle
echo   de securite : seuls les complements SIGNES par un editeur approuve par
echo   votre service informatique peuvent se charger. Word n'affichera pas les
echo   boutons et indiquera "Non charge. L'utilisateur a choisi de desactiver
echo   les macros".
echo.
echo   Ce reglage ne peut pas etre modifie depuis ce poste. Transmettez a votre
echo   service informatique le fichier INFORMATIQUE.txt de ce dossier : il
echo   explique comment signer le complement ^(quelques minutes^).
echo  =======================================================================
echo.
if not defined WTTE_NO_PAUSE pause
exit /b 0

:userblock
echo.
echo  =======================================================================
echo   Installation terminee, MAIS Word est regle pour refuser les complements
echo   non signes ^(il indiquerait "L'utilisateur a choisi de desactiver les
echo   macros"^). Pour l'autoriser, dans Word :
echo     Fichier ^> Options ^> Centre de gestion de la confidentialite ^>
echo     Parametres du Centre de gestion de la confidentialite ^> Complements
echo   decochez "Exiger que les complements d'application soient signes par un
echo   editeur approuve" et "Desactiver tous les complements d'application",
echo   puis redemarrez Word.
echo   Si ces cases sont grisees, voyez INFORMATIQUE.txt avec votre service
echo   informatique.
echo  =======================================================================
echo.
if not defined WTTE_NO_PAUSE pause
exit /b 0

:failed
echo.
echo  L'installation n'a pas abouti. Aucun droit administrateur n'est requis :
echo  verifiez le message ci-dessus, puis relancez install.cmd.
echo.
if not defined WTTE_NO_PAUSE pause
exit /b 1

rem ---------------------------------------------------------------------------
rem  :isset <cle> <valeur> : code 0 si la valeur DWORD existe et vaut 1
rem ---------------------------------------------------------------------------
:isset
reg query "%~1" /v %2 2>nul | findstr /R /I /C:"REG_DWORD *0x1$" >nul
exit /b %errorlevel%

rem ---------------------------------------------------------------------------
rem  :register <option de vue du registre> <chemin du chargeur>
rem ---------------------------------------------------------------------------
:register
set "VIEW=%~1"
set "SERVER=%~2"
set "KEY=HKCU\Software\Classes\CLSID\%CLSID%"
reg delete "%KEY%" /f %VIEW% >nul 2>&1
reg add "%KEY%" /ve /t REG_SZ /d "%NAME%" /f %VIEW% >nul || exit /b 1
reg add "%KEY%\InprocServer32" /ve /t REG_SZ /d "%SERVER%" /f %VIEW% >nul || exit /b 1
reg add "%KEY%\InprocServer32" /v ThreadingModel /t REG_SZ /d Apartment /f %VIEW% >nul || exit /b 1
reg add "%KEY%\ProgID" /ve /t REG_SZ /d "%PROGID%" /f %VIEW% >nul || exit /b 1
reg add "HKCU\Software\Classes\%PROGID%" /ve /t REG_SZ /d "%NAME%" /f %VIEW% >nul || exit /b 1
reg add "HKCU\Software\Classes\%PROGID%\CLSID" /ve /t REG_SZ /d "%CLSID%" /f %VIEW% >nul || exit /b 1
exit /b 0
