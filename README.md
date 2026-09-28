# Tableaux Word vers Excel

Complément COM pour Microsoft Word (Windows) qui exporte les tableaux du document ouvert vers un
classeur Excel, **une feuille par tableau**, en **conservant la mise en forme**, sans jamais modifier
le document Word.

- **Option A** : uniquement les tableaux qui ont une légende (champ `SEQ Tableau`, `SEQ Table`, `SEQ Tabla`…
  ou texte commençant par « Tabl… »). Chaque feuille porte le nom de sa légende.
- **Option B** : tous les tableaux du document. Les tableaux sans légende deviennent `Tableau_1`, `Tableau_2`…

---

## Sommaire

1. [Compatibilité](#compatibilité)
2. [Installation](#installation)
3. [Utilisation](#utilisation)
4. [Détection des légendes](#détection-des-légendes)
5. [Mise en forme conservée](#mise-en-forme-conservée)
6. [Intégrité du document Word](#intégrité-du-document-word)
7. [Désinstallation](#désinstallation)
8. [Dépannage](#dépannage)
9. [Compilation et tests](#compilation-et-tests)
10. [Architecture](#architecture)
11. [Réglages avancés](#réglages-avancés)

---

## Compatibilité

| Élément | Prise en charge |
|---|---|
| Word 2007, 2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365 | Bouton **Tableaux vers Excel** dans les onglets **Accueil** et **Références** du ruban |
| Word 2000, 2002 (XP), 2003 | Bouton **Tableaux → Excel** dans la barre d'outils *Standard* |
| Office 32 bits et 64 bits, installation MSI ou « Démarrer en un clic » | Oui (même DLL, enregistrée dans les deux vues du registre) |
| Windows | 7 SP1 et suivants, avec .NET Framework 4.x (inclus d'office depuis Windows 8) |
| Excel | **Non requis** pour créer le fichier : le classeur `.xlsx` est écrit directement |

Le complément n'utilise aucune bibliothèque d'interopérabilité liée à une version d'Office (PIA) :
tout l'accès à Word se fait en liaison tardive, ce qui le rend indépendant de la version installée.

## Installation

Aucun droit administrateur n'est nécessaire : l'installation se fait pour l'utilisateur courant.

1. Récupérez le dossier `WordTableToExcel` :
   - soit l'artefact **WordTableToExcel** produit par l'intégration continue GitHub (onglet *Actions*) ;
   - soit en compilant vous-même (voir [Compilation](#compilation-et-tests)) : le dossier est créé dans `dist\`.
2. **Fermez Word.**
3. Double-cliquez sur **`install.cmd`**.
4. Ouvrez Word : le bouton **Tableaux vers Excel** apparaît à droite de l'onglet **Accueil** (et de l'onglet **Références**).

Le script copie `WordTableToExcel.dll` dans `%LOCALAPPDATA%\WordTableToExcel`, l'enregistre comme
composant COM pour l'utilisateur (`HKCU`) et le déclare auprès de Word.

> Installation pour tous les utilisateurs d'un poste (administrateur) : exécutez
> `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase WordTableToExcel.dll`
> (Office 64 bits) ou `%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe /codebase WordTableToExcel.dll`
> (Office 32 bits) depuis une invite de commandes administrateur.

## Utilisation

1. Ouvrez le document Word qui contient les tableaux.
2. Cliquez sur **Accueil › Export Excel › Tableaux vers Excel**.
3. La boîte de dialogue indique le nombre de tableaux trouvés, combien ont une légende, et si les
   légendes sont placées au-dessus ou au-dessous des tableaux. Choisissez :
   - **A — Uniquement les tableaux qui ont une légende**, ou
   - **B — Tous les tableaux du document**.

   La liste montre les feuilles Excel qui seront créées, avec leur nom.
4. Options :
   - *Écrire la légende complète en haut de chaque feuille* : la légende figure en A1 (utile car un nom
     de feuille Excel est limité à 31 caractères) et le tableau commence en ligne 3 ;
   - *Convertir les nombres en valeurs numériques Excel* : « 1 234,50 », « 12,5 % », « 45 € » deviennent
     de vrais nombres (avec un format qui conserve leur apparence). Sinon, les cellules sont copiées
     telles quelles, sous forme de texte.
5. Cliquez sur **Exporter…**, puis choisissez le dossier et le nom du fichier dans la fenêtre
   d'enregistrement Windows (par défaut : dossier du document, nom `<document> - tableaux.xlsx`).
6. Une fenêtre de progression s'affiche (l'export peut être annulé), puis un compte rendu propose
   d'ouvrir le classeur.

Les choix sont mémorisés d'une fois sur l'autre.

## Détection des légendes

Pour chaque tableau, le complément examine le paragraphe **juste au-dessus** et **juste au-dessous**
(en ignorant jusqu'à trois paragraphes vides). Ce paragraphe est une légende de tableau :

1. s'il contient un **champ `SEQ`** dont l'identificateur désigne un tableau, dans n'importe quelle langue :
   `SEQ Table`, `SEQ Tableau`, `SEQ Tabla`, `SEQ Tabelle`, `SEQ Tabella`, `SEQ Tabel`, `SEQ Tabela`,
   `SEQ Tabell`, `SEQ Tabulka`, `SEQ Taulukko`, `SEQ Táblázat`, `SEQ Tablo`, `SEQ Таблица`, `SEQ Πίνακας`,
   `SEQ 表`, `SEQ 표`… (tout identificateur commençant par « Tabl », « Tabel », « Tabul »… est accepté) ;
2. ou si son **texte commence par « Tabl… »** ou un libellé équivalent (« Table 1. », « Tableau 3 : »,
   « Tabla 2 – », « Tabelle 4: », « Tab. 5 », « 表1 »…). Pour éviter les faux positifs (« Tablette… »),
   un mot qui n'est pas un libellé connu doit être suivi d'un numéro, sauf si le paragraphe utilise le
   style **Légende** de Word.

Un paragraphe contenant un champ `SEQ Figure` (ou toute autre séquence) n'est jamais pris pour une
légende de tableau. Le libellé « Tableau » de la langue d'interface de Word est aussi reconnu
automatiquement.

**Position automatique** : le complément détermine la convention du document (légendes au-dessus ou
au-dessous) d'après les tableaux dont la légende n'est possible que d'un seul côté, puis l'applique aux
cas ambigus. Une légende située entre deux tableaux n'est jamais attribuée aux deux.

**Nom des feuilles** : la légende nettoyée (« Tableau 3 : Coûts » devient « Tableau 3 - Coûts »),
tronquée à 31 caractères, sans les caractères interdits par Excel (`\ / ? * [ ] :`) et rendue unique
(« … (2) »). Sans légende : `Tableau_N`, où N est le rang du tableau dans le document.

## Mise en forme conservée

| Word | Excel |
|---|---|
| Police, taille, couleur du texte (y compris couleurs de thème) | Identiques, **par segment de texte** (texte enrichi dans une même cellule) |
| Gras, italique, souligné (simple/double), barré, exposant, indice | Identiques |
| Majuscules (attribut « Majuscules ») | Texte converti en majuscules |
| Texte masqué | Non exporté |
| Trame (fond) de cellule, y compris motifs en pourcentage | Couleur de remplissage |
| Mise en forme du **style de tableau** : ligne d'en-tête, dernière ligne, première/dernière colonne, lignes et colonnes à bandes | Fonds et bordures reproduits cellule par cellule |
| Surlignage | Couleur de fond de la cellule (Excel ne peut pas surligner une partie de cellule) |
| Cellules fusionnées (horizontalement et verticalement) | Cellules fusionnées |
| Bordures (style, épaisseur, couleur ; directes, du tableau ou du style) | Bordures équivalentes (fin, moyen, épais, double, pointillé, tirets…) |
| Alignement horizontal et vertical, orientation du texte | Identiques |
| Largeur des colonnes, hauteur de ligne « exacte » ou « au moins » | Reproduites |
| Sauts de ligne et paragraphes dans une cellule | Retours à la ligne dans la cellule (renvoi à la ligne activé) |

Limites connues :

- les **images** et objets insérés dans les cellules ne sont pas copiés (seul le texte l'est) ;
- un **tableau imbriqué** dans une cellule est exporté sous forme de texte dans la cellule parente ;
- seuls les tableaux du **corps du document** sont exportés (pas ceux des en-têtes, pieds de page
  ou zones de texte) ;
- le surlignage partiel d'une cellule colore toute la cellule ;
- les motifs de trame (rayures, quadrillages) sont rendus par une couleur unie.

## Intégrité du document Word

Le document ouvert est **uniquement lu** :

- aucune méthode de modification n'est appelée (pas de copier-coller, pas de sélection, pas d'enregistrement) ;
- l'indicateur « document modifié » est préservé : Word ne propose pas d'enregistrer un document que
  l'export n'a pas changé ;
- pendant l'export, la fenêtre de progression est modale : on ne peut pas modifier le document pendant sa lecture ;
- le presse-papiers n'est pas utilisé ;
- le classeur est d'abord écrit dans un fichier temporaire puis renommé : en cas d'incident, aucun
  fichier à moitié écrit n'est laissé et un classeur existant n'est remplacé qu'une fois le nouveau complet ;
- il est impossible de choisir le document Word lui-même comme fichier de destination ;
- toute erreur est interceptée et expliquée : elle ne peut ni faire planter Word, ni conduire Word à
  désactiver le complément.

## Désinstallation

Fermez Word, puis double-cliquez sur **`uninstall.cmd`** (dans le dossier d'installation ou le dossier
`scripts` du projet). Les clés de registre, les fichiers et les préférences sont supprimés
(`uninstall.ps1 -KeepSettings` conserve les préférences).

## Dépannage

| Problème | Solution |
|---|---|
| Le bouton n'apparaît pas | **Fichier › Options › Compléments**, liste *Gérer : Compléments COM* › **Atteindre…** : cochez « Tableaux Word vers Excel ». |
| Le complément a été désactivé par Word | **Fichier › Options › Compléments**, *Gérer : Éléments désactivés* › **Atteindre…** : réactivez-le, puis redémarrez Word. |
| « Comportement au chargement : non chargé. Une erreur d'exécution s'est produite » | Vérifiez que .NET Framework 4.x est installé ; relancez `install.cmd` (il débloque la DLL téléchargée). |
| Le classeur ne peut pas être enregistré | Le fichier est probablement ouvert dans Excel : fermez-le et recommencez. |
| Un tableau n'est pas exporté ou apparaît simplifié | Le compte rendu final et le journal le signalent. |

Journal de diagnostic : `%LOCALAPPDATA%\WordTableToExcel\WordTableToExcel.log`.

## Compilation et tests

Prérequis : [SDK .NET 8](https://dotnet.microsoft.com/download) (la DLL cible .NET Framework 4.0 grâce aux
assemblys de référence NuGet ; Visual Studio 2019+ ou Rider peuvent aussi ouvrir `WordTableToExcel.sln`).

```bat
build.cmd
```

exécute les tests, compile `WordTableToExcel.dll` (AnyCPU) et prépare `dist\WordTableToExcel\`
(DLL + scripts d'installation). Commandes équivalentes :

```bash
dotnet test tests/WordTableToExcel.Tests      # 150+ tests, exécutables sous Windows, Linux ou macOS
dotnet build src/WordTableToExcel -c Release
```

Les tests couvrent l'analyse de la structure des tableaux (fichiers réels générés par
`tests/fixtures/make_fixtures.py`), les styles de tableau, la détection des légendes, les noms de
feuilles, la conversion des nombres, l'écriture du `.xlsx` et la lecture Word via un faux modèle
objet Word (`tests/WordTableToExcel.Tests/Fakes`). Ils s'exécutent sans Word.

L'intégration continue (`.github/workflows/build.yml`, Windows) exécute ces tests, compile la DLL, puis
vérifie l'installation réelle : script `install.ps1`, activation COM du complément depuis un processus
32 bits et 64 bits, chargement du ruban, désinstallation. Le fonctionnement dans Word lui-même
(clic sur le bouton, lecture d'un vrai document) se valide sur un poste équipé de Word.

## Architecture

```
src/WordTableToExcel/
├── AddIn/            Point d'entrée COM (IDTExtensibility2 + ruban), barre d'outils Word 2000-2003,
│                     déclarations des interfaces Office (sans PIA), Ribbon.xml
├── Export/           Déroulement de l'export (ExportService) et sélection A/B + noms de feuilles (ExportPlan)
├── UI/               Boîte de dialogue de choix, fenêtre de progression, messages
├── Word/             Lecture du document via le modèle objet (liaison tardive) :
│                     tableaux, texte et mise en forme des caractères, légendes, protection du document
├── Core/             Code indépendant de Word et de Windows :
│   ├── Layout/       Analyse du XML des tableaux (grille, fusions, styles de tableau, bordures)
│   ├── Captions/     Reconnaissance multilingue des légendes, attribution, noms de feuilles
│   ├── Text/         Nettoyage du texte, reconnaissance des nombres
│   ├── Xlsx/         Écriture du classeur .xlsx (ZIP + SpreadsheetML, styles dédupliqués)
│   └── Model/        Modèle intermédiaire (tableau, cellule, segment de texte, bordures)
└── Infrastructure/   Journal, préférences
```

Déroulement d'un export :

1. **Détection** — pour chaque tableau, paragraphes voisins → légende candidate (`WordCaptionScanner`),
   puis attribution globale avec détection de la convention au-dessus/au-dessous (`CaptionAssigner`).
2. **Choix** — boîte de dialogue A/B, aperçu des feuilles, fenêtre d'enregistrement.
3. **Structure** — le XML du tableau (`Range.WordOpenXML`, Word 2010+, ou `Range.XML`, Word 2003+) donne la
   grille exacte : `gridSpan`, `vMerge`, `gridBefore`, largeurs, hauteurs, trames et bordures, y compris
   celles du style de tableau et de ses zones conditionnelles. Sans XML (Word 2000/2002), la grille est
   reconstruite à partir de la largeur des cellules.
4. **Contenu** — le texte et la mise en forme effective de chaque caractère sont lus via le modèle objet :
   la police d'une plage est interrogée d'un bloc ; si elle est hétérogène, la plage est coupée en deux
   (recherche dichotomique). Le coût dépend du nombre de changements de mise en forme, pas de la longueur du texte.
5. **Écriture** — `XlsxWorkbookWriter` produit un `.xlsx` standard (Office Open XML) ; les classeurs
   générés sont relus dans les tests et validés avec LibreOffice et openpyxl.

## Réglages avancés

Préférences stockées dans `HKEY_CURRENT_USER\Software\WordTableToExcel` :

| Valeur | Type | Rôle |
|---|---|---|
| `AllTables` | DWORD | Dernier choix A (0) / B (1) |
| `IncludeCaptionRow` | DWORD | Légende en A1 (1) ou non (0) |
| `ConvertNumbers` | DWORD | Conversion des nombres (1) ou non (0) |
| `LastFolder` | Chaîne | Dernier dossier d'enregistrement |
| `ExtraCaptionLabels` | Chaîne | Libellés de légende supplémentaires, séparés par `;` (ex. `Annexe;Tab.`) |
