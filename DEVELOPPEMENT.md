# Documentation développeur — Tableaux Word ↔ Excel

Ce document décrit l'implémentation complète du projet, pour le maintenir et le faire évoluer : principes,
arborescence, architecture, déroulement détaillé de l'export et de l'import, application autonome,
complément Word (gelé), réglages, compilation, tests, intégration continue, publication, conventions et
pièges connus de Word. Le mode d'emploi utilisateur est dans le [README](README.md).

Version en cours : **2.1.0** (non publiée : pas encore de tag).

---

## Sommaire

1. [Vue d'ensemble et principes](#1-vue-densemble-et-principes)
2. [Arborescence](#2-arborescence)
3. [Architecture en couches](#3-architecture-en-couches)
4. [Export Word → Excel](#4-export-word--excel)
5. [Import Excel → Word](#5-import-excel--word)
6. [Application autonome](#6-application-autonome)
7. [Complément Word (gelé)](#7-complément-word-gelé)
8. [Réglages, journal et fichiers écrits sur le poste](#8-réglages-journal-et-fichiers-écrits-sur-le-poste)
9. [Compilation](#9-compilation)
10. [Tests](#10-tests)
11. [Intégration continue et publication](#11-intégration-continue-et-publication)
12. [Conventions de code](#12-conventions-de-code)
13. [Pièges connus de Word et de COM](#13-pièges-connus-de-word-et-de-com)
14. [Évolutions possibles et idées écartées](#14-évolutions-possibles-et-idées-écartées)

---

## 1. Vue d'ensemble et principes

Deux produits partagent le même code :

| Produit | Dossier livré | Principe | Statut |
|---|---|---|---|
| **Application autonome** `TableauxWordExcel.exe` | `Application/` | Un seul exécutable qui pilote Word **depuis un autre processus** (automatisation COM). Rien n'est chargé dans Word : les stratégies de Word sur les compléments ne s'appliquent pas. | **Produit principal** |
| **Complément Word** `WordTableToExcel.dll` | `Installation/` | Complément COM chargé **dans** Word (boutons du ruban), via un chargeur natif enregistré pour l'utilisateur. | **Gelé** (voir § 7) |

Principes non négociables, vérifiés par les tests et la relecture :

- **Le document Word n'est jamais modifié par un export.** Lecture seule, pas de presse-papiers, pas de
  sélection ; l'indicateur « document modifié » est restauré (`WordDocumentGuard`).
- **Le classeur est écrit « tout ou rien ».** Fichier temporaire puis renommage (`ExportService.SaveSafely`) :
  jamais de fichier à moitié écrit, un classeur existant n'est remplacé qu'une fois le nouveau complet.
- **Aucune perte d'information silencieuse.** Tout ce qui n'arrive pas dans le classeur est signalé en tête du
  compte rendu (tableau non exporté, image, texte coupé…), et tout le détail est écrit dans le journal.
- **Jamais de valeur différente de ce qu'affiche Word.** La lecture rapide (XML) est comparée au texte de Word ;
  au moindre écart, le tableau est relu cellule par cellule.
- **Aucune dépendance à l'exécution** : .NET Framework 4.0, pas de PIA Office (liaison tardive `dynamic`),
  pas de bibliothèque tierce. ZIP, XLSX, formats de nombre Excel… sont implémentés dans le projet.
- **Aucun droit administrateur**, aucune installation pour l'application.
- **Aucune exception ne remonte jusqu'à Word** (elle ferait désactiver le complément) : tout est intercepté,
  expliqué à l'utilisateur et journalisé.
- **Interface, messages, commentaires et documentation en français.**

## 2. Arborescence

```
.
├── README.md                     Mode d'emploi utilisateur (court)
├── DEVELOPPEMENT.md              Ce document
├── WordTableToExcel.sln          Solution : complément, application, tests
├── build.cmd                     Windows : tests + recompilation de la DLL et de l'exe (copiés dans Installation/ et Application/)
├── Application/                  LIVRÉ — application prête à l'emploi
│   ├── TableauxWordExcel.exe     Exécutable unique (.NET Framework 4, AnyCPU) — à recompiler et recopier après chaque modification
│   └── LISEZMOI.txt              Mode d'emploi court
├── Installation/                 LIVRÉ — complément Word (gelé)
│   ├── WordTableToExcel.dll      Complément .NET (signé par nom fort)
│   ├── WordTableToExcel.Shim32.dll / Shim64.dll   Chargeurs natifs (C) pour Office 32 / 64 bits
│   ├── install.cmd / uninstall.cmd                 Installation pour l'utilisateur courant (HKCU), sans droits administrateur
│   ├── LISEZMOI.txt              Mode d'emploi du complément
│   ├── INFORMATIQUE.txt          Note au service informatique (signature Authenticode des DLL)
│   └── signer-les-dll.ps1        Script de signature (service informatique)
├── build/
│   ├── build-installation.sh     Linux/WSL : tests, DLL, exe, chargeurs natifs
│   ├── build-shim.sh             Compile shim.c en 32 et 64 bits (MinGW-w64)
│   └── make-icon.py              Génère TableauxWordExcel.ico (Pillow)
├── src/
│   ├── WordTableToExcel/         Bibliothèque partagée + complément (net40, signée : WordTableToExcel.snk)
│   │   ├── AddIn/                Complément uniquement (non compilé dans l'application)
│   │   │   ├── Connect.cs            Point d'entrée COM (IDTExtensibility2, IRibbonExtensibility), clics du ruban
│   │   │   ├── Ribbon.xml            Boutons « Tableaux vers Excel » / « Importer depuis Excel » (Accueil et Références)
│   │   │   ├── LegacyToolbar.cs      Barre d'outils Word 2000-2003 (sans ruban)
│   │   │   ├── OfficeInterop.cs      Interfaces Office déclarées à la main (pas de PIA)
│   │   │   └── ShimEntryPoint.cs     Entrée appelée par le chargeur natif (crée Connect, renvoie son IUnknown)
│   │   ├── Core/                 Code pur : ni Word, ni Windows, ni interface — entièrement testable
│   │   │   ├── Captions/
│   │   │   │   ├── CaptionMatcher.cs     Reconnaissance multilingue d'une légende (champ SEQ, texte « Tabl… », style Légende)
│   │   │   │   ├── CaptionAssigner.cs    Attribution des légendes (convention détectée ou imposée, jamais deux fois la même)
│   │   │   │   └── SheetNameBuilder.cs   Noms de feuilles Excel valides et uniques (31 caractères, caractères interdits)
│   │   │   ├── Layout/
│   │   │   │   ├── TableLayout.cs        Grille d'un tableau : cellules positionnées, fusions, largeurs, hauteurs, lignes supprimées…
│   │   │   │   ├── WordXmlTableParser.cs Analyse du XML d'un tableau Word (structure, styles de tableau, bordures) ; chargement allégé
│   │   │   │   ├── WordTableStyles.cs    Styles de tableau Word : zones conditionnelles, héritage, OoxmlXml (outils XML)
│   │   │   │   ├── WordXmlContent.cs     Texte et mise en forme effective des caractères lus dans le XML (styles, thème, champs, révisions)
│   │   │   │   └── WidthGridBuilder.cs   Grille reconstruite à partir des largeurs de cellules (sans XML)
│   │   │   ├── Model/                Modèle intermédiaire : TableModel, CellModel, TextRun, RunFormat, Rgb, bordures
│   │   │   ├── Text/
│   │   │   │   ├── CellTextSanitizer.cs  Nettoyage du texte (champs, caractères de contrôle, limite 32 767, comparaison)
│   │   │   │   └── NumberParser.cs       Texte → nombre Excel + format qui conserve l'apparence (« 1 234,50 », « 12,5 % »…)
│   │   │   ├── Xlsx/                 Écriture du classeur
│   │   │   │   ├── XlsxWorkbookWriter.cs Feuilles, texte enrichi, fusions, hauteurs, volets figés, feuille « Sommaire »
│   │   │   │   ├── XlsxStyleRegistry.cs  Styles dédupliqués (polices, fonds, bordures, formats, xf)
│   │   │   │   ├── ExcelUnits.cs         Points ↔ largeur de colonne Excel, références A1
│   │   │   │   └── ZipWriter.cs          Archive ZIP minimale (deflate, CRC32)
│   │   │   └── ExcelImport/          Lecture d'un classeur et conversion en tableau Word
│   │   │       ├── ZipReader.cs          Lecture ZIP (répertoire central, deflate, ZIP64, CRC, limites de taille)
│   │   │       ├── XlsxPackage.cs        Relations OPC, chargement XML sûr (DTD interdites), erreurs utilisateur
│   │   │       ├── XlsxWorkbook.cs       Classeur : feuilles, noms définis (zone d'impression, titres), styles, thème
│   │   │       ├── XlsxSheet.cs          Feuille : cellules (lecture en flux), fusions, lignes/colonnes, tableaux Excel
│   │   │       ├── XlsxStyles.cs         Formats de cellule, polices, remplissages, bordures ; ExcelColors.cs : couleurs
│   │   │       ├── ExcelNumberFormat.cs  Moteur de formats de nombre Excel (sections, conditions, dates, fractions, régional)
│   │   │       ├── ExcelDecimal.cs, ExcelDates.cs, ExcelLocaleTexts.cs   Arrondis Excel exacts, dates, textes localisés
│   │   │       ├── ExcelFormula.cs       Évaluateur de formules simples (mise en forme conditionnelle)
│   │   │       ├── ExcelConditionalFormatting.cs, ExcelTableStyles.cs   Mise en forme conditionnelle, styles de tableau Excel
│   │   │       ├── SheetConverter.cs     Feuille → TableModel (plage, valeurs affichées, mise en forme, largeur de page)
│   │   │       ├── WordTableXmlWriter.cs TableModel → tableau WordprocessingML (paquet Flat OPC pour InsertXML)
│   │   │       └── CellReference.cs      Références et plages A1
│   │   ├── Word/                 Accès à Word (liaison tardive) — lecture pour l'export, insertion pour l'import
│   │   │   ├── WordTableReader.cs    Lecture d'un tableau : XML (rapide) ou cellule par cellule (COM)
│   │   │   ├── WordRunReader.cs      Texte + mise en forme d'une plage par dichotomie sur les propriétés de police
│   │   │   ├── WordCaptionScanner.cs Paragraphes voisins d'un tableau → légendes candidates
│   │   │   ├── WordRevisions.cs      Suivi des modifications (suppressions) et renvois de notes (NOTEREF) à exclure
│   │   │   ├── WordColor.cs          Couleurs Word (RVB, automatique, thème, surlignage)
│   │   │   ├── WordDocumentGuard.cs  Restaure l'indicateur « document modifié » après la lecture
│   │   │   ├── WordTableInserter.cs  Insertion d'un tableau importé (InsertXML, repli .docx) et de sa légende
│   │   │   └── WordCom.cs            Conversions des valeurs COM, libération des références, IExportProgress
│   │   ├── Export/
│   │   │   ├── ExportService.cs      Déroulement complet de l'export (voir § 4)
│   │   │   └── ExportPlan.cs         TableEntry, sélection A/B et tableaux décochés, noms de feuilles, (ré)attribution des légendes
│   │   ├── Import/
│   │   │   ├── ImportService.cs      Déroulement complet de l'import (voir § 5)
│   │   │   ├── ExcelFileConverter.cs .xls/.xlsb/.ods/.csv → .xlsx par Excel en arrière-plan (s'il est installé)
│   │   │   └── TextMeasurer.cs       Largeur réelle d'un texte (GDI) pour dimensionner les colonnes
│   │   ├── UI/                   Boîtes de dialogue WinForms (partagées par le complément et l'application)
│   │   │   ├── ExportDialog.cs       Position des légendes, choix A/B, liste à cocher (pages, feuilles, légendes), options
│   │   │   ├── ImportDialog.cs       Feuilles à importer, plages, légendes, options
│   │   │   ├── ProgressDialog.cs     Progression modale, annulation, pompe de messages
│   │   │   ├── SelectAllHeader.cs    Case « Tout » dans l'en-tête d'une colonne de cases à cocher
│   │   │   ├── Messages.cs           Boîtes de message (information, avertissement, question, erreur)
│   │   │   └── WindowOwner.cs        Fenêtre Word propriétaire des boîtes de dialogue (complément)
│   │   ├── Infrastructure/
│   │   │   ├── Log.cs                Journal %LOCALAPPDATA%\WordTableToExcel\WordTableToExcel.log (1 Mo, 1 archive)
│   │   │   └── Settings.cs           Préférences HKCU\Software\WordTableToExcel
│   │   └── Properties/AssemblyInfo.cs
│   ├── TableauxWordExcel/        Application autonome (WinExe net40, AnyCPU, compile aussi Core/Word/Export/Import/UI/Infrastructure)
│   │   ├── Program.cs            Instance unique, nettoyage d'un Word invisible orphelin, filtre OLE, --selftest
│   │   ├── app.manifest          asInvoker (jamais d'élévation), version
│   │   ├── TableauxWordExcel.ico
│   │   └── App/
│   │       ├── MainForm.cs           Fenêtre principale : documents, export, point d'insertion, import
│   │       ├── WordInstances.cs      Instances et documents de Word vus depuis un autre processus
│   │       ├── DocumentList.cs       Fusion documents ouverts / fichiers ajoutés
│   │       ├── InsertionInfo.cs      Description du point d'insertion (texte de l'encadré)
│   │       ├── HiddenWord.cs         Word invisible réutilisable pour les fichiers non ouverts
│   │       ├── HiddenWordRecord.cs   Trace du Word invisible (reprise après un arrêt brutal)
│   │       ├── OleMessageFilter.cs   Filtre de messages OLE (Word occupé → nouvel essai, puis abandon propre)
│   │       ├── ComErrors.cs          Codes d'erreur COM, messages clairs
│   │       ├── NativeMethods.cs      Appels Win32 (fenêtres, accessibilité)
│   │       └── SelfTest.cs           Autotest sans Word (CI), captures PNG
│   └── WordTableToExcel.Shim/
│       ├── shim.c                Chargeur natif du complément (voir § 7)
│       └── shim.def
├── tests/
│   ├── WordTableToExcel.Tests/   Tests xunit (.NET 8) — voir § 10
│   │   └── Fakes/                Faux modèle objet de Word (FakeWord.cs, FakeInsertionWord.cs)
│   ├── fixtures/                 Documents et classeurs de test + scripts qui les génèrent
│   └── ci/                       Scripts PowerShell/VBScript de l'intégration continue (Windows)
└── .github/
    ├── workflows/build.yml       CI à chaque poussée et pull request
    ├── workflows/release.yml     Publication d'une version (tag v*)
    └── release-notes/            Notes de version (v1.0.0.md, v2.0.0.md, v2.1.0.md…)
```

## 3. Architecture en couches

```
AddIn (complément)      App (application autonome)       ← points d'entrée
        \                    /
         UI (WinForms : dialogues, progression, messages)
              |
   Export / Import (services : déroulement, compte rendu)
              |
   Word (modèle objet de Word en liaison tardive)          Infrastructure (journal, préférences)
              |
   Core (pur : XML, xlsx, formats, légendes, modèle)
```

Règles de dépendance :

- `Core` ne dépend de rien d'autre (ni COM, ni WinForms, ni registre) : c'est là que se trouve la logique
  délicate, testée sur toutes les plates-formes.
- `Word` ne fait que des appels COM en liaison tardive (`dynamic`) ; il se teste avec le faux Word des tests.
- `Export`/`Import` orchestrent et parlent à l'utilisateur via `UI`.
- `AddIn` n'est compilé que dans la DLL ; `App` que dans l'exe. L'exe compile les autres dossiers par liens
  (`TableauxWordExcel.csproj`, `EnableDefaultCompileItems=false`).

Tout s'exécute sur **le thread d'interface (STA)**. Les opérations longues tournent dans `ProgressDialog`, qui
pompe les messages (`Application.DoEvents`, au plus toutes les 100 ms) pour rester réactive et permettre
l'annulation : pas de thread de travail, donc pas de marshaling COM entre threads.

## 4. Export Word → Excel

### 4.1 Déroulement (`ExportService.Run`)

1. `WordDocumentGuard` mémorise `Document.Saved` (restauré à la fin, même en cas d'erreur).
2. `Tables.Count` = 0 → message et fin.
3. `WordRevisions.DocumentHasRevisions` : **une seule question** à Word. Sans modification suivie, aucune
   recherche de suppressions n'est faite ensuite (Word peut être lent sur `Range.Revisions`).
4. **Analyse** (`ScanTables`) — dans une fenêtre de progression annulable au-delà de
   `ScanProgressThreshold` (15) tableaux, sinon avec le sablier : pour chaque tableau,
   - tableau entièrement supprimé en suivi des modifications (`WordRevisions.IsDeletedTable`) → écarté et compté ;
   - légendes candidates au-dessus et au-dessous (`WordCaptionScanner`, § 4.2) ;
   - pages de début et de fin (`Range.Information(wdActiveEndPageNumber)`).
   Puis attribution des légendes (`ExportPlan.AssignCaptions`) avec la position mémorisée.
5. **Boîte de dialogue** (`ExportDialog`) : position des légendes (Automatique — avec la position détectée
   affichée —, au-dessus, au-dessous ; changer de choix réattribue aussitôt les légendes, sans relire Word),
   option A/B, tableaux à décocher, options (légende en A1, feuille Sommaire, conversion des nombres).
6. `ExportPlan.Build` : tableaux retenus et noms de feuilles (identique pour l'aperçu et pour l'export).
7. Fenêtre d'enregistrement (`AskTargetPath`, le document Word lui-même est refusé comme destination).
8. **Lecture et écriture** (`Export`, dans `ProgressDialog`) : pour chaque tableau, `WordTableReader.Read`
   (§ 4.3), puis `XlsxWorkbookWriter.AddTable` (§ 4.4). Un tableau illisible est ajouté aux échecs et l'export
   continue. Les références COM sont libérées au fil de l'eau (§ 4.6).
9. `SaveSafely` : écriture dans `~xxxxxxxx.tmp` à côté de la cible, suppression de l'ancien fichier, renommage.
10. **Compte rendu** (`ShowReport`, § 4.5).

### 4.2 Légendes

- **`WordCaptionScanner`** examine le premier paragraphe non vide au-dessus et au-dessous du tableau (au plus
  `MaxEmptyParagraphs` = 3 paragraphes vides ignorés). Il s'arrête sur un autre tableau. Le texte lu exclut les
  passages supprimés en suivi des modifications (`WordRevisions.VisibleText`) — sauf si le document n'en a
  aucun (`CheckRevisions = false`). Un champ `SEQ` de tableau → légende certaine ; un autre `SEQ` (Figure…) →
  pas une légende ; sinon `CaptionMatcher.LooksLikeCaptionText` (libellé connu, ou préfixe « Tabl… » suivi
  d'un numéro, ou style *Légende*).
- **`CaptionMatcher`** : libellés intégrés multilingues + libellé local de Word (`CaptionLabels(wdCaptionTable)`)
  + libellés supplémentaires (`ExtraCaptionLabels`).
- **`CaptionAssigner`** : convention détectée (`DetectConvention` : vote des tableaux dont la légende n'est
  possible que d'un côté ; à égalité, au-dessus ; si tout est ambigu, les champs SEQ départagent) ou **imposée
  par l'utilisateur**. 1er passage : côté de la convention ; 2e passage : l'autre côté, si la légende n'est pas
  celle d'un autre tableau du côté de la convention. Une légende n'est jamais attribuée deux fois.
  Cas typique : « légende 1, tableau 1, tableau 2, légende 2, tableau 3 » — la légende 2 va au tableau 3 si les
  légendes sont au-dessus, au tableau 2 si elles sont au-dessous.
- **`SheetNameBuilder`** : nom de feuille à partir de la légende, sinon `Tableau_N` (N = rang dans le document).

### 4.3 Lecture d'un tableau (`WordTableReader.Read`)

**Structure** — `TryReadXmlLayout` demande `Range.WordOpenXML` (Word 2007+ ; paquet « Flat OPC » complet :
document, styles, thème, polices, paramètres, images en base64…), à défaut `Range.XML` (WordprocessingML 2003).

- `WordXmlTableParser.Load` ne construit en mémoire que les parties utiles (document principal, `styles.xml`,
  thème) ; les autres sont sautées pendant la lecture. Si le tableau n'est pas trouvé dans ce qui reste
  (paquet inhabituel), le XML est relu en entier.
- `WordXmlTableParser.Parse` : grille (`tblGrid`), `gridSpan`, `vMerge`, `gridBefore`/`gridAfter`, largeurs,
  hauteurs (`trHeight`, exacte ou « au moins »), **lignes supprimées** (`w:trPr/w:del`, retirées de la grille
  mais gardées dans `SourceRowCount`/`DeletedRows` pour rester aligné sur Word), **lignes d'en-tête répétées**
  (`tblHeader` en tête → `HeaderRowCount`), **images et objets** (`CountObjects` : `w:drawing`, `w:pict`,
  `w:object`, sans compter `mc:Fallback`, les objets imbriqués ni le contenu supprimé), fonds et bordures de
  chaque cellule avec **styles de tableau** (zones conditionnelles : première/dernière ligne et colonne, bandes,
  coins ; `tblLook` ; héritage `basedOn` ; conflits de bordures entre cellules voisines).
- Contenu (`WordXmlContentReader`, utilisé par l'application : `ReadContentFromXml = true`) : pour chaque
  cellule, segments de texte avec leur mise en forme **effective**, résolue dans l'ordre de Word :
  `docDefaults` < style de tableau (zones conditionnelles) < style de paragraphe < style de caractère <
  mise en forme directe ; propriétés « bascule » (gras, italique…) ; polices et couleurs de thème. Sont exclus :
  texte masqué, codes de champ (résultat seul), **suppressions suivies** (`w:del`, `w:moveFrom`, marques de
  paragraphe supprimées, lignes supprimées), **renvois de notes** (champ `NOTEREF`, texte au style « Appel de
  note »). Les symboles (`w:sym`), tabulations, sauts de ligne et équations (texte linéaire) sont gérés.

**Contrôle** — `XmlTextMatchesWord` compare le texte du XML à `Table.Range.Text` (une seule question à Word),
après `CellTextSanitizer.Comparable` (sans codes de champ, caractères de contrôle ni espaces). Selon l'affichage,
Word inclut ou non le texte supprimé et le texte masqué : quatre variantes du texte XML sont acceptées
(`TableTexts`). Au moindre écart : lecture cellule par cellule.

**Lecture cellule par cellule (COM)** — utilisée par le complément, et par l'application en repli :

- `EnumerateCells` (cellules du niveau du tableau), `Pair` (appariement cellules COM ↔ cellules de la grille
  XML, lignes supprimées comprises) ; sans XML ou si l'appariement échoue : `WidthGridBuilder` reconstruit la
  grille à partir des largeurs.
- `ExcludedIntervals` : plages supprimées (`Revisions` de type suppression, déplacement, conflit) et résultats
  des champs `NOTEREF`, soustraits de chaque cellule (`WordRevisions.Subtract`).
- `WordRunReader.Read` : la police d'une plage est interrogée d'un bloc ; si une propriété est hétérogène
  (`wdUndefined`), la plage est coupée en deux (dichotomie, profondeur ≤ 48). Coût proportionnel au nombre de
  changements de mise en forme. Quand positions et caractères ne correspondent plus (champs, texte masqué),
  lecture paragraphe par paragraphe.
- Une cellule illisible est relue en texte brut (remarque dans le compte rendu), jamais d'échec du tableau.

**Nettoyage** — `CellTextSanitizer.Clean` : codes de champ, caractères interdits en XML, sauts de ligne,
fin de cellule, et **limite de 32 767 caractères** d'une cellule Excel (le dépassement est signalé :
`out truncated` → `TableModel.Omissions`).

### 4.4 Écriture du classeur (`XlsxWorkbookWriter`)

- `AddTable` : cellules placées dans une grille (`SortedDictionary`), fusions, bordures réparties sur les cases
  d'une zone fusionnée, texte simple ou **enrichi** (chaînes partagées dédupliquées), conversion des nombres
  (`NumberParser`, seulement si la mise en forme de la cellule est uniforme), hauteurs de ligne (exactes de Word,
  ou estimées pour les cellules fusionnées qu'Excel n'ajuste pas), **volets figés** sous les lignes d'en-tête
  répétées de Word (et la légende en A1 le cas échéant ; rien si tout le tableau est en-tête). Le XML de la
  feuille est produit aussitôt (le `TableModel` peut être libéré).
- `XlsxStyleRegistry` déduplique polices, remplissages, bordures, formats et `xf`.
- **Feuille « Sommaire »** (`IncludeSummary`, à partir de deux feuilles) : construite à l'enregistrement, placée
  en premier et active ; n°, légende complète, pages, lien interne (`<hyperlink location="'Feuille'!A1">`, sans
  relation) ; nom « Sommaire », ou « Sommaire (2) »… en cas de collision. À l'import, une feuille ainsi nommée
  n'est pas cochée par défaut (`IsSummarySheetName`).
- `Save` : `[Content_Types].xml`, relations, propriétés, `workbook.xml`, styles, chaînes partagées, feuilles,
  dans un `ZipWriter` (deflate, ou stockage si la compression n'apporte rien).

### 4.5 Compte rendu

`ExportReport` distingue :

| Liste | Contenu | Affichage |
|---|---|---|
| `Failures` | Tableaux non exportés (erreur de lecture) | En tête : « Attention : N tableau(x) non exporté(s) » (10 au plus) |
| `Omissions` | Informations des tableaux exportés absentes du classeur (`TableModel.Omissions`) : images et objets, texte coupé | « Informations absentes du classeur » (10 au plus) |
| `Warnings` | Remarques (`TableModel.Warnings`) : cellule lue en texte brut, grille reconstruite, tableau dont toutes les lignes sont supprimées… | « Remarques » (5 au plus) |

S'il y a des échecs ou des omissions, la question finale porte l'icône d'avertissement. **Toutes** les lignes
sont écrites dans le journal (« Rapport d'export : … »).

### 4.6 Mémoire et performance

Mesures (tableau synthétique de 8 colonnes, XML proche de celui de Word) : pic de l'ordre de 190 Mo pour
1 000 lignes, 500 Mo pour 5 000 lignes, 1 à 1,5 Go pour 20 000 lignes — le temps de lire **un** tableau
(chaîne XML UTF-16 + arbre XML + modèle). Entre deux tableaux, seul le XML des feuilles et les chaînes partagées
restent en mémoire (jusqu'à ~1 Ko par cellule de texte enrichi unique).

| Mesure | Où | Valeur |
|---|---|---|
| Parties inutiles du paquet XML non chargées | `WordXmlTableParser.Load` | document, styles, thème seulement |
| Recherche des suppressions seulement si le document en a | `WordRevisions.DocumentHasRevisions` | 1 appel par export |
| Libération des références COM (RCW) | `WordCom.ReleaseUnusedReferences` (`GC.Collect` + `WaitForPendingFinalizers`) | après l'analyse ; après chaque tableau lu cellule par cellule ; tous les 20 tableaux (`ReleaseEveryTables`) ; toutes les 500 cellules lues via COM (`ReleaseEveryCells`) |
| Analyse avec progression | `ExportService.ScanProgressThreshold` | au-delà de 15 tableaux |

Sans libération, chaque objet renvoyé par Word (plage, police, cellule) garde un objet vivant **dans Word**
jusqu'au passage du ramasse-miettes .NET, déclenché par la mémoire de .NET et non par celle de Word : Word
grossit pendant les longs exports. Le complément tourne dans la mémoire de Word (souvent 32 bits) : pour les
très gros documents, l'application (64 bits) est préférable.

## 5. Import Excel → Word

1. `ImportService.Run` : document vérifié (mode protégé, document marqué comme final, protection contre les
   modifications : signalés sans rien modifier), choix du classeur.
2. Formats autres que `.xlsx`/`.xlsm`/`.xltx`/`.xltm` : `ExcelFileConverter` les fait convertir par Excel
   (instance invisible, sans macros), s'il est installé.
3. `XlsxWorkbook.Load` : fichier lu en lecture partagée (il peut rester ouvert dans Excel), 200 Mo au plus
   (`MaxFileSize`) ; `ZipReader` (entrée ≤ 512 Mo, CRC contrôlé) ; styles, thème, palette, noms définis.
4. `Analyze` : pour chaque feuille, lecture (`XlsxSheet`, lecture en flux, 2 000 000 de cellules au plus —
   au-delà `Truncated`, signalé), plage détectée (zone d'impression, sinon cellules remplies et mises en
   forme, fusions, tableaux Excel), légende détectée en tête de feuille (« Tableau 3 : … » suivie d'une ligne
   vide, comme dans les classeurs exportés). Feuille cochée par défaut si visible, importable et pas « Sommaire ».
5. `ImportDialog` : feuilles cochées, plages, textes de légende, options.
6. Pour chaque feuille : `SheetConverter` → `TableModel` avec **le texte qu'Excel affiche**
   (`ExcelNumberFormat` : sections, conditions, couleurs, dates, durées, fractions, formats régionaux ; arrondis
   sur 15 chiffres significatifs avec `ExcelDecimal`), mise en forme (styles de cellule, styles de tableau
   Excel, mise en forme conditionnelle évaluée avec `ExcelFormula`), lignes/colonnes masquées, largeur ajustée
   à la page (`TextMeasurer`), 63 colonnes et 30 000 cellules au plus par tableau.
7. `WordTableXmlWriter` produit un paquet Flat OPC avec une mise en forme entièrement explicite ;
   `WordTableInserter` l'insère d'un bloc (`Range.InsertXML`), avec repli par un `.docx` temporaire
   (`InsertFile`), puis ajoute la légende par la fonction *Insérer une légende* de Word (champ `SEQ`), avec
   `[Titre du tableau]` surligné si aucun texte n'est saisi. Un paragraphe vide sépare toujours deux tableaux.
8. Tout est regroupé dans un seul enregistrement d'annulation (`UndoRecord`, Word 2010+) et `ScreenUpdating`
   est coupé pendant l'insertion.

## 6. Application autonome

- **`Program`** : `--selftest` → `SelfTest.Run` ; sinon instance unique (mutex `Local\TableauxWordExcel-…` ;
  une deuxième instance ramène la première au premier plan), `HiddenWord.CleanUpOrphan`, gestionnaires
  d'erreurs globaux, filtre de messages OLE, `MainForm`.
- **`MainForm`** : liste des documents ouverts dans Word et des fichiers ajoutés (bouton, glisser-déposer,
  arguments de la ligne de commande), actualisée à chaque retour dans la fenêtre (`OnActivated`) et par F5 ;
  encadré du point d'insertion ; boutons d'export et d'import. `RunAction` exécute une action à la fois
  (`_busy`), avec une attente maximale de 10 s d'un Word occupé, et traduit les erreurs COM en messages clairs.
- **`WordInstances`** : fenêtres principales **visibles** de Word (`OpusApp`) dans l'ordre d'empilement →
  fenêtre de document (`_WwG`) → `AccessibleObjectFromWindow(OBJID_NATIVEOM)` → `Window.Application`. Cela
  fonctionne avec plusieurs instances de Word ; à défaut, `GetActiveObject("Word.Application")` (seulement si
  visible). Un Word invisible n'est donc jamais proposé (ni pour l'export, ni pour l'import).
- **`OleMessageFilter`** : Word refuse les appels quand il est occupé (`RPC_E_CALL_REJECTED`,
  `SERVERCALL_RETRYLATER`) ; le filtre réessaie toutes les 200 ms jusqu'à l'attente maximale
  (10 s pour une action, 1,5 s pour l'actualisation automatique), puis abandonne proprement.
- **`HiddenWord`** (fichiers non ouverts dans Word) : instance de Word **distincte et invisible**
  (`Visible = false`, `DisplayAlerts = none`, `AutomationSecurity = ForceDisable` : aucune macro,
  `ScreenUpdating = false`) ; fichier ouvert en lecture seule, sans liste des fichiers récents, avec un faux mot
  de passe (un document protégé renvoie une erreur au lieu d'afficher une invite invisible). L'instance est
  **gardée 3 minutes** après un export (`IdleLifetime`) pour que le suivant démarre aussitôt ; une minuterie
  (5 s) la surveille hors opération :
  - un document que l'utilisateur ouvre entre-temps (double-clic dans l'Explorateur) peut arriver dans cette
    instance : elle est alors **rendue visible** avec ses réglages habituels et laissée à l'utilisateur
    (`HandOver`) — jamais refermée sans enregistrer ;
  - elle est refermée après inactivité, et à la fermeture de l'application ;
  - son processus est noté dans `%LOCALAPPDATA%\WordTableToExcel\word-invisible.txt` (numéro + heure de
    démarrage, `HiddenWordRecord`) ; si l'application s'arrête brutalement, le démarrage suivant ferme ce Word
    **seulement** s'il s'agit bien de lui (programme WINWORD, même heure de démarrage, aucune fenêtre visible).
- **`SelfTest`** (sans Word, en CI) : filtre OLE, nettoyage d'une trace qui ne désigne pas Word (aucun
  processus fermé), recherche de Word, fenêtre principale, boîte d'export (case « Tout », tableau décoché,
  **position des légendes** : automatique ↔ au-dessous imposée), lecture du classeur de test et boîte d'import ;
  captures PNG.

## 7. Complément Word (gelé)

**Statut : gelé.** Il reste compilé, testé en CI et livré dans `Installation/`, et bénéficie automatiquement des
améliorations du code partagé (Core, Word, Export, Import, UI). Aucune nouvelle fonction propre au complément
(ruban, installation, chargeur) n'est développée : l'application autonome est le produit recommandé, notamment
sur les postes où Word exige des compléments signés.

- **`Connect`** : `IDTExtensibility2` (chargement, déchargement) et `IRibbonExtensibility` (`Ribbon.xml`) ;
  clics « Tableaux vers Excel » / « Importer depuis Excel » ; toute exception est interceptée.
  Le contenu des tableaux est lu cellule par cellule (appels COM dans le même processus, rapides).
- **`LegacyToolbar`** : Word 2000-2003 (barre d'outils Standard, sans modifier le modèle Normal).
- **`OfficeInterop`** : interfaces COM d'Office déclarées à la main (pas de PIA).
- **Chargeur natif `shim.c`** : Windows n'active pas un serveur COM .NET (`mscoree.dll`) enregistré pour le seul
  utilisateur ; le chargeur (DLL native 32 et 64 bits, enregistrée dans HKCU) démarre le CLR 4
  (`ICLRMetaHost` → `ICLRRuntimeHost::ExecuteInDefaultAppDomain`) et appelle `ShimEntryPoint`, qui crée
  `Connect` et renvoie son `IUnknown`.
- **`install.cmd`** : copie dans `%LOCALAPPDATA%\WordTableToExcel`, retrait de la marque « provenant
  d'Internet », enregistrement COM (vues 32 et 64 bits de HKCU), déclaration `Office\Word\Addins`, détection des
  stratégies de signature. **`uninstall.cmd`** défait tout.
- **Signature** : DLL signée par nom fort (`WordTableToExcel.snk`) ; signature Authenticode par le service
  informatique (`INFORMATIQUE.txt`, `signer-les-dll.ps1`), vérifiée en CI avec un certificat de test.

## 8. Réglages, journal et fichiers écrits sur le poste

Préférences dans `HKEY_CURRENT_USER\Software\WordTableToExcel` (partagées par le complément et l'application) :

| Valeur | Type | Rôle |
|---|---|---|
| `AllTables` | DWORD | Dernier choix A (0) / B (1) |
| `IncludeCaptionRow` | DWORD | Légende en A1 (1) ou non (0) |
| `ExportSummary` | DWORD | Feuille « Sommaire » (1, par défaut) ou non (0) |
| `ConvertNumbers` | DWORD | Conversion des nombres (1) ou non (0) |
| `ExportCaptionPosition` | DWORD | Position des légendes : automatique (0), au-dessus (1), au-dessous (2) |
| `LastFolder` | Chaîne | Dernier dossier d'enregistrement |
| `ExtraCaptionLabels` | Chaîne | Libellés de légende supplémentaires, séparés par `;` (ex. `Annexe;Tab.`) |
| `ImportAddCaption` | DWORD | Import : ajouter une légende (1) ou non (0) |
| `ImportCaptionBelow` | DWORD | Import : légende sous le tableau (1) ou au-dessus (0) ; absente = selon le document |
| `ImportFitToPage` | DWORD | Import : réduire les tableaux trop larges (1) ou non (0) |
| `ImportSkipHidden` | DWORD | Import : ignorer les lignes et colonnes masquées (1) ou non (0) |
| `ImportGridlines` | DWORD | Import : quadrillage des cellules sans bordure (1) ou non (0) |
| `ImportLastFolder` | Chaîne | Import : dernier dossier ouvert |
| `AppTopMost` | DWORD | Application : fenêtre toujours visible (1, par défaut) ou non (0) |
| `AppBounds` | Chaîne | Application : position et taille de la fenêtre (`x,y,largeur,hauteur`) |
| `AppLastWordFolder` | Chaîne | Application : dossier du dernier fichier Word ajouté |

Fichiers dans `%LOCALAPPDATA%\WordTableToExcel` : `WordTableToExcel.log` (journal, 1 Mo puis archivé en
`.log.1`), `word-invisible.txt` (application, tant qu'un Word invisible tourne), et pour le complément ses
DLL. Le journal contient des chemins de fichiers, des noms de feuilles et des messages d'erreur ; il reste sur le
poste.

## 9. Compilation

Prérequis : [SDK .NET 8](https://dotnet.microsoft.com/download) (les projets ciblent .NET Framework 4.0 grâce
aux assemblies de référence NuGet, compilables sous Linux) ; MinGW-w64 pour le chargeur natif.

```bash
dotnet test tests/WordTableToExcel.Tests          # tests (Windows, Linux ou macOS)
dotnet build src/WordTableToExcel -c Release      # WordTableToExcel.dll (net40, AnyCPU, signée)
dotnet build src/TableauxWordExcel -c Release     # TableauxWordExcel.exe (net40, AnyCPU, Prefer32Bit=false)
build/build-shim.sh sortie/                       # chargeurs natifs 32/64 bits
build/build-installation.sh                       # tout, et recopie dans Installation/ et Application/
python3 build/make-icon.py                        # icône
```

Sous Windows, `build.cmd` exécute les tests et recopie la DLL et l'exe. **Les binaires livrés sont suivis par
Git** (`Application/TableauxWordExcel.exe`, `Installation/*.dll`) : après toute modification du code, les
recompiler et les recopier dans le même commit (la CI vérifie aussi l'exe du dépôt).

## 10. Tests

Projet `tests/WordTableToExcel.Tests` (xunit, .NET 8), environ 350 tests exécutables sans Word ni Windows. Il
compile par liens `Core/`, `Word/`, `Export/ExportPlan.cs`, `Infrastructure/Log.cs` et les fichiers purs de
l'application (`DocumentList`, `InsertionInfo`, `ComErrors`, `HiddenWordRecord`).

| Fichier | Couverture |
|---|---|
| `WordXmlTableParserTests` | Grilles, fusions, styles de tableau, bordures, fichiers réels (`fixtures/`) |
| `WordXmlContentTests` | Mise en forme effective (priorité des styles, bascules, thème), champs, texte masqué, suivi des modifications, renvois de notes, document complexe (13 tableaux comparés à LibreOffice) |
| `WordReaderTests` | Lecture via le faux Word : XML ou cellule par cellule, suppressions, renvois, tableaux supprimés, légendes, garde du document, bout en bout |
| `ExportSafetyTests` | Texte coupé, comptage des objets, chargement allégé du XML, omissions du compte rendu |
| `CaptionTests` | Reconnaissance, attribution, position imposée, noms de feuilles |
| `XlsxWriterTests`, `SummaryAndFreezeTests` | Paquet xlsx, styles, texte enrichi, fusions, Sommaire, volets figés |
| `TextTests` | Nettoyage du texte, reconnaissance des nombres |
| `WidthGridBuilderTests` | Grille reconstruite à partir des largeurs |
| `ExcelImportTests`, `ExcelNumberFormatTests`, `ExcelFormulaTests` | Lecture de classeurs, formats de nombre (comparés à LibreOffice), formules, aller-retour |
| `WordTableInserterTests` | Point d'insertion et insertion via un faux Word |
| `AppLogicTests` | Liste des documents, point d'insertion, erreurs COM, trace du Word invisible |

**Faux Word** (`Fakes/FakeWord.cs`) : un document est une suite de caractères avec leur mise en forme ; il imite
les positions de Word (marques de fin de cellule d'une position rendues « \r\a », codes de champ invisibles),
`Range`, `Font` (valeurs indéfinies sur une plage hétérogène), `Paragraphs`, `Tables`, `Cells`, `Fields`
(avec `Result`), `Revisions`, et compte les appels (`RangeCalls`, `FontCalls`, `RevisionsCalls`) pour vérifier
les optimisations. `FakeDocumentBuilder` construit tableaux, paragraphes, révisions et renvois.

**Données** (`tests/fixtures/`) : `medium_shading_merged.*` et `make_fixtures.py` (tableaux réels),
`rapport_test_complexe.docx` (styles intégrés, champs, fusions, tableau imbriqué),
`import_complexe*.xlsx` et `make_import_fixture.py` (classeur de test, recalculé par LibreOffice),
`libreoffice_word2003.xml` (format Word 2003).

Ce que les tests ne couvrent pas (pas de Word en CI) : le dialogue réel avec Word. Avant une version, faire une
**recette manuelle** sur un poste équipé de Word : document avec suivi des modifications et notes de bas de
page, gros document (> 100 tableaux), légendes enchaînées (au-dessus et au-dessous), fichier non ouvert (deux
exports de suite), import d'un classeur exporté.

## 11. Intégration continue et publication

**`build.yml`** (chaque poussée et pull request) :

1. *Chargeur natif* (Linux, MinGW-w64) : compilation 32 et 64 bits.
2. *Windows* : tests ; compilation de la DLL ; import vérifié sur .NET Framework dans Windows PowerShell
   (`framework-check.ps1`) ; **installation réelle** depuis `Installation/` (fichiers marqués « provenant
   d'Internet », `install.cmd` sans droits, registre, activation COM et lecture du ruban par un client natif
   en 32 et 64 bits depuis un compte standard, `uninstall.cmd` — `com-smoke-test.ps1`) ; signature et
   stratégies de Word (`enterprise-check.ps1`) ; compilation de l'exe ; autotest de l'exe compilé **et** de
   celui du dépôt (`app-selftest.ps1` : fichier unique, manifeste, icône, `--selftest`, démarrage réel, instance
   unique) ; installation depuis un dossier reconstruit à partir des sources.
3. Les captures de l'autotest sont publiées comme artefacts et imprimées dans le journal (lignes `PNG64`).

**`release.yml`** : déclenché par un tag `v*` (ou manuellement avec un tag) ; refait les vérifications puis crée
la release GitHub avec `TableauxWordExcel.exe` et `WordTableToExcel-<tag>.zip`, notes tirées de
`.github/release-notes/<tag>.md`.

**Publier une version** :

1. Numéro de version dans `src/WordTableToExcel/WordTableToExcel.csproj`,
   `src/TableauxWordExcel/TableauxWordExcel.csproj` (`Version`, `AssemblyVersion`, `FileVersion`) et
   `src/TableauxWordExcel/app.manifest` ;
2. notes dans `.github/release-notes/vX.Y.Z.md` ;
3. binaires recompilés et recopiés ; recette manuelle ; CI verte ;
4. fusion dans `main`, puis tag `vX.Y.Z` (déclenche la publication).

## 12. Conventions de code

- **Langue** : interface, messages, journal, commentaires et documentation en français. Les messages
  s'adressent à l'utilisateur (« Word est occupé… », jamais un code d'erreur seul).
- **Cible** : .NET Framework 4.0, C# 7.3. Pas d'`async`/`await`, pas d'`ExceptionDispatchInfo` (4.5), pas de
  dépendance NuGet à l'exécution. `dynamic` pour tout accès à Office.
- **COM** : chaque appel susceptible d'échouer est isolé (`try`/`catch` au plus près, valeur par défaut
  prudente, ligne dans le journal) ; valeurs converties par `WordCom.AsInt/AsString…` (`wdUndefined` = 9999999) ;
  aucune modification du document pendant l'export.
- **Commentaires** : un résumé `///` par classe et par méthode non triviale, qui explique le **pourquoi**
  (comportement de Word, limite d'Excel…). Constantes nommées pour les valeurs d'énumérations Office.
- **Tests** : toute correction de comportement s'accompagne d'un test ; logique pure dans `Core` ou dans un
  fichier sans dépendance pour rester testable.

## 13. Pièges connus de Word et de COM

- **`Range.Text` ≠ positions** : une marque de fin de cellule occupe une position mais donne « \r\a » ; les
  codes de champ, le texte masqué et les tableaux imbriqués décalent aussi texte et positions.
- **Suivi des modifications** : selon l'affichage (« Marques simples », « Toutes les marques », « Original »),
  `Range.Text` inclut ou non le texte supprimé. Les suppressions ne sont jamais exportées, acceptées ou non.
  `Range.Revisions` est lent sur les gros documents : on vérifie d'abord `Document.Revisions.Count`.
- **Renvois de notes** : un renvoi `NOTEREF` produit de vrais chiffres (souvent sans exposant) ; l'appel de note
  lui-même (`w:footnoteReference`) n'a pas de texte.
- **`Range.WordOpenXML`** renvoie un paquet complet (styles, thème, images…) pour chaque tableau : coûteux,
  d'où le chargement allégé.
- **`Information(wdActiveEndPageNumber)`** force la pagination : le premier appel peut prendre du temps.
- **Deux tableaux accolés fusionnent** : toujours un paragraphe entre eux à l'import.
- **Word occupé** (boîte de dialogue ouverte) : appels refusés → filtre de messages OLE.
- **Références COM (RCW)** : chacune garde un objet vivant dans Word jusqu'au ramasse-miettes → libération
  explicite dans les boucles longues.
- **Word invisible et documents de l'utilisateur** : Windows peut ouvrir un document double-cliqué dans une
  instance de Word lancée par automatisation → surveillance et restitution (`HiddenWord.HandOver`).
- **Composant COM .NET enregistré par utilisateur** : non activable par Windows → chargeur natif.
- **Libellé des légendes** : dépend de la langue de Word (`CaptionLabels(wdCaptionTable)`), d'où la
  reconnaissance multilingue.

## 14. Évolutions possibles et idées écartées

Évolutions possibles (issues des audits, non développées) :

- **Mise à jour des tableaux importés** : mémoriser la source (fichier, feuille, plage) dans le texte de
  remplacement du tableau Word et proposer « Mettre à jour les tableaux importés ».
- **Export de plusieurs documents** (ou d'un dossier) : un classeur par document, en réutilisant le Word invisible.
- **Liens hypertexte** des cellules exportés en vrais liens Excel (aujourd'hui : texte seul).
- **Lecture par blocs de lignes** des tableaux géants (pic mémoire borné) ; import : ne garder en mémoire que
  les feuilles cochées.

Écartées (trop complexes pour le gain) : export des images, regroupement de plusieurs tableaux sous une même
légende, tableaux sur plusieurs pages fusionnés.
