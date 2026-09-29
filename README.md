# Tableaux Word vers Excel

Complément COM pour Microsoft Word (Windows) qui échange des tableaux entre Word et Excel,
**une feuille par tableau**, en **conservant la mise en forme** :

- **Tableaux vers Excel** : exporte les tableaux du document ouvert vers un classeur Excel, sans jamais
  modifier le document Word.
  - **Option A** : uniquement les tableaux qui ont une légende (champ `SEQ Tableau`, `SEQ Table`, `SEQ Tabla`…
    ou texte commençant par « Tabl… »). Chaque feuille porte le nom de sa légende.
  - **Option B** : tous les tableaux du document. Les tableaux sans légende deviennent `Tableau_1`, `Tableau_2`…
- **Importer depuis Excel** : insère à l'emplacement du curseur un tableau Word par feuille choisie d'un
  classeur Excel, avec les **valeurs exactement telles qu'Excel les affiche** (nombres, monnaies, dates,
  pourcentages…) et leur mise en forme, et, en option, une **légende numérotée « Tableau N » à compléter**.

> **Installation en 3 clics — sans droits administrateur, sans rien compiler**
>
> 1. Sur la page GitHub du projet : **Code › Download ZIP**, puis clic droit sur le fichier › **Extraire tout**.
> 2. Fermez Word, ouvrez le dossier **`Installation`** extrait et double-cliquez sur **`install.cmd`**.
> 3. Ouvrez Word : boutons **Tableaux vers Excel** et **Importer depuis Excel** à droite des onglets
>    **Accueil** et **Références** (groupe **Excel**).

---

## Sommaire

1. [Compatibilité](#compatibilité)
2. [Installation](#installation)
3. [Utilisation](#utilisation)
4. [Détection des légendes](#détection-des-légendes)
5. [Mise en forme conservée](#mise-en-forme-conservée)
6. [Intégrité du document Word](#intégrité-du-document-word)
7. [Import depuis Excel](#import-depuis-excel)
8. [Désinstallation](#désinstallation)
9. [Dépannage](#dépannage)
10. [Compilation et tests](#compilation-et-tests)
11. [Architecture](#architecture)
12. [Réglages avancés](#réglages-avancés)

---

## Compatibilité

| Élément | Prise en charge |
|---|---|
| Word 2007, 2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365 | Boutons **Tableaux vers Excel** et **Importer depuis Excel** dans les onglets **Accueil** et **Références** du ruban |
| Word 2000, 2002 (XP), 2003 | Boutons **Tableaux → Excel** et **Excel → Tableaux** dans la barre d'outils *Standard* (import : nécessite le Pack de compatibilité Office 2007) |
| Office 32 bits et 64 bits, installation MSI ou « Démarrer en un clic » | Oui (chargeur 32 bits et 64 bits enregistrés tous les deux) |
| Droits nécessaires | **Aucun droit administrateur** : installation pour l'utilisateur courant |
| Windows | 7 SP1 et suivants, avec .NET Framework 4.x (inclus d'office depuis Windows 8) |
| Excel | **Non requis** : le classeur `.xlsx` est écrit (export) et lu (import) directement. Seuls les anciens formats (`.xls`, `.xlsb`, `.ods`, `.csv`) passent par Excel, s'il est installé, pour être convertis |

Le complément n'utilise aucune bibliothèque d'interopérabilité liée à une version d'Office (PIA) :
tout l'accès à Word se fait en liaison tardive, ce qui le rend indépendant de la version installée.

## Installation

Tout ce qu'il faut est **déjà dans le dépôt**, dossier [`Installation/`](Installation) : aucune
compilation, aucun outil de développement, **aucun droit administrateur**.

1. Sur la page GitHub du projet, cliquez sur **Code › Download ZIP**.
2. Clic droit sur le fichier `.zip` téléchargé › **Extraire tout…** (n'exécutez pas l'installation depuis
   l'intérieur du `.zip` : les fichiers doivent être extraits).
3. **Fermez Word.**
4. Ouvrez le dossier **`Installation`** et double-cliquez sur **`install.cmd`**.
   Si Windows affiche « Windows a protégé votre ordinateur », cliquez sur *Informations complémentaires*
   puis *Exécuter quand même* (le script est un simple fichier texte que vous pouvez lire avant).
5. Ouvrez Word : les boutons **Tableaux vers Excel** et **Importer depuis Excel** apparaissent à droite de
   l'onglet **Accueil** (et de l'onglet **Références**), dans le groupe **Excel**.

Contenu du dossier `Installation` :

| Fichier | Rôle |
|---|---|
| `WordTableToExcel.dll` | Le complément (.NET Framework 4) |
| `WordTableToExcel.Shim32.dll` | Chargeur natif pour Office 32 bits |
| `WordTableToExcel.Shim64.dll` | Chargeur natif pour Office 64 bits |
| `install.cmd`, `uninstall.cmd` | Installation / désinstallation pour l'utilisateur courant |
| `LISEZMOI.txt` | Mode d'emploi résumé |

`install.cmd` (un script `cmd` lisible, sans PowerShell) copie les trois DLL dans
`%LOCALAPPDATA%\WordTableToExcel`, retire la marque « fichier provenant d'Internet », enregistre le
complément pour l'utilisateur courant (`HKEY_CURRENT_USER`, Office 32 et 64 bits) et le déclare auprès de Word.

> **Pourquoi un chargeur natif ?** Windows refuse d'activer un composant COM écrit en .NET lorsqu'il est
> enregistré pour un seul utilisateur (vérifié dans l'intégration continue). Le petit chargeur
> `WordTableToExcel.Shim32/64.dll` (code C, dossier `src/WordTableToExcel.Shim`), lui, s'enregistre
> sans droits administrateur ; au chargement par Word il démarre .NET Framework 4 et crée le complément
> contenu dans `WordTableToExcel.dll`.

## Utilisation

1. Ouvrez le document Word qui contient les tableaux.
2. Cliquez sur **Accueil › Excel › Tableaux vers Excel**.
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

Lors d'un **export**, le document ouvert est **uniquement lu** (l'import, lui, insère des tableaux à
l'emplacement du curseur, sans rien supprimer ni remplacer, et s'annule d'un seul Ctrl+Z) :

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

## Import depuis Excel

1. Placez le curseur à l'endroit du document où les tableaux doivent être insérés.
2. Cliquez sur **Accueil › Excel › Importer depuis Excel** et choisissez un classeur
   (`.xlsx`, `.xlsm`, `.xltx`, `.xltm` ; `.xls`, `.xlsb`, `.ods`, `.csv` si Excel est installé pour les convertir).
   Le classeur peut rester ouvert dans Excel : il est lu, jamais modifié.
3. La boîte de dialogue liste les feuilles : cochez celles à importer (**un tableau par feuille**). Pour chacune :
   - **Plage** : détectée automatiquement (zone d'impression si elle existe, sinon cellules remplies et cellules
     voisines mises en forme, fusions et tableaux Excel compris) ; modifiable, par exemple `A3:F20` ;
   - **Texte de la légende** : texte placé après « Tableau N : ». Si la feuille commence par une légende
     (« Tableau 3 : Ventes… » en A1 suivie d'une ligne vide, comme dans les classeurs produits par l'export),
     elle est reprise et n'entre pas dans le tableau.
4. Options :
   - **Ajouter une légende numérotée** au-dessus ou au-dessous du tableau (position proposée : celle des
     légendes déjà présentes dans le document). La légende est créée par la fonction *Insérer une légende*
     de Word (style *Légende*, champ `SEQ`, numérotation automatique) ; sans texte saisi, elle contient
     **`[Titre du tableau]` surligné en jaune**, à compléter ;
   - **Réduire les tableaux trop larges pour la page** : les colonnes de texte cèdent d'abord (leurs mots
     passent à la ligne) ; les nombres, dates et montants ne sont jamais coupés ;
   - **Ignorer les lignes et colonnes masquées** (comme à l'impression) ;
   - **Ajouter un quadrillage gris** aux cellules sans bordure (automatique si la feuille imprime le quadrillage).
5. Cliquez sur **Importer**. Tout l'import s'annule d'un seul **Ctrl+Z** (Word 2010 et suivants).

### Valeurs importées

Chaque cellule reçoit **le texte qu'Excel affiche**, pas la valeur brute enregistrée : le moteur de formats
reproduit les règles d'Excel.

| Excel | Exemple de valeur | Texte dans Word |
|---|---|---|
| Format Standard (11 caractères au plus, notation scientifique au-delà) | `0,1+0,2` ; `123456789012` | `0,3` ; `1,23457E+11` |
| Nombres, séparateur de milliers, divisions par mille (`#,##0,"k"`) | `1234567,891` | `1 234 567,89` |
| Monnaies, formats comptables (`_-* # ##0,00 €_-`) | `-3` ; `0` | `-3,00 €` ; `- €` |
| Pourcentages, notation scientifique, fractions | `0,256` ; `12345` ; `3,14159` | `25,6%` ; `1,23E+04` ; `3 14/99` |
| Dates et heures, noms des mois et des jours, durées `[h]:mm` | `45366,75` ; `1,52` | `vendredi 15 mars 2024 18:00` ; `36:28` |
| Sections positif ; négatif ; zéro ; texte, conditions `[>1000]`, couleurs `[Rouge]` | `-1234,5` avec `# ##0,00;[Rouge](# ##0,00)` | `(1 234,50)` en rouge |
| Formats régionaux `[$€-40C]`, `[$-40C]dddd`, date système `[$-F800]` | | selon la langue indiquée |
| Valeurs logiques et erreurs | `TRUE`, `#VALUE!` | `VRAI`, `#VALEUR!` (langue d'Office) |
| Résultats de formules | `=SOMME(B4:E4)` | résultat enregistré par Excel |

Les arrondis sont faits comme dans Excel (sur 15 chiffres significatifs, 5 arrondi au-dessus : `1,005` au
format `0,00` donne `1,01`). Une colonne trop étroite n'a pas d'effet : là où Excel afficherait `#####` ou
moins de décimales au format Standard, Word reçoit la valeur complète. Séparateurs décimal et de milliers, date courte et noms des mois suivent les
paramètres régionaux de Windows, comme dans Excel.

### Mise en forme importée

| Excel | Word |
|---|---|
| Police, taille, gras, italique, souligné, barré, exposant, indice, couleur (thème, nuances, palette) | Identiques, **par segment** (texte enrichi dans une cellule) |
| Remplissage (uni, motifs rendus par leur couleur moyenne, dégradés) | Trame de cellule |
| Bordures (style, épaisseur, couleur), bordures communes à deux cellules | Bordures Word équivalentes, harmonisées entre cellules voisines |
| Cellules fusionnées, « Centrer sur plusieurs colonnes » | Cellules fusionnées |
| Texte qui déborde sur les cellules vides voisines | Cellule fusionnée avec ces cellules (comme à l'écran dans Excel) |
| Alignements horizontal et vertical (Standard : nombres à droite, texte à gauche), retrait, renvoi à la ligne, orientation verticale | Identiques |
| Largeur des colonnes, hauteur des lignes | Reproduites (hauteur « au moins », pour ne jamais masquer de texte) |
| Styles de tableau Excel (« Mettre sous forme de tableau ») : en-tête, bandes, totaux | Reproduits cellule par cellule (styles personnalisés exacts, styles intégrés reconstitués) |
| Mise en forme conditionnelle : valeurs, textes, 10 premiers, moyenne, doublons, nuances de couleurs, formules simples (`MOD(LIGNE();2)=0`, `$C2>100`…) | Appliquée selon les valeurs du classeur |
| Titres d'impression, ligne d'en-tête d'un tableau Excel | Lignes d'en-tête répétées en haut de chaque page |
| Feuille de droite à gauche | Tableau de droite à gauche |

Le tableau inséré a une mise en forme entièrement explicite (polices, espacements nuls, marges de cellule
réduites, disposition fixe) : il a le même aspect quel que soit le style *Normal* du document.
Chaque ligne reste d'un seul tenant d'une page à l'autre, sauf si elle est trop haute.

Limites connues : les images, graphiques, formes, commentaires, barres de données et jeux d'icônes ne sont
pas importés (le compte rendu le signale) ; les liens hypertexte sont importés comme du texte (avec leur
mise en forme) ; les styles de tableau intégrés d'Excel sont reconstitués de façon approchée ; un tableau Word compte au plus 63 colonnes ; les formules
d'un classeur jamais recalculé (produit par un autre logiciel) n'ont pas de résultat enregistré : les cellules
restent vides et le compte rendu invite à ouvrir et enregistrer le classeur dans Excel.

### Où le tableau est-il inséré ?

- Curseur sur un **paragraphe vide** : le tableau y est placé.
- Curseur au **début d'un paragraphe** : le tableau est placé juste avant.
- Curseur **au milieu ou à la fin d'un paragraphe**, ou **texte sélectionné** : le tableau est placé après le
  paragraphe ; **le texte sélectionné n'est jamais remplacé**.
- Curseur **dans un tableau** : le nouveau tableau est placé après ce tableau.
- Deux tableaux ne sont jamais accolés (Word les fusionnerait) : un paragraphe vide les sépare.

L'insertion se fait en un bloc (`Range.InsertXML`, Word 2007 et suivants) ; en cas de refus, un document
`.docx` temporaire est inséré à la place. Un document protégé contre les modifications ou ouvert en mode
protégé est signalé sans rien modifier.

## Désinstallation

Fermez Word, puis double-cliquez sur **`uninstall.cmd`** (dossier `Installation`). L'enregistrement COM,
la déclaration auprès de Word, les fichiers de `%LOCALAPPDATA%\WordTableToExcel` et les préférences sont
supprimés. Aucun droit administrateur n'est nécessaire.

## Dépannage

| Problème | Solution |
|---|---|
| Le bouton n'apparaît pas | **Fichier › Options › Compléments**, liste *Gérer : Compléments COM* › **Atteindre…** : cochez « Tableaux Word vers Excel ». |
| Le complément a été désactivé par Word | **Fichier › Options › Compléments**, *Gérer : Éléments désactivés* › **Atteindre…** : réactivez-le, puis redémarrez Word. |
| « Comportement au chargement : non chargé. Une erreur d'exécution s'est produite » | Vérifiez que .NET Framework 4.x est installé (inclus dans Windows 8, 10, 11) ; relancez `install.cmd`, qui recopie et débloque les fichiers. Le journal indique l'étape en cause. |
| `install.cmd` signale un fichier manquant | Le `.zip` n'a pas été extrait : clic droit › *Extraire tout*, puis lancez `install.cmd` depuis le dossier extrait. |
| Le classeur ne peut pas être enregistré | Le fichier est probablement ouvert dans Excel : fermez-le et recommencez. |
| Un tableau n'est pas exporté ou apparaît simplifié | Le compte rendu final et le journal le signalent. |
| Import : « format Excel 97-2003 » ou « .xlsb » | Excel n'est pas installé pour convertir le fichier : ouvrez-le dans Excel ailleurs et enregistrez-le au format `.xlsx`. |
| Import : cellules de formules vides | Le classeur n'a jamais été calculé par Excel : ouvrez-le dans Excel, enregistrez-le, puis recommencez. |
| Import : tableau trop large | Décochez « Réduire les tableaux trop larges », passez la section en orientation *Paysage* ou indiquez une plage plus étroite. |
| Import : annuler | **Ctrl+Z** retire en une fois tous les tableaux et légendes insérés. |

Journal de diagnostic : `%LOCALAPPDATA%\WordTableToExcel\WordTableToExcel.log`.

## Compilation et tests

**Les utilisateurs n'ont rien à compiler** : le dossier `Installation/` contient les fichiers prêts à
l'emploi. Cette section concerne uniquement la maintenance du projet.

- `build/build-installation.sh` (Linux, WSL ou Git Bash) : exécute les tests et régénère les trois DLL du
  dossier `Installation/`. Prérequis : [SDK .NET 8](https://dotnet.microsoft.com/download) et MinGW-w64
  (`apt install gcc-mingw-w64-i686 gcc-mingw-w64-x86-64`).
- `build.cmd` (Windows) : exécute les tests et régénère `Installation\WordTableToExcel.dll` (le chargeur
  natif ne change pas).

```bash
dotnet test tests/WordTableToExcel.Tests      # 280+ tests, exécutables sous Windows, Linux ou macOS
dotnet build src/WordTableToExcel -c Release  # DLL .NET Framework 4.0, AnyCPU, signée par nom fort
build/build-shim.sh sortie/                   # chargeurs natifs 32 et 64 bits
```

Les tests couvrent l'analyse de la structure des tableaux (fichiers réels générés par
`tests/fixtures/make_fixtures.py`), les styles de tableau, la détection des légendes, les noms de
feuilles, la conversion des nombres, l'écriture du `.xlsx` et la lecture Word via un faux modèle
objet Word (`tests/WordTableToExcel.Tests/Fakes`). Ils s'exécutent sans Word.

Pour l'import : lecture du classeur de test `tests/fixtures/import_complexe.xlsx` (généré par
`tests/fixtures/make_import_fixture.py` puis recalculé par LibreOffice), moteur de formats de nombre
(comparé à LibreOffice sur un millier de combinaisons valeur × format ; les écarts restants sont des
différences connues entre LibreOffice et Excel, où le complément suit Excel), styles de tableau, mise en
forme conditionnelle et formules, aller-retour Word → Excel → Word, XML Word généré relu par l'analyseur
de tableaux, et choix du point d'insertion avec un faux Word. L'intégration continue vérifie en plus
l'import avec la DLL .NET Framework 4 dans Windows PowerShell (`tests/ci/framework-check.ps1`), c'est-à-dire
dans le même environnement d'exécution que Word.

L'intégration continue (`.github/workflows/build.yml`) compile le chargeur natif (Linux) et le complément
(Windows), exécute les tests, puis vérifie l'installation réelle **deux fois** : avec le dossier
`Installation/` du dépôt, tel que vous le téléchargez, et avec un dossier reconstruit à partir des sources.
Chaque fois : fichiers marqués « provenant d'Internet », `install.cmd` sans droits administrateur, clés de
registre, activation du complément et lecture du ruban par un client COM natif (comme Word) en 32 et
64 bits depuis un compte utilisateur standard, puis `uninstall.cmd`. Le fonctionnement dans Word
lui-même (clic sur le bouton, lecture d'un vrai document) se valide sur un poste équipé de Word ;
`tests/fixtures/medium_shading_merged.docx` peut servir de document d'essai.

## Architecture

```
Installation/         Fichiers prêts à l'emploi (DLL + install.cmd) : ce que l'utilisateur utilise
src/WordTableToExcel.Shim/
└── shim.c            Chargeur natif (C) : enregistré dans HKCU, démarre .NET 4 et crée le complément
src/WordTableToExcel/
├── AddIn/            Point d'entrée COM (IDTExtensibility2 + ruban), entrée du chargeur natif,
│                     barre d'outils Word 2000-2003, interfaces Office (sans PIA), Ribbon.xml
├── Export/           Déroulement de l'export (ExportService) et sélection A/B + noms de feuilles (ExportPlan)
├── Import/           Déroulement de l'import (ImportService), conversion des anciens formats par Excel,
│                     mesure du texte (GDI)
├── UI/               Boîtes de dialogue (export, import), fenêtre de progression, messages
├── Word/             Accès au document via le modèle objet (liaison tardive) :
│                     lecture des tableaux et légendes, insertion des tableaux importés et de leur légende
├── Core/             Code indépendant de Word et de Windows :
│   ├── Layout/       Analyse du XML des tableaux (grille, fusions, styles de tableau, bordures)
│   ├── Captions/     Reconnaissance multilingue des légendes, attribution, noms de feuilles
│   ├── Text/         Nettoyage du texte, reconnaissance des nombres
│   ├── Xlsx/         Écriture du classeur .xlsx (ZIP + SpreadsheetML, styles dédupliqués)
│   ├── ExcelImport/  Lecture du .xlsx (ZIP, classeur, styles, thème, feuilles), formats de nombre et dates,
│   │                 styles de tableau, mise en forme conditionnelle, conversion feuille → tableau,
│   │                 écriture du tableau WordprocessingML
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
| `ImportAddCaption` | DWORD | Import : ajouter une légende (1) ou non (0) |
| `ImportCaptionBelow` | DWORD | Import : légende sous le tableau (1) ou au-dessus (0) ; absente = selon le document |
| `ImportFitToPage` | DWORD | Import : réduire les tableaux trop larges (1) ou garder les largeurs Excel (0) |
| `ImportSkipHidden` | DWORD | Import : ignorer les lignes et colonnes masquées (1) ou non (0) |
| `ImportGridlines` | DWORD | Import : ajouter un quadrillage aux cellules sans bordure (1) ou non (0) |
| `ImportLastFolder` | Chaîne | Import : dernier dossier ouvert |
