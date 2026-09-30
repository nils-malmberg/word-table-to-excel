# Tableaux Word ↔ Excel

Échange des tableaux entre Microsoft Word et Excel (Windows), **une feuille par tableau**, en **conservant la
mise en forme** :

- **Word → Excel** : exporte les tableaux d'un document Word vers un classeur Excel. Le document Word n'est
  jamais modifié.
- **Excel → Word** : insère, là où se trouve le curseur dans Word, un tableau par feuille choisie d'un classeur,
  avec les **valeurs exactement telles qu'Excel les affiche** et, en option, une légende numérotée.

L'outil est une **application** : un seul fichier, `TableauxWordExcel.exe`, sans installation ni droits
administrateur. Elle pilote Word de l'extérieur, donc Word ne peut pas la bloquer comme un complément.

> Documentation technique (implémentation, arborescence, compilation, tests) : [DEVELOPPEMENT.md](DEVELOPPEMENT.md).

---

## Démarrer

1. Téléchargez [`Application/TableauxWordExcel.exe`](Application/TableauxWordExcel.exe) (bouton
   **Download raw file**, ou la dernière version dans *Releases*) et rangez-le où vous voulez (par exemple
   *Documents*).
2. Double-cliquez dessus. Au premier lancement, si Windows affiche « Windows a protégé votre ordinateur » :
   **Informations complémentaires › Exécuter quand même**.
   Si l'antivirus ou une règle de l'entreprise bloque le programme, ne cherchez pas à contourner le blocage :
   demandez au service informatique.

Configuration : Windows 7 SP1 à 11, .NET Framework 4 (inclus depuis Windows 8), Word 2007 ou plus récent
(32 ou 64 bits). Excel n'est pas nécessaire, sauf pour importer des `.xls`, `.xlsb`, `.ods` ou `.csv`.

La fenêtre est petite ; « Toujours visible » la garde au-dessus de Word. Elle se met à jour chaque fois que
vous y revenez (et avec **F5**).

## Exporter les tableaux d'un document vers Excel

1. Dans la liste, sélectionnez un document **ouvert dans Word**, ou ajoutez un fichier avec **Ajouter un
   fichier Word…** (ou en le déposant sur la fenêtre). Un fichier non ouvert est lu par un Word invisible, en
   lecture seule, sans macros ; vos documents ouverts ne sont pas touchés.
2. Cliquez sur **Exporter vers Excel…**. Sur un gros document, une barre de progression montre la recherche
   des légendes (annulable).
3. Dans la fenêtre d'export :
   - **Position des légendes** : « Automatique — détectée : au-dessus des tableaux » (ou au-dessous). Si la
     position détectée est fausse, choisissez la bonne : les légendes de la liste sont réattribuées aussitôt ;
   - **A — Uniquement les tableaux qui ont une légende**, ou **B — Tous les tableaux** ;
   - la liste montre chaque tableau avec ses **pages de début et de fin**, sa **feuille Excel** et sa
     **légende** (↑ au-dessus, ↓ au-dessous). **Décochez** ceux à ne pas exporter (la case **Tout** les coche
     ou décoche tous) ;
   - options : légende complète en haut de chaque feuille (A1), **feuille « Sommaire »** en tête du classeur
     (liste des tableaux avec un lien vers chaque feuille), conversion des nombres (« 1 234,50 », « 12,5 % »,
     « 45 € » deviennent de vrais nombres).
4. **Exporter…**, puis choisissez où enregistrer le classeur.
5. Le compte rendu indique **en premier, avec une icône d'avertissement, tout ce qui manque** dans le
   classeur : tableau non exporté, images et objets non copiés, texte coupé (limite de 32 767 caractères
   d'une cellule Excel). Il propose ensuite d'ouvrir le classeur.

Les choix sont mémorisés d'une fois sur l'autre.

**Ce qui est repris** : polices, tailles, couleurs, gras, italique, souligné, barré, exposant, indice (par
morceau de texte dans une même cellule) ; fonds, bordures et styles de tableau ; cellules fusionnées ;
alignements et orientation ; largeurs de colonnes et hauteurs de lignes ; sauts de ligne. Les lignes d'en-tête
répétées sur chaque page dans Word restent **figées** en haut de la feuille Excel.

**Ce qui n'est pas exporté, volontairement** : texte masqué ; **texte, lignes et tableaux supprimés en suivi
des modifications** (même si la suppression n'est pas encore acceptée ; une légende supprimée n'est plus une
légende) ; **renvois vers les notes de bas de page** (ils seraient pris pour des chiffres). Le texte inséré en
suivi des modifications est exporté.

**Légendes reconnues** : paragraphe juste au-dessus ou au-dessous du tableau contenant un champ `SEQ Tableau`
(ou *Table*, *Tabla*… dans toutes les langues) ou commençant par « Tableau 3 : », « Table 1. », « Tab. 5 »…
Les légendes de figures ne sont jamais prises pour des légendes de tableaux. Le nom de la feuille reprend la
légende (31 caractères au plus) ; sans légende : `Tableau_N` (N = rang du tableau dans le document).

**Limites** : images, graphiques, formes et zones de texte des cellules ne sont pas copiés (le compte rendu
les compte) ; un tableau imbriqué devient du texte dans sa cellule ; seuls les tableaux du corps du document
sont exportés (pas ceux des en-têtes et pieds de page) ; un surlignage partiel colore toute la cellule.

## Importer des tableaux Excel dans Word

1. Dans Word, cliquez à l'endroit où insérer le tableau. L'encadré de l'application indique où il sera placé
   (document, page, « juste après le paragraphe … »).
2. Cliquez sur **Importer depuis Excel…** (ou déposez le classeur sur la fenêtre). Le classeur peut rester
   ouvert dans Excel : il est seulement lu.
3. Cochez les feuilles à importer (un tableau par feuille ; case **Tout**), vérifiez la **plage** détectée et
   le **texte de la légende**. La feuille « Sommaire » d'un classeur exporté n'est pas cochée par défaut.
4. Options : légende numérotée « Tableau N » au-dessus ou au-dessous (sans texte, « [Titre du tableau] » est
   surligné en jaune, à compléter) ; réduction des tableaux trop larges pour la page ; lignes et colonnes
   masquées ignorées ; quadrillage gris.
5. **Importer**. Dans Word, **Ctrl+Z** annule tout l'import d'un coup.

Les valeurs sont celles qu'Excel **affiche** (nombres, monnaies, pourcentages, dates, durées, formats
personnalisés et régionaux, résultats de formules), avec leur mise en forme : polices, couleurs, fonds,
bordures, fusions, styles de tableau Excel, mise en forme conditionnelle. Le texte sélectionné dans Word n'est
jamais remplacé, et deux tableaux ne sont jamais accolés.

**Limites** : images, graphiques, formes, commentaires, barres de données et jeux d'icônes ne sont pas importés
(le compte rendu le signale) ; 63 colonnes au plus par tableau Word ; les formules d'un classeur jamais calculé
par Excel restent vides (ouvrez-le et enregistrez-le dans Excel).

## Dépannage

| Problème | Solution |
|---|---|
| « Windows a protégé votre ordinateur » | Avertissement normal pour un programme téléchargé non signé : **Informations complémentaires › Exécuter quand même**. |
| Programme bloqué par l'antivirus ou une règle de l'entreprise | Ne pas contourner : demander au service informatique d'autoriser `TableauxWordExcel.exe`. |
| « Word est occupé » | Une boîte de dialogue est ouverte dans Word (enregistrement, impression…) : fermez-la et recommencez. |
| Le document n'apparaît pas dans la liste | Revenez dans la fenêtre ou appuyez sur **F5**. Ne lancez ni Word ni l'application « en tant qu'administrateur ». |
| Import impossible (mode protégé, document final ou protégé) | L'encadré l'indique : **Activer la modification**, **Modifier quand même** ou retirez la protection dans Word. |
| Le classeur ne peut pas être enregistré | Il est sans doute ouvert dans Excel : fermez-le et recommencez. |
| La légende d'un tableau est attribuée au tableau voisin | Changez la **Position des légendes** dans la fenêtre d'export. |
| Import d'un `.xls`, `.xlsb`, `.ods` ou `.csv` refusé | Excel n'est pas installé pour le convertir : enregistrez-le au format `.xlsx` sur un autre poste. |
| Cellules de formules vides à l'import | Ouvrez le classeur dans Excel, enregistrez-le, puis recommencez. |

L'application n'écrit que ses préférences (`HKCU\Software\WordTableToExcel`) et son journal de diagnostic
(`%LOCALAPPDATA%\WordTableToExcel\WordTableToExcel.log`). Pour la supprimer : effacez le fichier `.exe`, ce
dossier et cette clé.

## Complément Word (gelé)

Le dossier [`Installation/`](Installation) contient aussi une version **complément Word** (boutons dans le
ruban), pour les postes où Word accepte les compléments : fermer Word, double-cliquer sur `install.cmd`
(aucun droit administrateur), `uninstall.cmd` pour le retirer. Mode d'emploi :
[`Installation/LISEZMOI.txt`](Installation/LISEZMOI.txt).

Ce complément est **gelé** : il reçoit les améliorations communes (export, import, légendes…), mais plus de
nouvelles fonctions qui lui soient propres. Sur les postes où Word exige des compléments signés, utilisez
l'application.
