using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Text;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Lit un tableau Word (objet COM Word.Table) et produit un <see cref="TableModel"/> :
    /// <list type="number">
    /// <item>structure exacte (grille, fusions, fonds, bordures, styles de tableau) depuis le XML du tableau ;</item>
    /// <item>texte et mise en forme des caractères depuis le modèle objet (valeurs effectives calculées par Word),
    /// ou, avec <see cref="ReadContentFromXml"/>, depuis ce même XML (quelques appels à Word par tableau au lieu de
    /// plusieurs dizaines par cellule), après vérification que son texte est identique à celui renvoyé par Word ;</item>
    /// <item>mode de repli COM complet si le XML n'est pas disponible (anciennes versions, document protégé).</item>
    /// </list>
    /// Toutes les opérations sont en lecture seule.
    /// </summary>
    public sealed class WordTableReader
    {
        private sealed class ComCell
        {
            public object Cell;
            public int RowIndex;
            public int ColumnIndex;
            public int Start;
            public int End;
        }

        private readonly dynamic _document;
        private readonly WordRunReader _runReader;
        private readonly Func<int, Rgb?> _themeLookup;
        private readonly Action<string> _log;
        private readonly Dictionary<bool, string> _themeFonts = new Dictionary<bool, string>();

        /// <summary>Lecture cellule par cellule : références à Word libérées toutes les N cellules.</summary>
        private const int ReleaseEveryCells = 500;

        /// <summary>
        /// Lit le texte et la mise en forme des caractères dans le XML du tableau plutôt que caractère par caractère
        /// auprès de Word : beaucoup plus rapide quand Word est piloté depuis un autre programme. Le texte obtenu est
        /// comparé à celui que renvoie Word ; s'il diffère (texte masqué, suppressions suivies…), le tableau est lu
        /// cellule par cellule comme d'habitude.
        /// </summary>
        public bool ReadContentFromXml { get; set; }

        /// <summary>
        /// Vrai si le dernier tableau a été lu cellule par cellule auprès de Word (des milliers d'objets Word créés,
        /// à libérer ensuite : voir <see cref="WordCom.ReleaseUnusedReferences"/>).
        /// </summary>
        public bool LastReadCellByCell { get; private set; }

        /// <summary>
        /// Chercher les suppressions suivies dans les tableaux lus cellule par cellule ; inutile, et évité, si le
        /// document n'a aucune modification suivie (voir <see cref="WordRevisions.DocumentHasRevisions"/>).
        /// </summary>
        public bool CheckRevisions { get; set; } = true;

        public WordTableReader(object document, Action<string> log)
        {
            if (document == null) throw new ArgumentNullException("document");
            _document = document;
            _log = log ?? (s => { });
            _themeLookup = CreateThemeLookup(document);
            _runReader = new WordRunReader(document, _themeLookup);
        }

        /// <param name="table">Objet Word.Table.</param>
        /// <param name="documentIndex">Rang du tableau dans le document (base 1).</param>
        /// <param name="progress">Progression facultative (appelée régulièrement avec le nombre de cellules lues).</param>
        public TableModel Read(object table, int documentIndex, Action<int, int> progress)
        {
            dynamic t = table;
            var model = new TableModel { DocumentIndex = documentIndex };

            TableLayout layout = TryReadXmlLayout(t, model);
            int objects = layout != null ? layout.ObjectCount : CountObjectsFromCom(t);
            bool xmlContent = ReadContentFromXml && layout != null && layout.HasContent && XmlTextMatchesWord(t, layout, documentIndex);
            LastReadCellByCell = !xmlContent;
            Dictionary<LayoutCell, ComCell> pairs = null;
            List<int[]> excluded = null;

            if (!xmlContent)
            {
                excluded = ExcludedIntervals(t, documentIndex);
                var comCells = EnumerateCells(t);
                if (layout != null)
                {
                    pairs = Pair(layout, comCells, model);
                    if (pairs == null)
                    {
                        model.Warnings.Add("structure XML incohérente avec Word, reconstruction à partir des largeurs de cellules");
                        layout = null;
                    }
                }
                if (layout == null)
                {
                    layout = BuildLayoutFromWidths(comCells);
                    pairs = layout.Cells.Where(c => c.SourceTag is ComCell).ToDictionary(c => c, c => (ComCell)c.SourceTag);
                }
            }

            model.RowCount = layout.RowCount;
            model.ColumnCount = layout.ColumnCount;
            model.ColumnWidthsPt = layout.ColumnWidthsPt;
            model.RowHeightsPt = layout.RowHeightsPt;
            model.RowHeightExact = layout.RowHeightExact;
            model.HeaderRowCount = layout.HeaderRowCount;

            int done = 0;
            foreach (var lc in layout.Cells)
            {
                var cell = new CellModel
                {
                    Row = lc.Row,
                    Column = lc.Column,
                    RowSpan = lc.RowSpan,
                    ColumnSpan = lc.ColumnSpan,
                    Borders = lc.Borders ?? new CellBorders(),
                    VerticalAlignment = lc.VerticalAlignment,
                    TextRotation = lc.TextRotation,
                    Fill = lc.Fill
                };

                ComCell source;
                bool truncated = false;
                if (xmlContent)
                {
                    if (lc.Content != null)
                    {
                        cell.Runs.AddRange(CellTextSanitizer.Clean(lc.Content.Runs, out truncated));
                        cell.HorizontalAlignment = MapAlignment(lc.Content.Alignment);
                    }
                }
                else if (pairs.TryGetValue(lc, out source))
                {
                    try
                    {
                        truncated = ReadCellContent(source, cell, layout.HasCellFormatting, excluded);
                    }
                    catch (Exception ex)
                    {
                        // Une cellule illisible ne doit pas faire échouer tout le tableau.
                        model.Warnings.Add(string.Format("cellule L{0}C{1} lue en texte brut ({2})", lc.Row + 1, lc.Column + 1, ex.Message));
                        _log("Cellule " + lc + " : " + ex);
                        truncated = ReadPlainText(source, cell);
                    }
                }
                if (truncated)
                {
                    model.Omissions.Add(string.Format(CultureInfo.CurrentCulture, "cellule L{0}C{1} : texte coupé à {2:N0} caractères (limite d'une cellule Excel)",
                        lc.Row + 1, lc.Column + 1, CellTextSanitizer.ExcelMaxCellLength));
                }

                if (!cell.Fill.HasValue) cell.Fill = lc.ParagraphShading;
                if (!cell.Fill.HasValue)
                {
                    // Excel ne sait pas surligner une partie de cellule : le surlignage devient le fond.
                    foreach (var run in cell.Runs)
                    {
                        if (run.Highlight.HasValue && run.Text.Trim().Length > 0)
                        {
                            cell.Fill = run.Highlight;
                            break;
                        }
                    }
                }

                model.Cells.Add(cell);
                done++;
                if (progress != null && (done % 20 == 0 || done == layout.Cells.Count)) progress(done, layout.Cells.Count);
                // Lecture cellule par cellule : chaque plage, police… lue garde un objet vivant dans Word tant
                // qu'elle n'est pas libérée ; on les libère au fil de l'eau pour que Word ne grossisse pas.
                if (!xmlContent && done % ReleaseEveryCells == 0) WordCom.ReleaseUnusedReferences();
            }

            if (objects > 0)
            {
                model.Omissions.Add(objects == 1
                    ? "1 image ou objet (graphique, forme, zone de texte…) non exporté : seul le texte des cellules est copié"
                    : objects.ToString(CultureInfo.CurrentCulture) + " images ou objets (graphiques, formes, zones de texte…) non exportés : seul le texte des cellules est copié");
            }
            return model;
        }

        /// <summary>Images et objets du tableau, quand son XML n'est pas disponible.</summary>
        private static int CountObjectsFromCom(dynamic table)
        {
            int count = 0;
            try
            {
                count += WordCom.AsInt(table.Range.InlineShapes.Count);
            }
            catch (Exception)
            {
                // Propriété indisponible : rien à signaler.
            }
            try
            {
                count += WordCom.AsInt(table.Range.ShapeRange.Count);
            }
            catch (Exception)
            {
                // Aucune forme flottante ancrée dans le tableau.
            }
            return count;
        }

        // ------------------------------------------------------------------ structure

        private List<ComCell> EnumerateCells(dynamic table)
        {
            bool hasNested = false;
            int nesting = 1;
            try
            {
                hasNested = WordCom.AsInt(table.Tables.Count) > 0;
                if (hasNested) nesting = WordCom.AsInt(table.NestingLevel);
            }
            catch (Exception)
            {
                hasNested = false;
            }

            var list = new List<ComCell>();
            foreach (dynamic cell in table.Range.Cells)
            {
                if (hasNested && WordCom.AsInt(cell.NestingLevel) != nesting) continue;
                dynamic range = cell.Range;
                list.Add(new ComCell
                {
                    Cell = cell,
                    RowIndex = WordCom.AsInt(cell.RowIndex),
                    ColumnIndex = WordCom.AsInt(cell.ColumnIndex),
                    Start = WordCom.AsInt(range.Start),
                    End = WordCom.AsInt(range.End)
                });
            }
            return list;
        }

        private TableLayout TryReadXmlLayout(dynamic table, TableModel model)
        {
            dynamic range = table.Range;
            string xml = TryGetXml(() => range.WordOpenXML);          // Word 2007+
            if (xml == null) xml = TryGetXml(() => range.XML);        // Word 2003+ (WordprocessingML 2003)
            if (xml == null) xml = TryGetXml(() => range.XML(false)); // même propriété, appelée avec son paramètre facultatif
            if (xml == null)
            {
                _log("Tableau " + model.DocumentIndex + " : XML indisponible, mode de repli.");
                return null;
            }
            try
            {
                var layout = WordXmlTableParser.Parse(xml, ThemeFont);
                if (layout == null) _log("Tableau " + model.DocumentIndex + " : aucun tableau dans le XML.");
                return layout;
            }
            catch (Exception ex)
            {
                _log("Tableau " + model.DocumentIndex + " : XML illisible (" + ex.Message + ").");
                return null;
            }
        }

        /// <summary>
        /// Le texte lu dans le XML est-il celui que Word renvoie pour le tableau ? (un seul appel à Word). En cas de
        /// doute, le tableau est lu cellule par cellule : jamais de valeur différente de ce qu'affiche Word.
        /// </summary>
        private bool XmlTextMatchesWord(dynamic table, TableLayout layout, int documentIndex)
        {
            string wordText;
            try
            {
                wordText = WordCom.AsString(table.Range.Text);
            }
            catch (Exception ex)
            {
                _log("Tableau " + documentIndex + " : texte indisponible (" + ex.Message + "), lecture cellule par cellule.");
                return false;
            }
            // Selon l'affichage, Word inclut ou non le texte supprimé en suivi des modifications et le texte masqué :
            // chaque variante lue dans le XML est acceptée.
            string word = CellTextSanitizer.Comparable(wordText);
            foreach (var variant in layout.XmlTextVariants)
            {
                if (string.Equals(word, CellTextSanitizer.Comparable(variant), StringComparison.Ordinal)) return true;
            }
            _log("Tableau " + documentIndex + " : texte du XML différent de celui de Word, lecture cellule par cellule.");
            return false;
        }

        /// <summary>
        /// Passages du tableau à ne pas exporter quand il est lu cellule par cellule : texte supprimé en suivi des
        /// modifications (accepté ou non) et résultat des renvois vers une note (champs NOTEREF).
        /// </summary>
        private List<int[]> ExcludedIntervals(dynamic table, int documentIndex)
        {
            var intervals = new List<int[]>();
            try
            {
                if (CheckRevisions) intervals.AddRange(WordRevisions.DeletedIntervals((object)table.Range));
            }
            catch (Exception ex)
            {
                _log("Tableau " + documentIndex + " : révisions illisibles (" + ex.Message + ").");
            }
            try
            {
                intervals.AddRange(WordRevisions.NoteReferenceIntervals((object)table.Range));
            }
            catch (Exception ex)
            {
                _log("Tableau " + documentIndex + " : champs illisibles (" + ex.Message + ").");
            }
            return WordRevisions.Normalize(intervals);
        }

        /// <summary>Police du thème du document (titres ou corps), si le XML du tableau ne contient pas le thème.</summary>
        private string ThemeFont(bool major)
        {
            string name;
            if (_themeFonts.TryGetValue(major, out name)) return name;
            try
            {
                dynamic scheme = _document.DocumentTheme.ThemeFontScheme;
                dynamic fonts = major ? scheme.MajorFont : scheme.MinorFont;
                name = WordCom.AsString(fonts.Item(1).Name); // msoThemeLatin
            }
            catch (Exception)
            {
                name = null;
            }
            _themeFonts[major] = name;
            return name;
        }

        private static string TryGetXml(Func<object> getter)
        {
            try
            {
                string xml = getter() as string;
                return !string.IsNullOrEmpty(xml) && xml.TrimStart().StartsWith("<", StringComparison.Ordinal) ? xml : null;
            }
            catch (Exception)
            {
                return null; // propriété absente de cette version de Word, ou document protégé
            }
        }

        /// <summary>
        /// Associe les cellules de la structure XML aux cellules COM. Word ne présente pas les
        /// continuations de fusion verticale dans sa collection Cells : on associe donc, ligne par
        /// ligne, les cellules visibles dans l'ordre.
        /// </summary>
        private Dictionary<LayoutCell, ComCell> Pair(TableLayout layout, List<ComCell> comCells, TableModel model)
        {
            var byRow = comCells.GroupBy(c => c.RowIndex).ToDictionary(g => g.Key, g => g.OrderBy(c => c.ColumnIndex).ToList());
            // Lignes de Word, y compris celles supprimées en suivi des modifications (absentes de la grille).
            int sourceRows = Math.Max(layout.RowCount, layout.SourceRowCount);
            if (byRow.Count > 0 && byRow.Keys.Max() > sourceRows) return null;

            var pairs = new Dictionary<LayoutCell, ComCell>();
            int mismatchedRows = 0;
            for (int r = 0; r < sourceRows; r++)
            {
                if (layout.DeletedRows.Contains(r)) continue; // ligne supprimée : non exportée
                var visible = layout.CellsStartingOnSourceRow(r);
                List<ComCell> row;
                if (!byRow.TryGetValue(r + 1, out row)) row = new List<ComCell>();

                if (row.Count == visible.Count)
                {
                    for (int i = 0; i < visible.Count; i++) pairs[visible[i]] = row[i];
                }
                else if (layout.SourceCellCounts != null && row.Count == layout.SourceCellCounts[r])
                {
                    foreach (var lc in visible)
                    {
                        if (lc.SourceIndex < row.Count) pairs[lc] = row[lc.SourceIndex];
                    }
                }
                else
                {
                    mismatchedRows++;
                    foreach (var lc in visible)
                    {
                        var match = row.FirstOrDefault(c => c.ColumnIndex == lc.SourceIndex + 1);
                        if (match != null) pairs[lc] = match;
                    }
                }
            }

            if (mismatchedRows > 0)
            {
                if (mismatchedRows * 2 > layout.RowCount) return null;
                model.Warnings.Add(mismatchedRows + " ligne(s) à la structure inhabituelle");
            }
            return pairs;
        }

        private TableLayout BuildLayoutFromWidths(List<ComCell> comCells)
        {
            var infos = new List<ComCellInfo>();
            foreach (var c in comCells)
            {
                double width = 0;
                try
                {
                    width = WordCom.AsDouble(((dynamic)c.Cell).Width);
                }
                catch (Exception)
                {
                    width = 0;
                }
                if (WordCom.IsUndefined(width)) width = 0;
                infos.Add(new ComCellInfo { RowIndex = c.RowIndex, ColumnIndex = c.ColumnIndex, WidthPt = width, Tag = c });
            }
            return WidthGridBuilder.Build(infos);
        }

        // ------------------------------------------------------------------ contenu

        /// <param name="excluded">Passages à ne pas exporter (suppressions suivies, renvois de notes).</param>
        /// <returns>Vrai si le texte a été coupé à la limite d'une cellule Excel.</returns>
        private bool ReadCellContent(ComCell source, CellModel cell, bool formattingFromXml, List<int[]> excluded)
        {
            bool truncated = false;
            // La plage d'une cellule se termine par la marque de fin de cellule (1 position) : on l'exclut.
            int contentEnd = source.End - 1;
            if (contentEnd > source.Start)
            {
                var runs = new List<TextRun>();
                foreach (var piece in WordRevisions.Subtract(source.Start, contentEnd, excluded))
                {
                    runs.AddRange(_runReader.Read(piece[0], piece[1]));
                }
                cell.Runs.AddRange(CellTextSanitizer.Clean(runs, out truncated));
            }

            dynamic c = source.Cell;
            cell.HorizontalAlignment = ReadHorizontalAlignment(c);

            if (!formattingFromXml) ReadCellFormattingFromCom(c, cell);
            return truncated;
        }

        /// <returns>Vrai si le texte a été coupé à la limite d'une cellule Excel.</returns>
        private bool ReadPlainText(ComCell source, CellModel cell)
        {
            cell.Runs.Clear();
            try
            {
                string text = WordCom.AsString(((dynamic)source.Cell).Range.Text);
                bool truncated;
                cell.Runs.AddRange(CellTextSanitizer.Clean(new[] { new TextRun(text, new RunFormat(), null) }, out truncated));
                return truncated;
            }
            catch (Exception ex)
            {
                _log("Texte illisible : " + ex.Message);
                return false;
            }
        }

        private static HorizontalAlignment ReadHorizontalAlignment(dynamic cell)
        {
            int alignment;
            try
            {
                dynamic range = cell.Range;
                alignment = WordCom.AsInt(range.ParagraphFormat.Alignment);
                if (WordCom.IsUndefined(alignment)) alignment = WordCom.AsInt(range.Paragraphs.Item(1).Alignment);
            }
            catch (Exception)
            {
                return HorizontalAlignment.General;
            }
            return MapAlignment(alignment);
        }

        /// <summary>WdParagraphAlignment → alignement Excel.</summary>
        public static HorizontalAlignment MapAlignment(int wdAlignment)
        {
            switch (wdAlignment)
            {
                case 0: return HorizontalAlignment.Left;
                case 1: return HorizontalAlignment.Center;
                case 2: return HorizontalAlignment.Right;
                case 3:
                case 5:
                case 7:
                case 8:
                case 9: return HorizontalAlignment.Justify;
                case 4: return HorizontalAlignment.Distributed;
                default: return HorizontalAlignment.General;
            }
        }

        /// <summary>Mode de repli : fond, bordures, alignement vertical et orientation lus via COM.</summary>
        private void ReadCellFormattingFromCom(dynamic c, CellModel cell)
        {
            try
            {
                dynamic shading = c.Shading;
                Rgb? fill = WordColor.Decode(WordCom.AsInt(shading.BackgroundPatternColor), _themeLookup);
                if (!fill.HasValue && WordCom.AsInt(shading.Texture) == 1000) // wdTextureSolid
                {
                    fill = WordColor.Decode(WordCom.AsInt(shading.ForegroundPatternColor), _themeLookup);
                }
                cell.Fill = fill;
            }
            catch (Exception ex)
            {
                _log("Fond de cellule illisible : " + ex.Message);
            }

            try
            {
                int v = WordCom.AsInt(c.VerticalAlignment);
                cell.VerticalAlignment = v == 1 ? VerticalAlignment.Center : v == 3 ? VerticalAlignment.Bottom : VerticalAlignment.Top;
            }
            catch (Exception)
            {
                cell.VerticalAlignment = VerticalAlignment.Top;
            }

            try
            {
                int orientation = WordCom.AsInt(c.Range.Orientation);
                cell.TextRotation = orientation == 2 ? 90 : orientation == 3 ? 180 : 0;
            }
            catch (Exception)
            {
                cell.TextRotation = 0;
            }

            try
            {
                dynamic borders = c.Borders;
                cell.Borders = new CellBorders
                {
                    Top = ReadComBorder(borders.Item(-1)),
                    Left = ReadComBorder(borders.Item(-2)),
                    Bottom = ReadComBorder(borders.Item(-3)),
                    Right = ReadComBorder(borders.Item(-4))
                };
            }
            catch (Exception ex)
            {
                _log("Bordures illisibles : " + ex.Message);
            }
        }

        private BorderLine ReadComBorder(dynamic border)
        {
            int style = WordCom.AsInt(border.LineStyle);
            if (style == 0 || WordCom.IsUndefined(style)) return BorderLine.None;
            int width = WordCom.AsInt(border.LineWidth); // WdLineWidth = huitièmes de point
            if (WordCom.IsUndefined(width) || width <= 0) width = 4;
            Rgb? color = WordColor.Decode(WordCom.AsInt(border.Color), _themeLookup);
            return new BorderLine(OoxmlFormat.MapBorderStyle(ComLineStyleName(style), width), color);
        }

        private static string ComLineStyleName(int wdLineStyle)
        {
            switch (wdLineStyle)
            {
                case 2: return "dotted";
                case 3:
                case 4: return "dashed";
                case 5: return "dotdash";
                case 6: return "dotdotdash";
                case 7: return "double";
                case 8: return "triple";
                case 18: return "dashdotstroked";
                case 19:
                case 20:
                case 21:
                case 22: return "threedemboss";
                default:
                    return wdLineStyle >= 9 && wdLineStyle <= 17 ? "thinthicksmallgap" : "single";
            }
        }

        // ------------------------------------------------------------------ thème

        /// <summary>Couleurs du thème du document (Word 2007+), mises en cache.</summary>
        private Func<int, Rgb?> CreateThemeLookup(object document)
        {
            var cache = new Dictionary<int, Rgb?>();
            dynamic doc = document;
            return index =>
            {
                Rgb? color;
                if (cache.TryGetValue(index, out color)) return color;
                try
                {
                    int rgb = WordCom.AsInt(doc.DocumentTheme.ThemeColorScheme.Colors(index).RGB);
                    color = WordColor.IsPlainRgb(rgb) ? Rgb.FromBgr(rgb) : (Rgb?)null;
                }
                catch (Exception)
                {
                    color = null;
                }
                cache[index] = color;
                return color;
            };
        }
    }
}
