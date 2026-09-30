using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Text;

namespace WordTableToExcel.Core.Xlsx
{
    public sealed class XlsxExportOptions
    {
        /// <summary>Écrit la légende complète en A1 (le tableau commence alors en ligne 3).</summary>
        public bool IncludeCaptionRow = true;
        /// <summary>Convertit les textes numériques sans ambiguïté en nombres Excel.</summary>
        public bool ConvertNumbers;
        /// <summary>Culture utilisée pour trancher les séparateurs ambigus (« 1,234 »).</summary>
        public CultureInfo NumberCulture = CultureInfo.CurrentCulture;
        /// <summary>Titre du classeur (propriétés du document).</summary>
        public string Title;
        /// <summary>
        /// Ajoute en tête du classeur une feuille « Sommaire » : un tableau par ligne (n°, légende complète, pages dans le
        /// document Word, lien vers sa feuille).
        /// </summary>
        public bool IncludeSummary;
    }

    /// <summary>
    /// Génère un classeur Excel (.xlsx, Office Open XML) : une feuille par tableau,
    /// avec texte enrichi, polices, couleurs, fonds, bordures, fusions, alignements et largeurs.
    /// Ne nécessite pas qu'Excel soit installé.
    /// </summary>
    public sealed class XlsxWorkbookWriter
    {
        private enum CellKind
        {
            Empty,
            SharedString,
            Number
        }

        private struct SheetCell
        {
            public int Style;
            public CellKind Kind;
            public int StringIndex;
            public double Number;
        }

        private sealed class Sheet
        {
            public string Name;
            public byte[] Xml;
            // Pour la feuille « Sommaire ».
            public int DocumentIndex;
            public string Caption;
            public int StartPage;
            public int EndPage;
        }

        /// <summary>Nom de la feuille de sommaire (suivi de « (2) »… si un tableau porte déjà ce nom).</summary>
        public const string SummarySheetName = "Sommaire";

        /// <summary>Feuille « Sommaire » (ou « Sommaire (2) »…) : à l'import, elle n'est pas proposée comme un tableau.</summary>
        public static bool IsSummarySheetName(string sheetName)
        {
            return sheetName != null && System.Text.RegularExpressions.Regex.IsMatch(sheetName.Trim(),
                "^" + SummarySheetName + @"( \(\d+\))?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private readonly XlsxExportOptions _options;
        private readonly XlsxStyleRegistry _styles = new XlsxStyleRegistry();
        private readonly List<object> _strings = new List<object>(); // string ou List<TextRun>
        private readonly Dictionary<string, int> _stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<Sheet> _sheets = new List<Sheet>();
        private Sheet _summary;
        private int _stringReferences;

        public XlsxWorkbookWriter(XlsxExportOptions options)
        {
            _options = options ?? new XlsxExportOptions();
        }

        public int SheetCount
        {
            get { return _sheets.Count; }
        }

        /// <summary>Ajoute une feuille pour le tableau (le nom de feuille doit déjà être valide et unique).</summary>
        public void AddTable(TableModel table)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (string.IsNullOrEmpty(table.SheetName)) throw new ArgumentException("Nom de feuille manquant.", "table");
            if (_sheets.Any(s => string.Equals(s.Name, table.SheetName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException("Nom de feuille en double : " + table.SheetName, "table");
            }

            var cells = new SortedDictionary<int, SortedDictionary<int, SheetCell>>();
            var merges = new List<string>();
            var rowHeights = new Dictionary<int, double>();
            int rowOffset = 0;

            if (_options.IncludeCaptionRow && table.HasCaption)
            {
                var captionFont = new RunFormat { Bold = true, Size = 12 };
                int style = _styles.GetXfId(_styles.GetFontId(captionFont), 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false, 0);
                Put(cells, 0, 0, new SheetCell { Style = style, Kind = CellKind.SharedString, StringIndex = AddString(CellTextSanitizer.CleanPlain(table.Caption)) });
                rowOffset = 2;
            }

            foreach (var cell in table.Cells)
            {
                WriteCell(cells, merges, cell, rowOffset);
            }

            ComputeRowHeights(table, rowOffset, rowHeights);

            int lastRow = Math.Max(cells.Count == 0 ? 0 : cells.Keys.Max(), rowHeights.Count == 0 ? 0 : rowHeights.Keys.Max());
            int lastCol = Math.Max(0, table.ColumnCount - 1);
            foreach (var row in cells.Values)
            {
                if (row.Count > 0) lastCol = Math.Max(lastCol, row.Keys.Max());
            }

            // Lignes d'en-tête répétées dans Word : figées en haut de la feuille (avec la légende éventuelle au-dessus).
            int frozenRows = table.HeaderRowCount > 0 && table.HeaderRowCount < table.RowCount ? rowOffset + table.HeaderRowCount : 0;
            // Feuille active à l'ouverture : la première, sauf si le sommaire la précède.
            bool selected = _sheets.Count == 0 && !_options.IncludeSummary;
            _sheets.Add(new Sheet
            {
                Name = table.SheetName,
                DocumentIndex = table.DocumentIndex,
                Caption = table.Caption,
                StartPage = table.StartPage,
                EndPage = table.EndPage,
                Xml = BuildSheetXml(table, cells, merges, rowHeights, lastRow, lastCol, selected, frozenRows)
            });
        }

        public void Save(string path)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Save(fs);
            }
        }

        public void Save(Stream output)
        {
            if (_sheets.Count == 0) throw new InvalidOperationException("Le classeur ne contient aucune feuille.");
            // Sommaire construit d'abord : il ajoute ses textes et ses styles à ceux du classeur.
            var sheets = new List<Sheet>();
            if (_options.IncludeSummary)
            {
                if (_summary == null) _summary = BuildSummary();
                sheets.Add(_summary);
            }
            sheets.AddRange(_sheets);
            using (var zip = new ZipWriter(output))
            {
                zip.AddEntry("[Content_Types].xml", ContentTypesXml(sheets.Count));
                zip.AddEntry("_rels/.rels", RootRelsXml());
                zip.AddEntry("docProps/core.xml", CoreXml());
                zip.AddEntry("docProps/app.xml", AppXml());
                zip.AddEntry("xl/workbook.xml", WorkbookXml(sheets));
                zip.AddEntry("xl/_rels/workbook.xml.rels", WorkbookRelsXml(sheets.Count));
                zip.AddEntry("xl/styles.xml", Xml(w => _styles.Write(w)));
                zip.AddEntry("xl/sharedStrings.xml", Xml(WriteSharedStrings));
                for (int i = 0; i < sheets.Count; i++)
                {
                    zip.AddEntry("xl/worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml", sheets[i].Xml);
                }
                zip.Finish();
            }
        }

        // ------------------------------------------------------------------ sommaire

        private static readonly Rgb SummaryHeaderFill = new Rgb(0xDD, 0xEB, 0xF7);
        private static readonly Rgb SummaryLinkColor = new Rgb(0x05, 0x63, 0xC1);
        private static readonly Rgb SummaryGray = new Rgb(0x7F, 0x7F, 0x7F);
        private static readonly Rgb SummaryRule = new Rgb(0xBF, 0xBF, 0xBF);

        /// <summary>
        /// Feuille « Sommaire » : titre, puis un tableau par ligne — n° dans le document, légende complète, pages,
        /// et nom de la feuille sous forme de lien qui y mène. Ligne d'en-tête figée.
        /// </summary>
        private Sheet BuildSummary()
        {
            string name = SummarySheetName;
            for (int n = 2; _sheets.Any(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)); n++)
            {
                name = SummarySheetName + " (" + n.ToString(CultureInfo.InvariantCulture) + ")";
            }

            var rule = new BorderLine(BorderStyle.Thin, SummaryRule);
            int lines = _styles.GetBorderId(new CellBorders { Bottom = rule });
            int title = _styles.GetXfId(_styles.GetFontId(new RunFormat { Bold = true, Size = 13 }), 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false, 0);
            int header = _styles.GetXfId(_styles.GetFontId(new RunFormat { Bold = true }), _styles.GetFillId(SummaryHeaderFill),
                _styles.GetBorderId(new CellBorders { Bottom = new BorderLine(BorderStyle.Thin, SummaryGray) }), 0, HorizontalAlignment.General, VerticalAlignment.Center, false, 0);
            int number = _styles.GetXfId(_styles.GetFontId(new RunFormat()), 0, lines, 0, HorizontalAlignment.Center, VerticalAlignment.Top, false, 0);
            int text = _styles.GetXfId(_styles.GetFontId(new RunFormat()), 0, lines, 0, HorizontalAlignment.General, VerticalAlignment.Top, true, 0);
            int none = _styles.GetXfId(_styles.GetFontId(new RunFormat { Italic = true, Color = SummaryGray }), 0, lines, 0, HorizontalAlignment.General, VerticalAlignment.Top, true, 0);
            int pages = _styles.GetXfId(_styles.GetFontId(new RunFormat()), 0, lines, 0, HorizontalAlignment.Center, VerticalAlignment.Top, false, 0);
            int link = _styles.GetXfId(_styles.GetFontId(new RunFormat { Color = SummaryLinkColor, Underline = UnderlineKind.Single }), 0, lines, 0,
                HorizontalAlignment.General, VerticalAlignment.Top, false, 0);

            const int headerRow = 2; // ligne 3 (base 0)
            var rows = new List<KeyValuePair<int, SheetCell[]>>();
            string heading = string.IsNullOrEmpty(_options.Title) ? "Tableaux exportés" : "Tableaux exportés de « " + CellTextSanitizer.CleanPlain(_options.Title) + " »";
            rows.Add(new KeyValuePair<int, SheetCell[]>(0, new[] { Text(heading, title) }));
            rows.Add(new KeyValuePair<int, SheetCell[]>(headerRow, new[] { Text("N°", header), Text("Légende", header), Text("Pages", header), Text("Feuille", header) }));
            var links = new List<KeyValuePair<string, Sheet>>();
            for (int i = 0; i < _sheets.Count; i++)
            {
                var sheet = _sheets[i];
                int r = headerRow + 1 + i;
                bool captioned = !string.IsNullOrEmpty(sheet.Caption);
                rows.Add(new KeyValuePair<int, SheetCell[]>(r, new[]
                {
                    sheet.DocumentIndex > 0 ? new SheetCell { Kind = CellKind.Number, Number = sheet.DocumentIndex, Style = number } : new SheetCell { Style = number },
                    captioned ? Text(CellTextSanitizer.CleanPlain(sheet.Caption), text) : Text("(sans légende)", none),
                    Text(PageText(sheet.StartPage, sheet.EndPage), pages),
                    Text(sheet.Name, link)
                }));
                links.Add(new KeyValuePair<string, Sheet>(ExcelUnits.CellReference(r, 3), sheet));
            }

            int lastRow = headerRow + _sheets.Count;
            var xml = Xml(w =>
            {
                w.WriteStartElement("worksheet", XlsxNames.Main);
                w.WriteAttributeString("xmlns", "r", null, XlsxNames.Relationships);
                w.WriteStartElement("dimension");
                w.WriteAttributeString("ref", "A1:" + ExcelUnits.CellReference(lastRow, 3));
                w.WriteEndElement();
                WriteSheetViews(w, true, headerRow + 1);
                w.WriteStartElement("sheetFormatPr");
                w.WriteAttributeString("defaultRowHeight", "15");
                w.WriteEndElement();
                w.WriteStartElement("cols");
                double[] widths = { 7, 80, 10, 34 };
                for (int c = 0; c < widths.Length; c++)
                {
                    w.WriteStartElement("col");
                    w.WriteAttributeString("min", (c + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("max", (c + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("width", widths[c].ToString("0.##", CultureInfo.InvariantCulture));
                    w.WriteAttributeString("customWidth", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();

                w.WriteStartElement("sheetData");
                foreach (var row in rows)
                {
                    w.WriteStartElement("row");
                    w.WriteAttributeString("r", (row.Key + 1).ToString(CultureInfo.InvariantCulture));
                    if (row.Key == 0) w.WriteAttributeString("ht", "21");
                    if (row.Key == 0) w.WriteAttributeString("customHeight", "1");
                    for (int c = 0; c < row.Value.Length; c++) WriteCell(w, row.Key, c, row.Value[c]);
                    w.WriteEndElement();
                }
                w.WriteEndElement();

                // Liens internes vers chaque feuille (aucune relation nécessaire : « location » désigne une cellule du classeur).
                w.WriteStartElement("hyperlinks");
                foreach (var entry in links)
                {
                    w.WriteStartElement("hyperlink");
                    w.WriteAttributeString("ref", entry.Key);
                    w.WriteAttributeString("location", "'" + XmlSafe(entry.Value.Name).Replace("'", "''") + "'!A1");
                    w.WriteAttributeString("display", XmlSafe(entry.Value.Name));
                    w.WriteEndElement();
                }
                w.WriteEndElement();

                WritePageMargins(w);
                w.WriteEndElement(); // worksheet
            });
            return new Sheet { Name = name, Xml = xml };
        }

        private SheetCell Text(string text, int style)
        {
            if (string.IsNullOrEmpty(text)) return new SheetCell { Kind = CellKind.Empty, Style = style };
            return new SheetCell { Kind = CellKind.SharedString, StringIndex = AddString(text), Style = style };
        }

        /// <summary>« 3 », « 3-4 », ou vide si les pages sont inconnues.</summary>
        internal static string PageText(int start, int end)
        {
            if (start <= 0) return string.Empty;
            if (end <= start) return start.ToString(CultureInfo.InvariantCulture);
            return start.ToString(CultureInfo.InvariantCulture) + "-" + end.ToString(CultureInfo.InvariantCulture);
        }

        // ------------------------------------------------------------------ cellules

        private void WriteCell(SortedDictionary<int, SortedDictionary<int, SheetCell>> cells, List<string> merges, CellModel cell, int rowOffset)
        {
            var runs = CellTextSanitizer.Clean(cell.Runs);
            var primary = FirstVisibleFormat(runs) ?? new RunFormat();
            int fontId = _styles.GetFontId(primary);
            int fillId = _styles.GetFillId(cell.Fill);
            int rotation = cell.TextRotation;

            var value = new SheetCell { Kind = CellKind.Empty };
            int numFmtId = 0;
            string text = string.Concat(runs.Select(r => r.Text).ToArray());

            if (text.Trim().Length > 0)
            {
                bool uniform = runs.All(r => r.Format.Equals(runs[0].Format));
                double number;
                string format;
                if (_options.ConvertNumbers && uniform && text.IndexOf('\n') < 0
                    && NumberParser.TryParse(text, _options.NumberCulture, out number, out format))
                {
                    value.Kind = CellKind.Number;
                    value.Number = number;
                    numFmtId = _styles.GetNumFmtId(format);
                }
                else if (uniform)
                {
                    value.Kind = CellKind.SharedString;
                    value.StringIndex = AddString(text);
                }
                else
                {
                    value.Kind = CellKind.SharedString;
                    value.StringIndex = AddRichString(runs);
                }
            }

            for (int r = cell.Row; r < cell.Row + cell.RowSpan; r++)
            {
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan; c++)
                {
                    // Dans une zone fusionnée, chaque case porte la partie de bordure qui lui revient.
                    var borders = new CellBorders
                    {
                        Top = r == cell.Row ? cell.Borders.Top : BorderLine.None,
                        Bottom = r == cell.Row + cell.RowSpan - 1 ? cell.Borders.Bottom : BorderLine.None,
                        Left = c == cell.Column ? cell.Borders.Left : BorderLine.None,
                        Right = c == cell.Column + cell.ColumnSpan - 1 ? cell.Borders.Right : BorderLine.None
                    };
                    int style = _styles.GetXfId(fontId, fillId, _styles.GetBorderId(borders), numFmtId,
                        cell.HorizontalAlignment, cell.VerticalAlignment, true, rotation);
                    var sheetCell = r == cell.Row && c == cell.Column ? value : new SheetCell { Kind = CellKind.Empty };
                    sheetCell.Style = style;
                    Put(cells, r + rowOffset, c, sheetCell);
                }
            }

            if (cell.IsMerged)
            {
                merges.Add(ExcelUnits.CellReference(cell.Row + rowOffset, cell.Column) + ":"
                         + ExcelUnits.CellReference(cell.Row + cell.RowSpan - 1 + rowOffset, cell.Column + cell.ColumnSpan - 1));
            }
        }

        private static RunFormat FirstVisibleFormat(List<TextRun> runs)
        {
            foreach (var run in runs)
            {
                if (run.Text.Trim().Length > 0) return run.Format;
            }
            return runs.Count > 0 ? runs[0].Format : null;
        }

        private static void Put(SortedDictionary<int, SortedDictionary<int, SheetCell>> cells, int row, int col, SheetCell cell)
        {
            SortedDictionary<int, SheetCell> line;
            if (!cells.TryGetValue(row, out line))
            {
                line = new SortedDictionary<int, SheetCell>();
                cells[row] = line;
            }
            line[col] = cell;
        }

        private int AddString(string text)
        {
            _stringReferences++;
            string key = "P" + text;
            int index;
            if (_stringIndex.TryGetValue(key, out index)) return index;
            index = _strings.Count;
            _strings.Add(text);
            _stringIndex[key] = index;
            return index;
        }

        private int AddRichString(List<TextRun> runs)
        {
            _stringReferences++;
            var sb = new StringBuilder("R");
            foreach (var run in runs)
            {
                sb.Append('\u0001').Append(XlsxStyleRegistry.NormalizeFont(run.Format)).Append('\u0002').Append(run.Text);
            }
            string key = sb.ToString();
            int index;
            if (_stringIndex.TryGetValue(key, out index)) return index;
            index = _strings.Count;
            _strings.Add(runs.Select(r => new TextRun(r.Text, r.Format, null)).ToList());
            _stringIndex[key] = index;
            return index;
        }

        // ------------------------------------------------------------------ hauteurs de ligne

        /// <summary>
        /// Excel ajuste lui-même la hauteur des lignes renvoyant à la ligne, sauf pour les cellules
        /// fusionnées : pour celles-ci (et pour les hauteurs imposées par Word), on calcule une hauteur.
        /// </summary>
        private static void ComputeRowHeights(TableModel table, int rowOffset, Dictionary<int, double> heights)
        {
            for (int r = 0; r < table.RowCount; r++)
            {
                double word = table.RowHeightsPt != null && r < table.RowHeightsPt.Length ? table.RowHeightsPt[r] : 0;
                bool exact = table.RowHeightExact != null && r < table.RowHeightExact.Length && table.RowHeightExact[r];

                if (exact && word > 0)
                {
                    heights[r + rowOffset] = Math.Min(word, ExcelUnits.MaxRowHeight);
                    continue;
                }

                double estimate = 0;
                bool hasMergedCell = false;
                foreach (var cell in table.Cells)
                {
                    if (cell.Row != r || cell.RowSpan != 1) continue;
                    if (cell.ColumnSpan > 1) hasMergedCell = true;
                    estimate = Math.Max(estimate, EstimateHeight(table, cell));
                }

                if (hasMergedCell && estimate > ExcelUnits.DefaultRowHeight + 0.5 || word > ExcelUnits.DefaultRowHeight + 0.5)
                {
                    heights[r + rowOffset] = Math.Min(Math.Max(word, estimate), ExcelUnits.MaxRowHeight);
                }
            }
        }

        internal static double EstimateHeight(TableModel table, CellModel cell)
        {
            string text = cell.PlainText.TrimEnd('\n');
            double size = 0;
            foreach (var run in cell.Runs)
            {
                if (run.Format.Size.HasValue && run.Format.Size.Value > 0 && run.Format.Size.Value < 410) size = Math.Max(size, run.Format.Size.Value);
            }
            if (size <= 0) size = XlsxStyleRegistry.DefaultFontSize;

            double lineHeight = size * 1.3;
            if (text.Length == 0) return lineHeight + 2;

            double charWidthPt = size * 0.5;
            if (cell.TextRotation != 0)
            {
                int longest = text.Split('\n').Max(l => l.Length);
                return Math.Min(longest * charWidthPt + 8, ExcelUnits.MaxRowHeight);
            }

            double width = 0;
            for (int c = cell.Column; c < cell.Column + cell.ColumnSpan; c++)
            {
                double w = table.ColumnWidthsPt != null && c < table.ColumnWidthsPt.Length ? table.ColumnWidthsPt[c] : 0;
                width += w > 0 ? w : ExcelUnits.ColumnWidthToPoints(8.43);
            }
            double available = Math.Max(10, width - 6);

            int lines = 0;
            foreach (var line in text.Split('\n'))
            {
                double needed = line.Length * charWidthPt;
                lines += Math.Max(1, (int)Math.Ceiling(needed / available));
            }
            return lines * lineHeight + 3;
        }

        // ------------------------------------------------------------------ XML

        /// <param name="frozenRows">Lignes figées en haut de la feuille (0 : aucune).</param>
        private byte[] BuildSheetXml(TableModel table, SortedDictionary<int, SortedDictionary<int, SheetCell>> cells,
            List<string> merges, Dictionary<int, double> rowHeights, int lastRow, int lastCol, bool selected, int frozenRows)
        {
            return Xml(w =>
            {
                w.WriteStartElement("worksheet", XlsxNames.Main);
                w.WriteAttributeString("xmlns", "r", null, XlsxNames.Relationships);

                w.WriteStartElement("dimension");
                w.WriteAttributeString("ref", "A1:" + ExcelUnits.CellReference(lastRow, lastCol));
                w.WriteEndElement();

                WriteSheetViews(w, selected, frozenRows);

                w.WriteStartElement("sheetFormatPr");
                w.WriteAttributeString("defaultRowHeight", "15");
                w.WriteEndElement();

                if (table.ColumnWidthsPt != null && table.ColumnWidthsPt.Any(x => x > 0))
                {
                    w.WriteStartElement("cols");
                    for (int c = 0; c < table.ColumnWidthsPt.Length; c++)
                    {
                        double width = ExcelUnits.PointsToColumnWidth(table.ColumnWidthsPt[c]);
                        if (width <= 0) continue;
                        w.WriteStartElement("col");
                        w.WriteAttributeString("min", (c + 1).ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("max", (c + 1).ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("width", width.ToString("0.##", CultureInfo.InvariantCulture));
                        w.WriteAttributeString("customWidth", "1");
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }

                w.WriteStartElement("sheetData");
                var rowIndexes = new SortedSet<int>(cells.Keys);
                foreach (var r in rowHeights.Keys) rowIndexes.Add(r);
                foreach (int r in rowIndexes)
                {
                    w.WriteStartElement("row");
                    w.WriteAttributeString("r", (r + 1).ToString(CultureInfo.InvariantCulture));
                    double height;
                    if (rowHeights.TryGetValue(r, out height))
                    {
                        w.WriteAttributeString("ht", height.ToString("0.##", CultureInfo.InvariantCulture));
                        w.WriteAttributeString("customHeight", "1");
                    }
                    SortedDictionary<int, SheetCell> line;
                    if (cells.TryGetValue(r, out line))
                    {
                        foreach (var entry in line) WriteCell(w, r, entry.Key, entry.Value);
                    }
                    w.WriteEndElement();
                }
                w.WriteEndElement(); // sheetData

                if (merges.Count > 0)
                {
                    w.WriteStartElement("mergeCells");
                    w.WriteAttributeString("count", merges.Count.ToString(CultureInfo.InvariantCulture));
                    foreach (var m in merges)
                    {
                        w.WriteStartElement("mergeCell");
                        w.WriteAttributeString("ref", m);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }

                WritePageMargins(w);

                w.WriteEndElement(); // worksheet
            });
        }

        private static void WriteCell(XmlWriter w, int row, int column, SheetCell cell)
        {
            w.WriteStartElement("c");
            w.WriteAttributeString("r", ExcelUnits.CellReference(row, column));
            if (cell.Style != 0) w.WriteAttributeString("s", cell.Style.ToString(CultureInfo.InvariantCulture));
            if (cell.Kind == CellKind.SharedString)
            {
                w.WriteAttributeString("t", "s");
                w.WriteElementString("v", XlsxNames.Main, cell.StringIndex.ToString(CultureInfo.InvariantCulture));
            }
            else if (cell.Kind == CellKind.Number)
            {
                w.WriteElementString("v", XlsxNames.Main, cell.Number.ToString("R", CultureInfo.InvariantCulture));
            }
            w.WriteEndElement();
        }

        /// <summary>Vue de la feuille : onglet actif ou non, lignes du haut figées (volets) si <paramref name="frozenRows"/> &gt; 0.</summary>
        private static void WriteSheetViews(XmlWriter w, bool selected, int frozenRows)
        {
            w.WriteStartElement("sheetViews");
            w.WriteStartElement("sheetView");
            if (selected) w.WriteAttributeString("tabSelected", "1");
            w.WriteAttributeString("workbookViewId", "0");
            if (frozenRows > 0)
            {
                string first = "A" + (frozenRows + 1).ToString(CultureInfo.InvariantCulture);
                w.WriteStartElement("pane");
                w.WriteAttributeString("ySplit", frozenRows.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("topLeftCell", first);
                w.WriteAttributeString("activePane", "bottomLeft");
                w.WriteAttributeString("state", "frozen");
                w.WriteEndElement();
                w.WriteStartElement("selection");
                w.WriteAttributeString("pane", "bottomLeft");
                w.WriteAttributeString("activeCell", first);
                w.WriteAttributeString("sqref", first);
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WritePageMargins(XmlWriter w)
        {
            w.WriteStartElement("pageMargins");
            w.WriteAttributeString("left", "0.7");
            w.WriteAttributeString("right", "0.7");
            w.WriteAttributeString("top", "0.75");
            w.WriteAttributeString("bottom", "0.75");
            w.WriteAttributeString("header", "0.3");
            w.WriteAttributeString("footer", "0.3");
            w.WriteEndElement();
        }

        private void WriteSharedStrings(XmlWriter w)
        {
            w.WriteStartElement("sst", XlsxNames.Main);
            w.WriteAttributeString("count", _stringReferences.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("uniqueCount", _strings.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var item in _strings)
            {
                w.WriteStartElement("si");
                var plain = item as string;
                if (plain != null)
                {
                    WriteText(w, plain);
                }
                else
                {
                    foreach (var run in (List<TextRun>)item)
                    {
                        w.WriteStartElement("r");
                        w.WriteStartElement("rPr");
                        XlsxStyleRegistry.WriteFontProperties(w, run.Format, "rFont");
                        w.WriteEndElement();
                        WriteText(w, run.Text);
                        w.WriteEndElement();
                    }
                }
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static void WriteText(XmlWriter w, string text)
        {
            w.WriteStartElement("t");
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(XmlSafe(text));
            w.WriteEndElement();
        }

        /// <summary>Supprime les caractères interdits en XML 1.0 (sécurité supplémentaire).</summary>
        internal static string XmlSafe(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder sb = null;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                bool ok;
                if (char.IsHighSurrogate(c))
                {
                    ok = i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]);
                    if (ok)
                    {
                        if (sb != null) sb.Append(c).Append(text[i + 1]);
                        i++;
                        continue;
                    }
                }
                else if (char.IsLowSurrogate(c)) ok = false;
                else ok = c == '\t' || c == '\n' || c == '\r' || (c >= ' ' && c != '\uFFFE' && c != '\uFFFF');

                if (!ok)
                {
                    if (sb == null) sb = new StringBuilder(text, 0, i, text.Length);
                    continue;
                }
                if (sb != null) sb.Append(c);
            }
            return sb == null ? text : sb.ToString();
        }

        private static byte[] ContentTypesXml(int sheetCount)
        {
            return Xml(w =>
            {
                w.WriteStartElement("Types", XlsxNames.ContentTypes);
                Default(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
                Default(w, "xml", "application/xml");
                Override(w, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                for (int i = 0; i < sheetCount; i++)
                {
                    Override(w, "/xl/worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                }
                Override(w, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                Override(w, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                Override(w, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
                Override(w, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                w.WriteEndElement();
            });
        }

        private static byte[] RootRelsXml()
        {
            return Xml(w =>
            {
                w.WriteStartElement("Relationships", XlsxNames.PackageRelationships);
                Relationship(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
                Relationship(w, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
                Relationship(w, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
                w.WriteEndElement();
            });
        }

        private static byte[] WorkbookXml(List<Sheet> sheets)
        {
            return Xml(w =>
            {
                w.WriteStartElement("workbook", XlsxNames.Main);
                w.WriteAttributeString("xmlns", "r", null, XlsxNames.Relationships);
                w.WriteStartElement("bookViews");
                w.WriteStartElement("workbookView");
                w.WriteAttributeString("activeTab", "0");
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteStartElement("sheets");
                for (int i = 0; i < sheets.Count; i++)
                {
                    w.WriteStartElement("sheet");
                    w.WriteAttributeString("name", XmlSafe(sheets[i].Name));
                    w.WriteAttributeString("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("id", XlsxNames.Relationships, "rId" + (i + 1).ToString(CultureInfo.InvariantCulture));
                    w.WriteEndElement();
                }
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        private static byte[] WorkbookRelsXml(int sheetCount)
        {
            return Xml(w =>
            {
                w.WriteStartElement("Relationships", XlsxNames.PackageRelationships);
                for (int i = 0; i < sheetCount; i++)
                {
                    Relationship(w, "rId" + (i + 1).ToString(CultureInfo.InvariantCulture),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        "worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml");
                }
                Relationship(w, "rId" + (sheetCount + 1).ToString(CultureInfo.InvariantCulture),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
                Relationship(w, "rId" + (sheetCount + 2).ToString(CultureInfo.InvariantCulture),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
                w.WriteEndElement();
            });
        }

        private byte[] CoreXml()
        {
            const string cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            const string dc = "http://purl.org/dc/elements/1.1/";
            const string dcterms = "http://purl.org/dc/terms/";
            const string xsi = "http://www.w3.org/2001/XMLSchema-instance";
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
            return Xml(w =>
            {
                w.WriteStartElement("cp", "coreProperties", cp);
                w.WriteAttributeString("xmlns", "dc", null, dc);
                w.WriteAttributeString("xmlns", "dcterms", null, dcterms);
                w.WriteAttributeString("xmlns", "xsi", null, xsi);
                if (!string.IsNullOrEmpty(_options.Title)) w.WriteElementString("dc", "title", dc, XmlSafe(_options.Title));
                w.WriteElementString("dc", "creator", dc, "Tableaux Word vers Excel");
                w.WriteStartElement("dcterms", "created", dcterms);
                w.WriteAttributeString("xsi", "type", xsi, "dcterms:W3CDTF");
                w.WriteString(now);
                w.WriteEndElement();
                w.WriteStartElement("dcterms", "modified", dcterms);
                w.WriteAttributeString("xsi", "type", xsi, "dcterms:W3CDTF");
                w.WriteString(now);
                w.WriteEndElement();
                w.WriteEndElement();
            });
        }

        private static byte[] AppXml()
        {
            return Xml(w =>
            {
                w.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
                w.WriteElementString("Application", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Microsoft Excel");
                w.WriteEndElement();
            });
        }

        private static void Default(XmlWriter w, string extension, string contentType)
        {
            w.WriteStartElement("Default");
            w.WriteAttributeString("Extension", extension);
            w.WriteAttributeString("ContentType", contentType);
            w.WriteEndElement();
        }

        private static void Override(XmlWriter w, string part, string contentType)
        {
            w.WriteStartElement("Override");
            w.WriteAttributeString("PartName", part);
            w.WriteAttributeString("ContentType", contentType);
            w.WriteEndElement();
        }

        private static void Relationship(XmlWriter w, string id, string type, string target)
        {
            w.WriteStartElement("Relationship");
            w.WriteAttributeString("Id", id);
            w.WriteAttributeString("Type", type);
            w.WriteAttributeString("Target", target);
            w.WriteEndElement();
        }

        private static byte[] Xml(Action<XmlWriter> write)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                CheckCharacters = true,
                NewLineHandling = NewLineHandling.None
            };
            using (var ms = new MemoryStream())
            {
                using (var w = XmlWriter.Create(ms, settings))
                {
                    w.WriteStartDocument(true);
                    write(w);
                    w.WriteEndDocument();
                }
                return ms.ToArray();
            }
        }
    }
}
