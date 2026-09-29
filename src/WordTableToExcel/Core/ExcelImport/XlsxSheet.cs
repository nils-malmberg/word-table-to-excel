using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;

namespace WordTableToExcel.Core.ExcelImport
{
    public enum XlsxCellType
    {
        /// <summary>Cellule sans valeur (seulement mise en forme).</summary>
        Blank,
        Number,
        /// <summary>Texte (chaîne partagée, chaîne en ligne ou résultat texte d'une formule).</summary>
        Text,
        Boolean,
        Error
    }

    /// <summary>Cellule lue dans une feuille : valeur enregistrée par Excel et index de style.</summary>
    public sealed class XlsxCell
    {
        public int Row;
        public int Column;
        public int StyleIndex;
        public XlsxCellType Type;
        /// <summary>Valeur numérique (nombre, date, heure) ou booléenne (0 / 1).</summary>
        public double Number;
        /// <summary>Texte (Type = Text) ou code d'erreur (« #DIV/0! »…).</summary>
        public XlsxRichText Text;
        public bool HasFormula;

        /// <summary>Formule sans résultat enregistré (classeur produit par un autre logiciel, non recalculé).</summary>
        public bool FormulaWithoutValue
        {
            get { return HasFormula && Type == XlsxCellType.Blank; }
        }

        public bool HasValue
        {
            get { return Type != XlsxCellType.Blank && !(Type == XlsxCellType.Text && Text.Text.Length == 0); }
        }

        public override string ToString()
        {
            string value = Type == XlsxCellType.Number || Type == XlsxCellType.Boolean ? Number.ToString("R", CultureInfo.InvariantCulture)
                         : Text == null ? string.Empty : Text.Text;
            return CellReference.Format(Row, Column) + " " + Type + " " + value;
        }
    }

    public sealed class XlsxRowInfo
    {
        public int Index;
        /// <summary>Hauteur en points (null = hauteur par défaut de la feuille).</summary>
        public double? Height;
        public bool CustomHeight;
        public bool Hidden;
        /// <summary>Style appliqué à toute la ligne (cellules non définies), ou -1.</summary>
        public int StyleIndex = -1;
    }

    public sealed class XlsxColumnInfo
    {
        /// <summary>Première et dernière colonne concernées (base 0).</summary>
        public int First;
        public int Last;
        /// <summary>Largeur en nombre de caractères (unité du fichier), null = largeur par défaut.</summary>
        public double? Width;
        public bool Hidden;
        /// <summary>Style appliqué à toute la colonne (cellules non définies), ou -1.</summary>
        public int StyleIndex = -1;
    }

    /// <summary>Tableau Excel (« Mettre sous forme de tableau ») présent dans la feuille.</summary>
    public sealed class XlsxTableInfo
    {
        public string Name;
        public CellRange Range;
        public int HeaderRowCount = 1;
        public int TotalsRowCount;
        public string StyleName;
        public bool ShowFirstColumn;
        public bool ShowLastColumn;
        public bool ShowRowStripes = true;
        public bool ShowColumnStripes;
        /// <summary>false si le tableau n'a aucune information de style (aucune mise en forme de tableau).</summary>
        public bool HasStyleInfo;
    }

    /// <summary>Règle de mise en forme conditionnelle (seules les règles simples sont reproduites).</summary>
    public sealed class XlsxConditionalRule
    {
        public List<CellRange> Ranges = new List<CellRange>();
        public string Type;
        public string Operator;
        public int DifferentialFormatId = -1;
        public int Priority;
        public bool StopIfTrue;
        public List<string> Formulas = new List<string>();
        /// <summary>Texte recherché (règles containsText, beginsWith, endsWith).</summary>
        public string Text;
        /// <summary>Règles top10 : rang, en pourcentage, par le bas.</summary>
        public int Rank = 10;
        public bool Percent;
        public bool Bottom;
        /// <summary>Règles aboveAverage : au-dessus (true) ou en dessous, égalité incluse, écart type.</summary>
        public bool AboveAverage = true;
        public bool EqualAverage;
        public int StandardDeviations;
        /// <summary>Nuances de couleurs : seuils (type, valeur) et couleurs correspondantes.</summary>
        public List<KeyValuePair<string, string>> ScaleThresholds = new List<KeyValuePair<string, string>>();
        public List<XElement> ScaleColors = new List<XElement>();

        /// <summary>Cellule d'ancrage des références relatives : coin supérieur gauche de la première plage.</summary>
        public int AnchorRow
        {
            get { return Ranges.Count > 0 ? Ranges[0].FirstRow : 0; }
        }

        public int AnchorColumn
        {
            get { return Ranges.Count > 0 ? Ranges[0].FirstColumn : 0; }
        }

        public bool Applies(int row, int column)
        {
            foreach (var r in Ranges)
            {
                if (r.Contains(row, column)) return true;
            }
            return false;
        }
    }

    public sealed class XlsxHyperlink
    {
        public CellRange Range;
        /// <summary>Adresse externe (URL, fichier) ; null pour un lien interne au classeur.</summary>
        public string Target;
        public string Location;
    }

    /// <summary>Contenu d'une feuille de calcul (xl/worksheets/sheetN.xml).</summary>
    public sealed class XlsxSheet
    {
        /// <summary>Nombre maximal de cellules lues dans une feuille (au-delà, l'import serait inexploitable dans Word).</summary>
        public const int MaxCells = 2000000;

        private readonly Dictionary<long, XlsxCell> _cells = new Dictionary<long, XlsxCell>();
        private readonly Dictionary<int, XlsxRowInfo> _rows = new Dictionary<int, XlsxRowInfo>();
        private readonly List<XlsxColumnInfo> _columns = new List<XlsxColumnInfo>();
        private readonly List<CellRange> _merges = new List<CellRange>();
        private readonly List<XlsxTableInfo> _tables = new List<XlsxTableInfo>();
        private readonly List<XlsxConditionalRule> _conditionalRules = new List<XlsxConditionalRule>();
        private readonly List<XlsxHyperlink> _hyperlinks = new List<XlsxHyperlink>();

        private XlsxSheet(XlsxWorkbook workbook, XlsxSheetInfo info)
        {
            Workbook = workbook;
            Info = info;
            BaseColumnWidth = 8;
            DefaultRowHeight = 15;
            ShowGridLines = true;
        }

        public XlsxWorkbook Workbook { get; private set; }
        public XlsxSheetInfo Info { get; private set; }

        public string Name
        {
            get { return Info.Name; }
        }

        /// <summary>Largeur par défaut des colonnes (defaultColWidth, en caractères avec marges), null = calculée depuis baseColWidth.</summary>
        public double? DefaultColumnWidth { get; private set; }
        public int BaseColumnWidth { get; private set; }
        public double DefaultRowHeight { get; private set; }
        /// <summary>Les lignes sont masquées par défaut (seules les lignes déclarées sont visibles).</summary>
        public bool RowsHiddenByDefault { get; private set; }
        public bool ShowGridLines { get; private set; }
        public bool PrintGridLines { get; private set; }
        public bool RightToLeft { get; private set; }
        public bool HasDrawings { get; private set; }
        public bool HasComments { get; private set; }
        public bool HasSparklines { get; private set; }
        /// <summary>Plage déclarée par Excel (dimension) ; indicative seulement.</summary>
        public CellRange? Dimension { get; private set; }
        public bool Truncated { get; private set; }

        public int CellCount
        {
            get { return _cells.Count; }
        }

        public IEnumerable<XlsxCell> Cells
        {
            get { return _cells.Values; }
        }

        public IList<XlsxColumnInfo> Columns
        {
            get { return _columns; }
        }

        public IList<CellRange> MergedRanges
        {
            get { return _merges; }
        }

        public IList<XlsxTableInfo> Tables
        {
            get { return _tables; }
        }

        public IList<XlsxConditionalRule> ConditionalRules
        {
            get { return _conditionalRules; }
        }

        public IList<XlsxHyperlink> Hyperlinks
        {
            get { return _hyperlinks; }
        }

        public XlsxCell Cell(int row, int column)
        {
            XlsxCell cell;
            return _cells.TryGetValue(Key(row, column), out cell) ? cell : null;
        }

        public XlsxRowInfo Row(int row)
        {
            XlsxRowInfo info;
            return _rows.TryGetValue(row, out info) ? info : null;
        }

        public IEnumerable<XlsxRowInfo> Rows
        {
            get { return _rows.Values; }
        }

        public XlsxColumnInfo Column(int column)
        {
            // Les déclarations sont peu nombreuses : recherche linéaire (la dernière l'emporte).
            XlsxColumnInfo found = null;
            foreach (var c in _columns)
            {
                if (column >= c.First && column <= c.Last) found = c;
            }
            return found;
        }

        public bool IsRowHidden(int row)
        {
            var info = Row(row);
            if (info != null) return info.Hidden;
            return RowsHiddenByDefault;
        }

        public bool IsColumnHidden(int column)
        {
            var info = Column(column);
            return info != null && info.Hidden;
        }

        /// <summary>Style effectif d'une position : cellule, sinon ligne, sinon colonne, sinon 0.</summary>
        public int StyleAt(int row, int column)
        {
            var cell = Cell(row, column);
            if (cell != null) return cell.StyleIndex;
            var r = Row(row);
            if (r != null && r.StyleIndex >= 0) return r.StyleIndex;
            var c = Column(column);
            if (c != null && c.StyleIndex >= 0) return c.StyleIndex;
            return 0;
        }

        /// <summary>Plage fusionnée contenant la position, ou null.</summary>
        public CellRange? MergeAt(int row, int column)
        {
            foreach (var m in _merges)
            {
                if (m.Contains(row, column)) return m;
            }
            return null;
        }

        private static long Key(int row, int column)
        {
            return ((long)row << 15) | (uint)column;
        }

        // ------------------------------------------------------------------ lecture

        internal static XlsxSheet Read(XlsxWorkbook workbook, ZipReader zip, XlsxSheetInfo info)
        {
            var sheet = new XlsxSheet(workbook, info);
            var rels = XlsxPackage.Relationships(zip, info.PartName);
            using (var reader = XlsxPackage.OpenReader(zip, info.PartName))
            {
                if (reader == null) throw new ExcelImportException("La feuille « " + info.Name + " » est introuvable dans le classeur.");
                sheet.Parse(reader, rels);
            }

            foreach (var rel in rels)
            {
                if (rel.External) continue;
                if (rel.IsOfType("table")) sheet.ReadTable(zip, rel.Target);
                else if (rel.IsOfType("drawing")) sheet.HasDrawings = true;
                else if (rel.IsOfType("comments")) sheet.HasComments = true;
            }
            return sheet;
        }

        private void Parse(XmlReader reader, List<OpcRelationship> rels)
        {
            int currentRow = -1;
            int nextColumn = 0;
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element)
                {
                    reader.Read();
                    continue;
                }
                switch (reader.LocalName)
                {
                    case "sheetPr":
                    case "sheetViews":
                    case "sheetFormatPr":
                    case "cols":
                    case "mergeCells":
                    case "conditionalFormatting":
                    case "hyperlinks":
                    case "printOptions":
                    case "extLst":
                        ReadSection((XElement)XNode.ReadFrom(reader), rels);
                        continue;
                    case "dimension":
                        CellRange dim;
                        if (CellRange.TryParse(reader.GetAttribute("ref"), out dim)) Dimension = dim;
                        reader.Read();
                        continue;
                    case "row":
                        currentRow = ReadRowStart(reader, currentRow);
                        nextColumn = 0;
                        reader.Read();
                        continue;
                    case "c":
                        if (currentRow < 0) currentRow = 0;
                        var cell = ReadCell(reader, currentRow, nextColumn);
                        if (cell != null)
                        {
                            nextColumn = cell.Column + 1;
                            if (_cells.Count >= MaxCells)
                            {
                                Truncated = true;
                                return;
                            }
                            _cells[Key(cell.Row, cell.Column)] = cell;
                        }
                        continue;
                    default:
                        reader.Read();
                        continue;
                }
            }
        }

        private int ReadRowStart(XmlReader reader, int previousRow)
        {
            int row = previousRow + 1;
            string r = reader.GetAttribute("r");
            int parsed;
            if (r != null && int.TryParse(r, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) && parsed >= 1 && parsed <= CellReference.MaxRows)
            {
                row = parsed - 1;
            }
            var info = new XlsxRowInfo { Index = row };
            double ht;
            string htText = reader.GetAttribute("ht");
            if (htText != null && double.TryParse(htText, NumberStyles.Float, CultureInfo.InvariantCulture, out ht) && ht >= 0) info.Height = ht;
            info.CustomHeight = OnOff(reader.GetAttribute("customHeight"));
            info.Hidden = OnOff(reader.GetAttribute("hidden"));
            if (OnOff(reader.GetAttribute("customFormat")))
            {
                int s;
                if (int.TryParse(reader.GetAttribute("s"), NumberStyles.Integer, CultureInfo.InvariantCulture, out s)) info.StyleIndex = s;
            }
            _rows[row] = info;
            return row;
        }

        private XlsxCell ReadCell(XmlReader reader, int row, int nextColumn)
        {
            var cell = new XlsxCell { Row = row, Column = nextColumn };
            string reference = reader.GetAttribute("r");
            int r, c;
            if (reference != null && CellReference.TryParse(reference, out r, out c))
            {
                cell.Row = r;
                cell.Column = c;
            }
            int style;
            if (int.TryParse(reader.GetAttribute("s"), NumberStyles.Integer, CultureInfo.InvariantCulture, out style) && style >= 0) cell.StyleIndex = style;
            string type = reader.GetAttribute("t") ?? "n";

            string value = null;
            XlsxRichText inline = null;
            if (reader.IsEmptyElement)
            {
                reader.Read();
            }
            else
            {
                int depth = reader.Depth;
                reader.Read();
                while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        switch (reader.LocalName)
                        {
                            case "v":
                                value = reader.ReadElementContentAsString();
                                continue;
                            case "f":
                                cell.HasFormula = true;
                                reader.Skip();
                                continue;
                            case "is":
                                inline = XlsxWorkbook.ReadStringItem(reader);
                                continue;
                            default:
                                reader.Skip();
                                continue;
                        }
                    }
                    reader.Read();
                }
                reader.Read(); // </c>
            }

            switch (type)
            {
                case "s":
                    int index;
                    if (value != null && int.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
                    {
                        var shared = Workbook.SharedString(index);
                        if (shared != null)
                        {
                            cell.Type = XlsxCellType.Text;
                            cell.Text = shared;
                        }
                    }
                    break;
                case "inlineStr":
                    if (inline != null)
                    {
                        cell.Type = XlsxCellType.Text;
                        cell.Text = inline;
                    }
                    else if (value != null)
                    {
                        cell.Type = XlsxCellType.Text;
                        cell.Text = new XlsxRichText(XlsxPackage.DecodeEscapes(value), null);
                    }
                    break;
                case "str":
                    if (value != null)
                    {
                        cell.Type = XlsxCellType.Text;
                        cell.Text = new XlsxRichText(XlsxPackage.DecodeEscapes(value), null);
                    }
                    break;
                case "b":
                    if (value != null)
                    {
                        string v = value.Trim();
                        cell.Type = XlsxCellType.Boolean;
                        cell.Number = v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    }
                    break;
                case "e":
                    if (value != null)
                    {
                        cell.Type = XlsxCellType.Error;
                        cell.Text = new XlsxRichText(value.Trim(), null);
                    }
                    break;
                case "d":
                    double serial;
                    if (value != null && ExcelDates.TryParseIso(value.Trim(), Workbook.Date1904, out serial))
                    {
                        cell.Type = XlsxCellType.Number;
                        cell.Number = serial;
                    }
                    else if (!string.IsNullOrEmpty(value))
                    {
                        cell.Type = XlsxCellType.Text;
                        cell.Text = new XlsxRichText(value, null);
                    }
                    break;
                default:
                    double number;
                    if (value != null && double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                        && !double.IsNaN(number) && !double.IsInfinity(number))
                    {
                        cell.Type = XlsxCellType.Number;
                        cell.Number = number;
                    }
                    else if (!string.IsNullOrEmpty(value))
                    {
                        // Valeur non numérique dans une cellule numérique : on l'affiche telle quelle plutôt que de la perdre.
                        cell.Type = XlsxCellType.Text;
                        cell.Text = new XlsxRichText(value, null);
                    }
                    break;
            }
            return cell;
        }

        private void ReadSection(XElement section, List<OpcRelationship> rels)
        {
            switch (section.Name.LocalName)
            {
                case "sheetPr":
                    break;
                case "sheetViews":
                    var view = section.Elements().FirstOrDefault(e => OoxmlXml.Is(e, "sheetView"));
                    if (view != null)
                    {
                        bool b;
                        if (OoxmlXml.TryBoolAttr(view, "showGridLines", out b)) ShowGridLines = b;
                        if (OoxmlXml.TryBoolAttr(view, "rightToLeft", out b)) RightToLeft = b;
                    }
                    break;
                case "sheetFormatPr":
                    double d;
                    string defaultWidth = OoxmlXml.Attr(section, "defaultColWidth");
                    if (defaultWidth != null && double.TryParse(defaultWidth, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d > 0) DefaultColumnWidth = d;
                    int baseWidth = OoxmlXml.IntAttr(section, "baseColWidth", 8);
                    if (baseWidth > 0 && baseWidth <= 255) BaseColumnWidth = baseWidth;
                    string defaultHeight = OoxmlXml.Attr(section, "defaultRowHeight");
                    if (defaultHeight != null && double.TryParse(defaultHeight, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d > 0) DefaultRowHeight = d;
                    bool zero;
                    if (OoxmlXml.TryBoolAttr(section, "zeroHeight", out zero)) RowsHiddenByDefault = zero;
                    break;
                case "cols":
                    foreach (var col in section.Elements().Where(e => OoxmlXml.Is(e, "col")))
                    {
                        int min = OoxmlXml.IntAttr(col, "min", 0);
                        int max = OoxmlXml.IntAttr(col, "max", min);
                        if (min < 1 || max < min) continue;
                        var info = new XlsxColumnInfo { First = min - 1, Last = Math.Min(max, CellReference.MaxColumns) - 1 };
                        string w = OoxmlXml.Attr(col, "width");
                        double width;
                        if (w != null && double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out width) && width >= 0) info.Width = width;
                        bool hidden;
                        if (OoxmlXml.TryBoolAttr(col, "hidden", out hidden)) info.Hidden = hidden;
                        if (info.Width.HasValue && info.Width.Value == 0) info.Hidden = true;
                        string s = OoxmlXml.Attr(col, "style");
                        int styleIndex;
                        if (s != null && int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out styleIndex)) info.StyleIndex = styleIndex;
                        _columns.Add(info);
                    }
                    break;
                case "mergeCells":
                    foreach (var mc in section.Elements().Where(e => OoxmlXml.Is(e, "mergeCell")))
                    {
                        CellRange range;
                        if (CellRange.TryParse(OoxmlXml.Attr(mc, "ref"), out range) && (range.RowCount > 1 || range.ColumnCount > 1))
                        {
                            if (!_merges.Any(m => m.Intersects(range))) _merges.Add(range); // fusions qui se chevauchent : fichier incohérent
                        }
                    }
                    break;
                case "conditionalFormatting":
                    ReadConditionalFormatting(section);
                    break;
                case "hyperlinks":
                    foreach (var h in section.Elements().Where(e => OoxmlXml.Is(e, "hyperlink")))
                    {
                        CellRange range;
                        if (!CellRange.TryParse(OoxmlXml.Attr(h, "ref"), out range)) continue;
                        var link = new XlsxHyperlink { Range = range, Location = OoxmlXml.Attr(h, "location") };
                        string id = null;
                        foreach (var a in h.Attributes())
                        {
                            if (a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None) id = a.Value;
                        }
                        if (id != null)
                        {
                            var rel = rels.FirstOrDefault(r => string.Equals(r.Id, id, StringComparison.Ordinal));
                            if (rel != null && rel.External) link.Target = rel.Target;
                        }
                        if (link.Target != null || link.Location != null) _hyperlinks.Add(link);
                    }
                    break;
                case "printOptions":
                    bool grid;
                    if (OoxmlXml.TryBoolAttr(section, "gridLines", out grid)) PrintGridLines = grid;
                    break;
                case "extLst":
                    foreach (var e in section.Descendants())
                    {
                        if (e.Name.LocalName == "sparklineGroups") HasSparklines = true;
                        else if (e.Name.LocalName == "conditionalFormatting") ReadConditionalFormatting(e);
                    }
                    break;
            }
        }

        private void ReadConditionalFormatting(XElement section)
        {
            var ranges = new List<CellRange>();
            string sqref = OoxmlXml.Attr(section, "sqref");
            if (sqref == null)
            {
                var sq = section.Elements().FirstOrDefault(e => e.Name.LocalName == "sqref");
                if (sq != null) sqref = sq.Value;
            }
            if (sqref != null)
            {
                foreach (var part in sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    CellRange range;
                    if (CellRange.TryParse(part, out range)) ranges.Add(range);
                }
            }
            foreach (var rule in section.Elements().Where(e => OoxmlXml.Is(e, "cfRule")))
            {
                var r = new XlsxConditionalRule
                {
                    Ranges = ranges,
                    Type = OoxmlXml.Attr(rule, "type"),
                    Operator = OoxmlXml.Attr(rule, "operator"),
                    DifferentialFormatId = OoxmlXml.IntAttr(rule, "dxfId", -1),
                    Priority = OoxmlXml.IntAttr(rule, "priority", int.MaxValue),
                    Text = OoxmlXml.Attr(rule, "text")
                };
                bool flag;
                if (OoxmlXml.TryBoolAttr(rule, "stopIfTrue", out flag)) r.StopIfTrue = flag;
                if (OoxmlXml.TryBoolAttr(rule, "percent", out flag)) r.Percent = flag;
                if (OoxmlXml.TryBoolAttr(rule, "bottom", out flag)) r.Bottom = flag;
                if (OoxmlXml.TryBoolAttr(rule, "aboveAverage", out flag)) r.AboveAverage = flag;
                if (OoxmlXml.TryBoolAttr(rule, "equalAverage", out flag)) r.EqualAverage = flag;
                r.Rank = OoxmlXml.IntAttr(rule, "rank", 10);
                r.StandardDeviations = OoxmlXml.IntAttr(rule, "stdDev", 0);
                foreach (var f in rule.Elements().Where(e => e.Name.LocalName == "formula" || e.Name.LocalName == "f")) r.Formulas.Add(f.Value);
                var scale = rule.Elements().FirstOrDefault(e => e.Name.LocalName == "colorScale");
                if (scale != null)
                {
                    foreach (var e in scale.Elements())
                    {
                        if (e.Name.LocalName == "cfvo")
                        {
                            string val = OoxmlXml.Attr(e, "val");
                            if (val == null)
                            {
                                var f = e.Elements().FirstOrDefault(x => x.Name.LocalName == "f");
                                if (f != null) val = f.Value;
                            }
                            r.ScaleThresholds.Add(new KeyValuePair<string, string>(OoxmlXml.Attr(e, "type") ?? "min", val));
                        }
                        else if (e.Name.LocalName == "color")
                        {
                            r.ScaleColors.Add(e);
                        }
                    }
                }
                _conditionalRules.Add(r);
            }
        }

        private void ReadTable(ZipReader zip, string partName)
        {
            XDocument doc;
            try
            {
                doc = XlsxPackage.LoadXml(zip, partName);
            }
            catch (ExcelImportException)
            {
                return; // tableau illisible : les cellules restent, seule sa mise en forme est perdue
            }
            if (doc == null || doc.Root == null) return;
            var root = doc.Root;
            CellRange range;
            if (!CellRange.TryParse(OoxmlXml.Attr(root, "ref"), out range)) return;
            var table = new XlsxTableInfo
            {
                Name = OoxmlXml.Attr(root, "displayName") ?? OoxmlXml.Attr(root, "name"),
                Range = range,
                HeaderRowCount = Math.Max(0, OoxmlXml.IntAttr(root, "headerRowCount", 1)),
                TotalsRowCount = Math.Max(0, OoxmlXml.IntAttr(root, "totalsRowCount", 0))
            };
            var info = OoxmlXml.Child(root, "tableStyleInfo");
            if (info != null)
            {
                table.HasStyleInfo = true;
                table.StyleName = OoxmlXml.Attr(info, "name");
                bool b;
                table.ShowFirstColumn = OoxmlXml.TryBoolAttr(info, "showFirstColumn", out b) && b;
                table.ShowLastColumn = OoxmlXml.TryBoolAttr(info, "showLastColumn", out b) && b;
                table.ShowRowStripes = !OoxmlXml.TryBoolAttr(info, "showRowStripes", out b) || b;
                table.ShowColumnStripes = OoxmlXml.TryBoolAttr(info, "showColumnStripes", out b) && b;
            }
            _tables.Add(table);
        }

        private static bool OnOff(string value)
        {
            return value != null && OoxmlXml.ParseOnOff(value);
        }
    }
}
