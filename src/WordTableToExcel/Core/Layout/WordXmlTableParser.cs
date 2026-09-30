using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Layout
{
    /// <summary>
    /// Analyse le XML d'un tableau fourni par Word (<c>Range.WordOpenXML</c>, Word 2010+,
    /// ou <c>Range.XML</c>, Word 2003+) pour obtenir la structure exacte du tableau :
    /// grille, cellules fusionnées horizontalement (gridSpan) et verticalement (vMerge),
    /// décalages de ligne (gridBefore), largeurs de colonnes, hauteurs de lignes,
    /// ainsi que le fond, les bordures et l'orientation de chaque cellule, y compris ceux
    /// hérités du style de tableau (ligne d'en-tête, lignes à bandes, première colonne…).
    /// </summary>
    public static class WordXmlTableParser
    {
        private const double TwipsPerPoint = 20.0;

        private sealed class RawCell
        {
            public XElement Tc;
            public XElement TcPr;
            public int SourceRow;
            public int SourceIndex;
            public int SourceOrdinal;
            public int GridColumn;
            public int GridSpan;
            public bool IsContinuation;
            public LayoutCell Layout;
        }

        /// <summary>Analyse le XML ; renvoie null si aucun tableau n'y figure.</summary>
        /// <param name="themeFontFallback">Polices du thème (vrai = titres), si le XML ne contient pas le thème.</param>
        public static TableLayout Parse(string xml, Func<bool, string> themeFontFallback = null)
        {
            if (string.IsNullOrEmpty(xml)) return null;
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CheckCharacters = false,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true
            };
            XDocument document;
            using (var sr = new System.IO.StringReader(xml))
            using (var reader = XmlReader.Create(sr, settings))
            {
                document = XDocument.Load(reader);
            }
            return Parse(document, themeFontFallback);
        }

        public static TableLayout Parse(XDocument document, Func<bool, string> themeFontFallback = null)
        {
            var root = document.Root;
            if (root == null) return null;

            XElement body = root.Descendants().FirstOrDefault(e => OoxmlXml.Is(e, "body"));
            XElement tbl = (body ?? root).Descendants().FirstOrDefault(e => OoxmlXml.Is(e, "tbl"));
            if (tbl == null) return null;
            return Parse(root, tbl, themeFontFallback);
        }

        /// <summary>Analyse le tableau <paramref name="tbl"/> du document <paramref name="root"/> (styles et thème compris).</summary>
        public static TableLayout Parse(XElement root, XElement tbl, Func<bool, string> themeFontFallback = null)
        {
            var styles = new TableStyleSheet(root);
            var tblPr = OoxmlXml.Child(tbl, "tblPr");
            string styleId = OoxmlXml.Attr(OoxmlXml.Child(tblPr, "tblStyle"), "val");
            var chain = styles.Chain(styleId);
            var look = TableLook.Parse(OoxmlXml.Child(tblPr, "tblLook"));

            var layout = new TableLayout { HasCellFormatting = true };

            // Grille déclarée.
            var gridWidths = new List<double>();
            var tblGrid = OoxmlXml.Child(tbl, "tblGrid");
            if (tblGrid != null)
            {
                foreach (var gc in tblGrid.Elements().Where(e => OoxmlXml.Is(e, "gridCol")))
                {
                    gridWidths.Add(Math.Max(0, OoxmlXml.IntAttr(gc, "w", 0)) / TwipsPerPoint);
                }
            }

            // Lignes supprimées en suivi des modifications (même non acceptées) : absentes du tableau exporté.
            // Leur rang d'origine est conservé (SourceRow) pour rester aligné sur les lignes de Word.
            var allRows = StructuralChildren(tbl, "tr").ToList();
            var rows = new List<XElement>();
            var sourceRows = new List<int>();
            for (int i = 0; i < allRows.Count; i++)
            {
                if (IsDeletedRow(allRows[i]))
                {
                    layout.DeletedRows.Add(i);
                    continue;
                }
                rows.Add(allRows[i]);
                sourceRows.Add(i);
            }
            layout.SourceRowCount = allRows.Count;
            layout.RowCount = rows.Count;
            layout.RowHeightsPt = new double[rows.Count];
            layout.RowHeightExact = new bool[rows.Count];
            layout.SourceCellCounts = new int[allRows.Count];
            foreach (int deleted in layout.DeletedRows) layout.SourceCellCounts[deleted] = StructuralChildren(allRows[deleted], "tc").Count();

            var rawRows = new List<List<RawCell>>();
            int headerRows = 0;
            bool headerRun = true;
            int maxColumns = gridWidths.Count;

            for (int r = 0; r < rows.Count; r++)
            {
                var tr = rows[r];
                var trPr = OoxmlXml.Child(tr, "trPr");
                int column = Math.Max(0, OoxmlXml.IntAttr(OoxmlXml.Child(trPr, "gridBefore"), "val", 0));

                var trHeight = OoxmlXml.Child(trPr, "trHeight");
                if (trHeight != null)
                {
                    layout.RowHeightsPt[r] = Math.Max(0, OoxmlXml.IntAttr(trHeight, "val", 0)) / TwipsPerPoint;
                    string rule = (OoxmlXml.Attr(trHeight, "hRule") ?? OoxmlXml.Attr(trHeight, "h-rule") ?? "atLeast").ToLowerInvariant();
                    layout.RowHeightExact[r] = rule == "exact";
                    if (rule == "auto") layout.RowHeightsPt[r] = 0;
                }

                if (headerRun && OoxmlXml.IsOn(OoxmlXml.Child(trPr, "tblHeader"))) headerRows++;
                else headerRun = false;

                var rawRow = new List<RawCell>();
                int index = 0, ordinal = 0;
                foreach (var tc in StructuralChildren(tr, "tc"))
                {
                    var tcPr = OoxmlXml.Child(tc, "tcPr");
                    int span = Math.Max(1, OoxmlXml.IntAttr(OoxmlXml.Child(tcPr, "gridSpan"), "val", 1));
                    var vMerge = OoxmlXml.Child(tcPr, "vMerge");
                    string vMergeVal = OoxmlXml.Attr(vMerge, "val");
                    bool continuation = vMerge != null && (vMergeVal == null || !string.Equals(vMergeVal, "restart", StringComparison.OrdinalIgnoreCase));

                    var raw = new RawCell
                    {
                        Tc = tc,
                        TcPr = tcPr,
                        SourceRow = sourceRows[r],
                        SourceIndex = index++,
                        SourceOrdinal = continuation ? -1 : ordinal,
                        GridColumn = column,
                        GridSpan = span,
                        IsContinuation = continuation
                    };
                    if (!continuation) ordinal++;
                    rawRow.Add(raw);
                    column += span;
                }
                column += Math.Max(0, OoxmlXml.IntAttr(OoxmlXml.Child(trPr, "gridAfter"), "val", 0));
                maxColumns = Math.Max(maxColumns, column);
                layout.SourceCellCounts[sourceRows[r]] = index;
                rawRows.Add(rawRow);
            }

            layout.ColumnCount = maxColumns;
            layout.ColumnWidthsPt = new double[maxColumns];
            for (int c = 0; c < maxColumns && c < gridWidths.Count; c++) layout.ColumnWidthsPt[c] = gridWidths[c];

            // Placement dans la grille et fusions verticales.
            var owners = new LayoutCell[rows.Count, Math.Max(1, maxColumns)];
            for (int r = 0; r < rawRows.Count; r++)
            {
                foreach (var raw in rawRows[r])
                {
                    LayoutCell above = r > 0 && raw.GridColumn < maxColumns ? owners[r - 1, raw.GridColumn] : null;
                    if (raw.IsContinuation && above != null && above.Column == raw.GridColumn)
                    {
                        above.RowSpan = r - above.Row + 1;
                        raw.Layout = above;
                    }
                    else
                    {
                        if (raw.IsContinuation)
                        {
                            // Continuation orpheline (XML incohérent) : traitée comme une cellule normale.
                            raw.IsContinuation = false;
                            raw.SourceOrdinal = rawRows[r].Where(x => x.SourceIndex < raw.SourceIndex && !x.IsContinuation).Count();
                            foreach (var later in rawRows[r].Where(x => x.SourceIndex > raw.SourceIndex && !x.IsContinuation)) later.SourceOrdinal++;
                        }
                        raw.Layout = new LayoutCell
                        {
                            Row = r,
                            Column = raw.GridColumn,
                            ColumnSpan = raw.GridSpan,
                            SourceRow = raw.SourceRow,
                            SourceIndex = raw.SourceIndex,
                            SourceOrdinal = raw.SourceOrdinal,
                            SourceTag = raw.Tc
                        };
                        layout.Cells.Add(raw.Layout);
                    }
                    for (int c = raw.GridColumn; c < raw.GridColumn + raw.GridSpan && c < maxColumns; c++)
                    {
                        owners[r, c] = raw.Layout;
                    }
                }
            }

            // Mise en forme des cellules (directe + style de tableau).
            var context = new StyleContext(layout, chain, look, tblPr, Math.Max(1, headerRows));
            var tcPrByCell = rawRows.SelectMany(x => x).Where(x => !x.IsContinuation).ToDictionary(x => x.Layout, x => x);
            foreach (var cell in layout.Cells)
            {
                RawCell raw;
                tcPrByCell.TryGetValue(cell, out raw);
                context.Apply(cell, raw == null ? null : raw.TcPr, raw == null ? null : raw.Tc);
                cell.SourceTag = null;
            }

            // Texte et mise en forme des caractères, lus dans le même XML (Open XML uniquement).
            try
            {
                var content = WordXmlContentReader.TryCreate(root, tbl, themeFontFallback);
                if (content != null)
                {
                    foreach (var cell in layout.Cells)
                    {
                        RawCell raw;
                        cell.Content = tcPrByCell.TryGetValue(cell, out raw) ? content.ReadCell(raw.Tc, context.Formats(cell)) : new XmlCellContent();
                    }
                    layout.XmlTextVariants.AddRange(content.TableTexts(tbl));
                    layout.XmlText = layout.XmlTextVariants[0];
                    layout.HasContent = true;
                }
            }
            catch (Exception)
            {
                // XML inattendu : le contenu sera lu par Word, cellule par cellule.
                foreach (var cell in layout.Cells) cell.Content = null;
                layout.XmlText = null;
                layout.XmlTextVariants.Clear();
                layout.HasContent = false;
            }
            return layout;
        }

        /// <summary>Ligne supprimée en suivi des modifications (w:trPr/w:del), acceptée ou non.</summary>
        internal static bool IsDeletedRow(XElement tr)
        {
            return OoxmlXml.Child(tr, "trPr", "del") != null;
        }

        /// <summary>Enfants « structurels » (tr, tc) en traversant les contrôles de contenu et le XML personnalisé.</summary>
        internal static IEnumerable<XElement> StructuralChildren(XElement parent, string localName)
        {
            foreach (var child in parent.Elements())
            {
                if (OoxmlXml.Is(child, localName))
                {
                    yield return child;
                    continue;
                }
                XElement container = null;
                if (OoxmlXml.Is(child, "sdt")) container = OoxmlXml.Child(child, "sdtContent");
                else if (OoxmlXml.Is(child, "customXml") || OoxmlXml.Is(child, "smartTag")) container = child;
                else if (child.Name.Namespace != parent.Name.Namespace && !OoxmlXml.Is(child, "AlternateContent")) container = child; // XML personnalisé Word 2003
                if (container == null) continue;
                foreach (var nested in StructuralChildren(container, localName)) yield return nested;
            }
        }

        /// <summary>Résolution de la mise en forme d'une cellule selon les règles de priorité de Word.</summary>
        private sealed class StyleContext
        {
            private readonly TableLayout _layout;
            private readonly List<XElement> _chain;
            private readonly TableLook _look;
            private readonly XElement _tblPr;
            private readonly int _headerRows;
            private readonly int _rowBand;
            private readonly int _colBand;

            public StyleContext(TableLayout layout, List<XElement> chain, TableLook look, XElement tblPr, int headerRows)
            {
                _layout = layout;
                _chain = chain;
                _look = look;
                _tblPr = tblPr;
                _headerRows = Math.Min(headerRows, Math.Max(1, layout.RowCount));
                _rowBand = TableStyleSheet.BandSize(chain, "tblStyleRowBandSize");
                _colBand = TableStyleSheet.BandSize(chain, "tblStyleColBandSize");
            }

            public void Apply(LayoutCell cell, XElement tcPr, XElement tc)
            {
                var regions = Regions(cell); // priorité décroissante, « tableau entier » en dernier

                // Fond : cellule > zones conditionnelles > style (tableau entier) > trame du tableau.
                Defined<Rgb?> fill = OoxmlFormat.ReadShading(OoxmlXml.Child(tcPr, "shd"));
                foreach (var pr in FormatsFor(regions))
                {
                    if (fill.IsDefined) break;
                    fill = OoxmlFormat.ReadShading(OoxmlXml.Child(pr, "tcPr", "shd"));
                    if (!fill.IsDefined) fill = OoxmlFormat.ReadShading(OoxmlXml.Child(pr, "tblPr", "shd"));
                }
                if (!fill.IsDefined) fill = OoxmlFormat.ReadShading(OoxmlXml.Child(_tblPr, "shd"));
                cell.Fill = fill.IsDefined ? fill.Value : null;

                // Alignement vertical.
                var valign = OoxmlFormat.ReadVerticalAlignment(tcPr);
                foreach (var pr in FormatsFor(regions))
                {
                    if (valign.IsDefined) break;
                    valign = OoxmlFormat.ReadVerticalAlignment(OoxmlXml.Child(pr, "tcPr"));
                }
                cell.VerticalAlignment = valign.IsDefined ? valign.Value : VerticalAlignment.Top;
                cell.TextRotation = OoxmlFormat.ReadTextRotation(tcPr);

                // Bordures.
                var tcBorders = OoxmlXml.Child(tcPr, "tcBorders");
                cell.Borders = new CellBorders
                {
                    Top = ResolveEdge(cell, tcBorders, regions, "top"),
                    Bottom = ResolveEdge(cell, tcBorders, regions, "bottom"),
                    Left = ResolveEdge(cell, tcBorders, regions, "left"),
                    Right = ResolveEdge(cell, tcBorders, regions, "right")
                };

                // Trame de paragraphe (souvent utilisée à la place du fond de cellule).
                if (tc != null)
                {
                    foreach (var p in tc.Elements().Where(e => OoxmlXml.Is(e, "p")))
                    {
                        var shd = OoxmlFormat.ReadShading(OoxmlXml.Child(p, "pPr", "shd"));
                        if (shd.IsDefined && shd.Value.HasValue)
                        {
                            cell.ParagraphShading = shd.Value;
                            break;
                        }
                    }
                }
            }

            /// <summary>Mise en forme du style de tableau applicable à la cellule, de la plus prioritaire à la moins prioritaire.</summary>
            public List<XElement> Formats(LayoutCell cell)
            {
                return FormatsFor(Regions(cell)).ToList();
            }

            private BorderLine ResolveEdge(LayoutCell cell, XElement tcBorders, List<TableRegion> regions, string edge)
            {
                // 1. Bordure directe de la cellule.
                var direct = OoxmlFormat.ReadEdge(tcBorders, edge);
                if (direct.IsDefined) return direct.Value;

                // 2. Bordures directes du tableau (bord extérieur ou intérieur selon la position).
                var table = OoxmlFormat.ReadEdge(OoxmlXml.Child(_tblPr, "tblBorders"), EdgeKey(cell, TableRegion.WholeTable, edge));
                if (table.IsDefined) return table.Value;

                // 3. Style de tableau : zones conditionnelles puis tableau entier, styles parents compris.
                foreach (var region in regions)
                {
                    string key = EdgeKey(cell, region, edge);
                    foreach (var pr in FormatsFor(region))
                    {
                        var b = OoxmlFormat.ReadEdge(OoxmlXml.Child(pr, "tcPr", "tcBorders"), key);
                        if (!b.IsDefined) b = OoxmlFormat.ReadEdge(OoxmlXml.Child(pr, "tblPr", "tblBorders"), key);
                        if (b.IsDefined) return b.Value;
                    }
                }
                return BorderLine.None;
            }

            private IEnumerable<XElement> FormatsFor(IEnumerable<TableRegion> regions)
            {
                foreach (var region in regions)
                {
                    foreach (var pr in FormatsFor(region)) yield return pr;
                }
            }

            /// <summary>
            /// Éléments porteurs de mise en forme pour une zone. Pour le tableau entier : la zone
            /// « wholeTable » éventuelle puis les propriétés du style lui-même (tblPr / tcPr).
            /// </summary>
            private IEnumerable<XElement> FormatsFor(TableRegion region)
            {
                if (region != TableRegion.WholeTable)
                {
                    foreach (var pr in TableStyleSheet.ConditionalFormats(_chain, region)) yield return pr;
                    yield break;
                }
                foreach (var style in _chain)
                {
                    foreach (var pr in TableStyleSheet.ConditionalFormats(new List<XElement> { style }, region)) yield return pr;
                    yield return style;
                }
            }

            /// <summary>Zones conditionnelles qui s'appliquent à la cellule, de la plus prioritaire à la moins prioritaire.</summary>
            private List<TableRegion> Regions(LayoutCell cell)
            {
                var list = new List<TableRegion>();
                int lastRow = _layout.RowCount - 1;
                int lastCol = _layout.ColumnCount - 1;

                bool inFirstRow = _look.FirstRow && cell.Row < _headerRows;
                bool inLastRow = _look.LastRow && cell.LastRow == lastRow;
                bool inFirstCol = _look.FirstColumn && cell.Column == 0;
                bool inLastCol = _look.LastColumn && cell.LastColumn == lastCol;

                list.Add(TableRegion.WholeTable);

                if (!_look.NoVerticalBand)
                {
                    int start = _look.FirstColumn ? 1 : 0;
                    if (cell.Column >= start)
                    {
                        int band = (cell.Column - start) / _colBand;
                        list.Add(band % 2 == 0 ? TableRegion.Band1Vertical : TableRegion.Band2Vertical);
                    }
                }
                if (!_look.NoHorizontalBand)
                {
                    int start = _look.FirstRow ? _headerRows : 0;
                    if (cell.Row >= start)
                    {
                        int band = (cell.Row - start) / _rowBand;
                        list.Add(band % 2 == 0 ? TableRegion.Band1Horizontal : TableRegion.Band2Horizontal);
                    }
                }
                if (inFirstCol) list.Add(TableRegion.FirstColumn);
                if (inLastCol) list.Add(TableRegion.LastColumn);
                if (inFirstRow) list.Add(TableRegion.FirstRow);
                if (inLastRow) list.Add(TableRegion.LastRow);
                if (inFirstRow && inLastCol) list.Add(TableRegion.NorthEastCell);
                if (inFirstRow && inFirstCol) list.Add(TableRegion.NorthWestCell);
                if (inLastRow && inLastCol) list.Add(TableRegion.SouthEastCell);
                if (inLastRow && inFirstCol) list.Add(TableRegion.SouthWestCell);

                list.Reverse();
                return list;
            }

            /// <summary>
            /// Nom du bord à lire pour une zone : bord extérieur de la zone (top/bottom/left/right)
            /// ou bord intérieur (insideH/insideV).
            /// </summary>
            private string EdgeKey(LayoutCell cell, TableRegion region, string edge)
            {
                int regionTop = 0, regionBottom = _layout.RowCount - 1;
                int regionLeft = 0, regionRight = _layout.ColumnCount - 1;

                switch (region)
                {
                    case TableRegion.FirstRow:
                        regionBottom = _headerRows - 1;
                        break;
                    case TableRegion.LastRow:
                        regionTop = regionBottom;
                        break;
                    case TableRegion.FirstColumn:
                        regionRight = 0;
                        break;
                    case TableRegion.LastColumn:
                        regionLeft = regionRight;
                        break;
                    case TableRegion.Band1Horizontal:
                    case TableRegion.Band2Horizontal:
                    {
                        int start = _look.FirstRow ? _headerRows : 0;
                        int band = (cell.Row - start) / _rowBand;
                        regionTop = start + band * _rowBand;
                        regionBottom = Math.Min(regionBottom, regionTop + _rowBand - 1);
                        break;
                    }
                    case TableRegion.Band1Vertical:
                    case TableRegion.Band2Vertical:
                    {
                        int start = _look.FirstColumn ? 1 : 0;
                        int band = (cell.Column - start) / _colBand;
                        regionLeft = start + band * _colBand;
                        regionRight = Math.Min(regionRight, regionLeft + _colBand - 1);
                        break;
                    }
                    case TableRegion.NorthEastCell:
                    case TableRegion.NorthWestCell:
                    case TableRegion.SouthEastCell:
                    case TableRegion.SouthWestCell:
                        regionTop = cell.Row;
                        regionBottom = cell.LastRow;
                        regionLeft = cell.Column;
                        regionRight = cell.LastColumn;
                        break;
                }

                switch (edge)
                {
                    case "top": return cell.Row <= regionTop ? "top" : "insideH";
                    case "bottom": return cell.LastRow >= regionBottom ? "bottom" : "insideH";
                    case "left": return cell.Column <= regionLeft ? "left" : "insideV";
                    default: return cell.LastColumn >= regionRight ? "right" : "insideV";
                }
            }
        }
    }
}
