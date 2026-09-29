using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Options de conversion d'une feuille Excel en tableau Word.</summary>
    public sealed class ExcelImportOptions
    {
        /// <summary>Ne pas importer les lignes et colonnes masquées (comme à l'impression).</summary>
        public bool SkipHiddenRowsAndColumns = true;
        /// <summary>Ajouter un quadrillage fin aux cellules sans bordure (automatique si la feuille imprime le quadrillage).</summary>
        public bool AddGridlines;
        /// <summary>Reconnaître une légende en tête de feuille (« Tableau 3 : … » suivi d'une ligne vide), telle que l'export l'écrit.</summary>
        public bool DetectCaption = true;
        /// <summary>Largeur disponible dans la page Word (points) ; le tableau est réduit s'il est plus large. 0 = largeurs Excel conservées.</summary>
        public double AvailableWidthPt;
        /// <summary>Fusionner une cellule de texte non renvoyé à la ligne avec les cellules vides voisines qu'il recouvre dans Excel.</summary>
        public bool MergeOverflowingText = true;
        /// <summary>Mesure de la largeur d'un texte en points (null = estimation).</summary>
        public Func<string, RunFormat, double> MeasureText;
    }

    /// <summary>Résultat de la conversion d'une feuille.</summary>
    public sealed class SheetImport
    {
        public XlsxSheetInfo Sheet;
        /// <summary>Plage importée (coordonnées de la feuille).</summary>
        public CellRange Range;
        /// <summary>Légende trouvée dans la feuille (texte complet), ou null.</summary>
        public string Caption;
        public TableModel Table;
        /// <summary>Raison pour laquelle la feuille ne peut pas être importée (null si le tableau est prêt).</summary>
        public string Error;
        public readonly List<string> Warnings = new List<string>();

        public bool Succeeded
        {
            get { return Error == null && Table != null; }
        }
    }

    /// <summary>
    /// Convertit une feuille Excel en <see cref="TableModel"/> : texte affiché exactement comme dans Excel,
    /// polices, couleurs, remplissages, bordures (partagées entre cellules voisines), fusions, alignements,
    /// retraits, orientation, largeurs de colonnes et hauteurs de lignes.
    /// </summary>
    public sealed class SheetConverter
    {
        /// <summary>Nombre maximal de colonnes d'un tableau Word.</summary>
        public const int MaxWordColumns = 63;
        /// <summary>Au-delà, l'insertion dans Word deviendrait très lente et le document inexploitable.</summary>
        public const int MaxCells = 30000;
        /// <summary>Marge intérieure gauche et droite des cellules Word créées (points).</summary>
        public const double CellMarginPt = 1.5;
        /// <summary>Couleur du quadrillage ajouté.</summary>
        public static readonly Rgb GridlineColor = new Rgb(0xA6, 0xA6, 0xA6);

        private const int MaxExtraFormattedRows = 10;
        private const int MaxExtraFormattedColumns = 5;
        private const char NoBreakSpace = '\u00A0';

        private readonly XlsxWorkbook _workbook;
        private readonly XlsxStyles _styles;
        private readonly ExcelFormatSettings _format;
        private readonly ExcelImportOptions _options;
        private readonly CaptionMatcher _captionMatcher;
        private readonly Dictionary<int, bool> _visibleFormatting = new Dictionary<int, bool>();
        private readonly ExcelTableStyles _tableStyles;
        private readonly double _maxDigitWidthPx;

        public SheetConverter(XlsxWorkbook workbook, ExcelFormatSettings format, ExcelImportOptions options, CaptionMatcher captionMatcher)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            _workbook = workbook;
            _styles = workbook.Styles;
            _format = format ?? new ExcelFormatSettings(null, null, workbook.Date1904, _styles.Colors);
            _options = options ?? new ExcelImportOptions();
            _captionMatcher = captionMatcher ?? new CaptionMatcher();
            _maxDigitWidthPx = MaxDigitWidth(_styles.DefaultFont);
            _tableStyles = new ExcelTableStyles(_styles);
        }

        // ================================================================== plage utile

        /// <summary>
        /// Plage à importer : zone d'impression si elle est définie, sinon cellules non vides, étendue aux cellules
        /// voisines mises en forme (bordures, remplissage), aux fusions et aux tableaux Excel qui les touchent.
        /// null si la feuille ne contient aucune valeur.
        /// </summary>
        public CellRange? DetectRange(XlsxSheet sheet)
        {
            CellRange? content = ContentRange(sheet);
            var printArea = sheet.Info.PrintArea;
            if (printArea.HasValue)
            {
                var area = printArea.Value;
                bool wholeRows = area.FirstColumn == 0 && area.LastColumn == CellReference.MaxColumns - 1;
                bool wholeColumns = area.FirstRow == 0 && area.LastRow == CellReference.MaxRows - 1;
                if (!wholeRows && !wholeColumns) return area;
                return content.HasValue ? content.Value.Intersect(area) : null;
            }
            return content;
        }

        private CellRange? ContentRange(XlsxSheet sheet)
        {
            int r1 = int.MaxValue, c1 = int.MaxValue, r2 = -1, c2 = -1;
            var formatted = new List<XlsxCell>();
            foreach (var cell in sheet.Cells)
            {
                if (IsNonEmpty(cell))
                {
                    r1 = Math.Min(r1, cell.Row);
                    c1 = Math.Min(c1, cell.Column);
                    r2 = Math.Max(r2, cell.Row);
                    c2 = Math.Max(c2, cell.Column);
                }
                else if (HasVisibleFormatting(cell.StyleIndex))
                {
                    formatted.Add(cell);
                }
            }
            if (r2 < 0) return null;

            var values = new CellRange(r1, c1, r2, c2);
            var box = values;
            bool changed = true;
            for (int guard = 0; changed && guard < 100; guard++)
            {
                changed = false;
                foreach (var m in sheet.MergedRanges)
                {
                    if (m.Intersects(box) && !Contains(box, m))
                    {
                        box = Union(box, m);
                        changed = true;
                    }
                }
                foreach (var t in sheet.Tables)
                {
                    if (t.Range.Intersects(box) && !Contains(box, t.Range))
                    {
                        box = Union(box, t.Range);
                        changed = true;
                    }
                }
                foreach (var cell in formatted)
                {
                    if (box.Contains(cell.Row, cell.Column)) continue;
                    bool adjacent = cell.Row >= box.FirstRow - 1 && cell.Row <= box.LastRow + 1 && cell.Column >= box.FirstColumn - 1 && cell.Column <= box.LastColumn + 1;
                    if (!adjacent) continue;
                    bool withinLimits = cell.Row >= values.FirstRow - MaxExtraFormattedRows && cell.Row <= values.LastRow + MaxExtraFormattedRows
                                     && cell.Column >= values.FirstColumn - MaxExtraFormattedColumns && cell.Column <= values.LastColumn + MaxExtraFormattedColumns;
                    if (!withinLimits) continue;
                    box = Union(box, new CellRange(cell.Row, cell.Column, cell.Row, cell.Column));
                    changed = true;
                }
            }
            return ExtendForOverflowingText(sheet, box);
        }

        /// <summary>
        /// Un texte non renvoyé à la ligne de la dernière colonne peut déborder, dans Excel, sur les colonnes vides
        /// de droite : elles font partie de ce qui est affiché (et imprimé).
        /// </summary>
        private CellRange ExtendForOverflowingText(XlsxSheet sheet, CellRange box)
        {
            int last = box.LastColumn;
            for (int r = box.FirstRow; r <= box.LastRow; r++)
            {
                XlsxCell cell = null;
                for (int c = box.LastColumn; c >= box.FirstColumn; c--)
                {
                    var candidate = sheet.Cell(r, c);
                    if (candidate != null && IsNonEmpty(candidate))
                    {
                        cell = candidate;
                        break;
                    }
                }
                if (cell == null || cell.Type != XlsxCellType.Text || sheet.MergeAt(cell.Row, cell.Column).HasValue) continue;
                var xf = _styles.CellFormat(cell.StyleIndex);
                string h = (xf.Horizontal ?? "general").ToLowerInvariant();
                if (xf.WrapText || xf.TextRotation != 0 || (h != "general" && h != "left")) continue;
                double needed = Measure(cell.Text.Text, _styles.Font(xf.FontId)) + 2 * CellMarginPt;
                double available = 0;
                for (int c = cell.Column; c <= box.LastColumn; c++) available += ColumnWidthPt(sheet, c);
                int end = box.LastColumn;
                while (available < needed && end + 1 < CellReference.MaxColumns && end - box.LastColumn < 10)
                {
                    var next = sheet.Cell(r, end + 1);
                    if (next != null && (IsNonEmpty(next) || next.HasFormula)) break;
                    end++;
                    available += ColumnWidthPt(sheet, end);
                }
                if (available >= needed) last = Math.Max(last, end);
            }
            return new CellRange(box.FirstRow, box.FirstColumn, box.LastRow, last);
        }

        private static bool IsNonEmpty(XlsxCell cell)
        {
            if (cell.FormulaWithoutValue) return true;
            if (cell.Type == XlsxCellType.Blank) return false;
            if (cell.Type == XlsxCellType.Text) return cell.Text.Text.Trim().Length > 0;
            return true;
        }

        private bool HasVisibleFormatting(int styleIndex)
        {
            bool visible;
            if (_visibleFormatting.TryGetValue(styleIndex, out visible)) return visible;
            var xf = _styles.CellFormat(styleIndex);
            var b = _styles.Border(xf.BorderId);
            visible = _styles.Fill(xf.FillId).HasValue || b.Left.IsVisible || b.Right.IsVisible || b.Top.IsVisible || b.Bottom.IsVisible;
            _visibleFormatting[styleIndex] = visible;
            return visible;
        }

        private static bool Contains(CellRange outer, CellRange inner)
        {
            return inner.FirstRow >= outer.FirstRow && inner.LastRow <= outer.LastRow && inner.FirstColumn >= outer.FirstColumn && inner.LastColumn <= outer.LastColumn;
        }

        private static CellRange Union(CellRange a, CellRange b)
        {
            return new CellRange(Math.Min(a.FirstRow, b.FirstRow), Math.Min(a.FirstColumn, b.FirstColumn), Math.Max(a.LastRow, b.LastRow), Math.Max(a.LastColumn, b.LastColumn));
        }

        /// <summary>
        /// Légende en tête de plage (disposition produite par l'export : légende en A1, ligne vide, tableau) :
        /// une seule cellule non vide dans la première ligne, qui ressemble à une légende de tableau, et une ligne vide ensuite.
        /// </summary>
        public bool TryDetectCaption(XlsxSheet sheet, CellRange range, out string caption, out CellRange tableRange)
        {
            caption = null;
            tableRange = range;
            if (range.RowCount < 3) return false;

            XlsxCell only = null;
            for (int c = range.FirstColumn; c <= range.LastColumn; c++)
            {
                var cell = sheet.Cell(range.FirstRow, c);
                if (cell == null || !IsNonEmpty(cell)) continue;
                if (only != null) return false;
                only = cell;
            }
            if (only == null || only.Type != XlsxCellType.Text) return false;
            for (int c = range.FirstColumn; c <= range.LastColumn; c++)
            {
                var cell = sheet.Cell(range.FirstRow + 1, c);
                if (cell != null && IsNonEmpty(cell)) return false;
            }
            string text = CaptionMatcher.CleanCaptionText(only.Text.Text);
            if (!_captionMatcher.LooksLikeCaptionText(text, false)) return false;

            // Le tableau commence à la première ligne non vide après la légende.
            int first = range.FirstRow + 2;
            while (first <= range.LastRow && RowIsEmpty(sheet, first, range)) first++;
            if (first > range.LastRow) return false;
            caption = text;
            tableRange = new CellRange(first, range.FirstColumn, range.LastRow, range.LastColumn);
            return true;
        }

        private static bool RowIsEmpty(XlsxSheet sheet, int row, CellRange range)
        {
            for (int c = range.FirstColumn; c <= range.LastColumn; c++)
            {
                var cell = sheet.Cell(row, c);
                if (cell != null && IsNonEmpty(cell)) return false;
            }
            return true;
        }

        private static readonly Regex CaptionPrefix = new Regex(
            @"^\s*\S+[\s\u00A0]*(?:(?:n°|no\.?|nr\.?)\s*)?[0-9A-Z]+(?:[.\-\u2013][0-9]+)*[\s\u00A0]*[:.\-\u2013\u2014][\s\u00A0]*",
            RegexOptions.CultureInvariant);

        /// <summary>« Tableau 3 : Ventes 2024 » → « Ventes 2024 » (titre à placer après le numéro créé par Word).</summary>
        public static string CaptionTitle(string caption)
        {
            if (string.IsNullOrEmpty(caption)) return string.Empty;
            var m = CaptionPrefix.Match(caption);
            return m.Success ? caption.Substring(m.Length).Trim() : caption.Trim();
        }

        // ================================================================== conversion

        private sealed class MergeInfo
        {
            public int Row, Column, RowSpan, ColumnSpan;
            /// <summary>Cellule d'Excel qui porte la valeur et la mise en forme.</summary>
            public int SourceRow, SourceColumn;
            /// <summary>Positions de feuille couvertes (visibles).</summary>
            public List<int> SheetRows, SheetColumns;
        }

        public SheetImport Convert(XlsxSheet sheet, CellRange? requestedRange)
        {
            var result = new SheetImport { Sheet = sheet.Info };
            CellRange? detected = requestedRange ?? DetectRange(sheet);
            if (!detected.HasValue)
            {
                result.Error = "La feuille est vide.";
                return result;
            }
            var range = detected.Value;
            if (!requestedRange.HasValue && _options.DetectCaption)
            {
                string caption;
                CellRange tableRange;
                if (TryDetectCaption(sheet, range, out caption, out tableRange))
                {
                    result.Caption = caption;
                    range = tableRange;
                }
            }
            result.Range = range;

            // Lignes et colonnes visibles.
            var rows = new List<int>();
            var columns = new List<int>();
            int hiddenRows = 0, hiddenColumns = 0;
            for (int r = range.FirstRow; r <= range.LastRow; r++)
            {
                if (_options.SkipHiddenRowsAndColumns && sheet.IsRowHidden(r)) hiddenRows++;
                else rows.Add(r);
                if (rows.Count > MaxCells) break;
            }
            for (int c = range.FirstColumn; c <= range.LastColumn; c++)
            {
                if (_options.SkipHiddenRowsAndColumns && sheet.IsColumnHidden(c)) hiddenColumns++;
                else columns.Add(c);
                if (columns.Count > MaxWordColumns) break;
            }
            if (rows.Count == 0 || columns.Count == 0)
            {
                result.Error = "Toutes les lignes ou toutes les colonnes de la plage " + range + " sont masquées.";
                return result;
            }
            if (columns.Count > MaxWordColumns)
            {
                result.Error = "La plage " + range + " compte plus de " + MaxWordColumns + " colonnes visibles, le maximum d'un tableau Word. "
                             + "Indiquez une plage plus étroite.";
                return result;
            }
            if ((long)rows.Count * columns.Count > MaxCells)
            {
                result.Error = "La plage " + range + " est trop grande pour un tableau Word (" + rows.Count + " lignes × " + columns.Count + " colonnes). "
                             + "Indiquez une plage plus petite (au plus " + MaxCells.ToString("N0", CultureInfo.CurrentCulture) + " cellules).";
                return result;
            }
            if (hiddenRows > 0 || hiddenColumns > 0)
            {
                result.Warnings.Add(Plural(hiddenRows, "ligne masquée ignorée", "lignes masquées ignorées") + ", "
                                  + Plural(hiddenColumns, "colonne masquée ignorée", "colonnes masquées ignorées") + ".");
            }

            var rowIndex = new Dictionary<int, int>();
            for (int i = 0; i < rows.Count; i++) rowIndex[rows[i]] = i;
            var columnIndex = new Dictionary<int, int>();
            for (int i = 0; i < columns.Count; i++) columnIndex[columns[i]] = i;

            var table = new TableModel
            {
                RowCount = rows.Count,
                ColumnCount = columns.Count,
                ColumnWidthsPt = new double[columns.Count],
                RowHeightsPt = new double[rows.Count],
                RowHeightExact = new bool[rows.Count],
                RightToLeft = sheet.RightToLeft,
                SheetName = sheet.Name
            };
            for (int i = 0; i < columns.Count; i++) table.ColumnWidthsPt[i] = ColumnWidthPt(sheet, columns[i]);
            for (int i = 0; i < rows.Count; i++)
            {
                var info = sheet.Row(rows[i]);
                table.RowHeightsPt[i] = info != null && info.Height.HasValue ? info.Height.Value : sheet.DefaultRowHeight;
            }

            // Fusions.
            var merges = new List<MergeInfo>();
            var covered = new int[rows.Count, columns.Count]; // 0 = libre, n = fusion n-1
            foreach (var m in sheet.MergedRanges)
            {
                var clipped = m.Intersect(range);
                if (!clipped.HasValue) continue;
                var mr = rows.Where(r => r >= clipped.Value.FirstRow && r <= clipped.Value.LastRow).ToList();
                var mc = columns.Where(c => c >= clipped.Value.FirstColumn && c <= clipped.Value.LastColumn).ToList();
                if (mr.Count == 0 || mc.Count == 0) continue;
                AddMerge(merges, covered, new MergeInfo
                {
                    Row = rowIndex[mr[0]],
                    Column = columnIndex[mc[0]],
                    RowSpan = mr.Count,
                    ColumnSpan = mc.Count,
                    SourceRow = m.FirstRow,
                    SourceColumn = m.FirstColumn,
                    SheetRows = mr,
                    SheetColumns = mc
                });
            }
            AddCenterAcrossSelection(sheet, rows, columns, merges, covered);
            if (_options.MergeOverflowingText) AddOverflowMerges(sheet, rows, columns, table.ColumnWidthsPt, merges, covered);

            // Cellules.
            var stats = new ConversionStats { Conditional = new ExcelConditionalFormatting(sheet, _styles) };
            var owners = new CellModel[rows.Count, columns.Count];
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < columns.Count; c++)
                {
                    int merge = covered[r, c];
                    if (merge > 0)
                    {
                        var info = merges[merge - 1];
                        if (info.Row != r || info.Column != c) continue;
                        var cell = BuildCell(sheet, info.SourceRow, info.SourceColumn, r, c, stats);
                        cell.RowSpan = info.RowSpan;
                        cell.ColumnSpan = info.ColumnSpan;
                        cell.Borders = MergedBorders(sheet, info);
                        table.Cells.Add(cell);
                        for (int i = 0; i < info.RowSpan; i++)
                        {
                            for (int j = 0; j < info.ColumnSpan; j++) owners[r + i, c + j] = cell;
                        }
                    }
                    else
                    {
                        var cell = BuildCell(sheet, rows[r], columns[c], r, c, stats);
                        table.Cells.Add(cell);
                        owners[r, c] = cell;
                    }
                }
            }

            ResolveSharedBorders(owners, rows.Count, columns.Count);
            if (_options.AddGridlines || sheet.PrintGridLines) AddGridlines(table);

            table.HeaderRowCount = HeaderRows(sheet, rows);
            FitToPage(table, stats.Values, result);
            AddWarnings(sheet, stats, result);
            result.Table = table;
            return result;
        }

        private static void AddMerge(List<MergeInfo> merges, int[,] covered, MergeInfo info)
        {
            for (int i = 0; i < info.RowSpan; i++)
            {
                for (int j = 0; j < info.ColumnSpan; j++)
                {
                    if (covered[info.Row + i, info.Column + j] != 0) return; // chevauchement : fusion ignorée
                }
            }
            merges.Add(info);
            for (int i = 0; i < info.RowSpan; i++)
            {
                for (int j = 0; j < info.ColumnSpan; j++) covered[info.Row + i, info.Column + j] = merges.Count;
            }
        }

        /// <summary>« Centré sur plusieurs colonnes » : la cellule s'étend sur les cellules vides suivantes qui ont le même alignement.</summary>
        private void AddCenterAcrossSelection(XlsxSheet sheet, List<int> rows, List<int> columns, List<MergeInfo> merges, int[,] covered)
        {
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < columns.Count; c++)
                {
                    if (covered[r, c] != 0) continue;
                    var cell = sheet.Cell(rows[r], columns[c]);
                    if (cell == null || !IsNonEmpty(cell)) continue;
                    if (!IsCenterContinuous(sheet.StyleAt(rows[r], columns[c]))) continue;
                    int end = c;
                    while (end + 1 < columns.Count && covered[r, end + 1] == 0)
                    {
                        var next = sheet.Cell(rows[r], columns[end + 1]);
                        if (next != null && IsNonEmpty(next)) break;
                        if (!IsCenterContinuous(sheet.StyleAt(rows[r], columns[end + 1]))) break;
                        end++;
                    }
                    if (end == c) continue;
                    AddMerge(merges, covered, new MergeInfo
                    {
                        Row = r,
                        Column = c,
                        RowSpan = 1,
                        ColumnSpan = end - c + 1,
                        SourceRow = rows[r],
                        SourceColumn = columns[c],
                        SheetRows = new List<int> { rows[r] },
                        SheetColumns = columns.GetRange(c, end - c + 1)
                    });
                    c = end;
                }
            }
        }

        private bool IsCenterContinuous(int styleIndex)
        {
            return string.Equals(_styles.CellFormat(styleIndex).Horizontal, "centerContinuous", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Texte non renvoyé à la ligne plus large que sa colonne : dans Excel il déborde sur les cellules vides
        /// de droite. Dans Word, la cellule est fusionnée avec ces cellules (tant que cela ne masque ni bordure ni remplissage).
        /// </summary>
        private void AddOverflowMerges(XlsxSheet sheet, List<int> rows, List<int> columns, double[] widths, List<MergeInfo> merges, int[,] covered)
        {
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < columns.Count - 1; c++)
                {
                    if (covered[r, c] != 0) continue;
                    var cell = sheet.Cell(rows[r], columns[c]);
                    if (cell == null || cell.Type != XlsxCellType.Text || !IsNonEmpty(cell)) continue;
                    var xf = _styles.CellFormat(cell.StyleIndex);
                    string h = (xf.Horizontal ?? "general").ToLowerInvariant();
                    if (xf.WrapText || xf.TextRotation != 0 || (h != "general" && h != "left")) continue;
                    var number = _format.Get(NumberFormatCode(xf.NumberFormatId));
                    if (number.HasTextSection) continue;

                    var font = _styles.Font(xf.FontId);
                    double needed = Measure(cell.Text.Text, font) + 2 * CellMarginPt + xf.Indent * IndentStepPt(font);
                    double available = widths[c];
                    if (needed <= available) continue;

                    var fill = _styles.Fill(xf.FillId);
                    var borders = _styles.Border(xf.BorderId);
                    if (borders.Right.IsVisible) continue;
                    int end = c;
                    while (available < needed && end + 1 < columns.Count && covered[r, end + 1] == 0)
                    {
                        var next = sheet.Cell(rows[r], columns[end + 1]);
                        if (next != null && (IsNonEmpty(next) || next.HasFormula)) break;
                        var nxf = _styles.CellFormat(sheet.StyleAt(rows[r], columns[end + 1]));
                        var nb = _styles.Border(nxf.BorderId);
                        if (!Nullable.Equals(_styles.Fill(nxf.FillId), fill)) break;
                        if (nb.Left.IsVisible || !Equals(nb.Top, borders.Top) || !Equals(nb.Bottom, borders.Bottom)) break;
                        end++;
                        available += widths[end];
                        if (nb.Right.IsVisible) break;
                    }
                    if (end == c) continue;
                    AddMerge(merges, covered, new MergeInfo
                    {
                        Row = r,
                        Column = c,
                        RowSpan = 1,
                        ColumnSpan = end - c + 1,
                        SourceRow = rows[r],
                        SourceColumn = columns[c],
                        SheetRows = new List<int> { rows[r] },
                        SheetColumns = columns.GetRange(c, end - c + 1)
                    });
                    c = end;
                }
            }
        }

        private sealed class ConversionStats
        {
            public ExcelConditionalFormatting Conditional;
            /// <summary>Cellules dont la valeur ne doit pas être coupée (nombres, dates, valeurs logiques, erreurs).</summary>
            public readonly HashSet<CellModel> Values = new HashSet<CellModel>();
            public int FormulasWithoutValue;
            public int Unrepresentable;
            public int SlantedText;
            public int StackedText;
        }

        private CellModel BuildCell(XlsxSheet sheet, int sheetRow, int sheetColumn, int row, int column, ConversionStats stats)
        {
            var model = new CellModel { Row = row, Column = column };
            var xlsxCell = sheet.Cell(sheetRow, sheetColumn);
            int styleIndex = xlsxCell != null ? xlsxCell.StyleIndex : sheet.StyleAt(sheetRow, sheetColumn);
            var xf = _styles.CellFormat(styleIndex);
            var font = _styles.Font(xf.FontId);
            Rgb? fill = _styles.Fill(xf.FillId);
            var borders = _styles.Border(xf.BorderId);

            // Style de tableau Excel (sous la mise en forme de la cellule), puis mise en forme conditionnelle (au-dessus).
            var table = TableOverlay(sheet, sheetRow, sheetColumn);
            if (table != null) ApplyTableOverlay(table, font, ref fill, borders);
            var conditional = stats.Conditional.IsEmpty ? null : stats.Conditional.Evaluate(sheetRow, sheetColumn);
            if (conditional != null) ApplyConditional(conditional, font, ref fill, borders);

            model.Fill = fill;
            model.Borders = borders;
            model.VerticalAlignment = MapVertical(xf.Vertical);
            model.IndentLevel = Math.Min(xf.Indent, 15);
            MapRotation(xf.TextRotation, model, stats);

            var type = xlsxCell == null ? XlsxCellType.Blank : xlsxCell.Type;
            if (xlsxCell != null && xlsxCell.FormulaWithoutValue) stats.FormulasWithoutValue++;
            model.HorizontalAlignment = MapHorizontal(xf.Horizontal, type);

            FillText(model, xlsxCell, xf, font, stats, conditional);
            if (model.Runs.Count == 0) model.Runs.Add(new TextRun(string.Empty, font, null));
            if (type == XlsxCellType.Number || type == XlsxCellType.Boolean || type == XlsxCellType.Error) stats.Values.Add(model);
            return model;
        }

        private CellFormatOverlay TableOverlay(XlsxSheet sheet, int row, int column)
        {
            foreach (var t in sheet.Tables)
            {
                if (t.Range.Contains(row, column)) return _tableStyles.Evaluate(t, row, column);
            }
            return null;
        }

        /// <summary>Le style de tableau ne s'applique que là où la cellule n'a pas sa propre mise en forme.</summary>
        private void ApplyTableOverlay(CellFormatOverlay o, RunFormat font, ref Rgb? fill, CellBorders borders)
        {
            if (o.HasFill && !fill.HasValue) fill = o.Fill;
            if (o.Bold == true) font.Bold = true;
            if (o.Italic == true) font.Italic = true;
            if (o.HasFontColor && (!font.Color.HasValue || Nullable.Equals(font.Color, _styles.DefaultFont.Color))) font.Color = o.FontColor;
            if (o.Top != null && !borders.Top.IsVisible) borders.Top = o.Top;
            if (o.Bottom != null && !borders.Bottom.IsVisible) borders.Bottom = o.Bottom;
            if (o.Left != null && !borders.Left.IsVisible) borders.Left = o.Left;
            if (o.Right != null && !borders.Right.IsVisible) borders.Right = o.Right;
        }

        /// <summary>La mise en forme conditionnelle l'emporte sur celle de la cellule.</summary>
        private static void ApplyConditional(CellFormatOverlay o, RunFormat font, ref Rgb? fill, CellBorders borders)
        {
            if (o.HasFill) fill = o.Fill;
            ApplyConditionalFont(o, font);
            if (o.Top != null) borders.Top = o.Top;
            if (o.Bottom != null) borders.Bottom = o.Bottom;
            if (o.Left != null) borders.Left = o.Left;
            if (o.Right != null) borders.Right = o.Right;
        }

        private static void ApplyConditionalFont(CellFormatOverlay o, RunFormat font)
        {
            if (o.Bold.HasValue) font.Bold = o.Bold.Value;
            if (o.Italic.HasValue) font.Italic = o.Italic.Value;
            if (o.Strike.HasValue) font.Strike = o.Strike.Value;
            if (o.Underline.HasValue) font.Underline = o.Underline.Value;
            if (o.HasFontColor) font.Color = o.FontColor;
        }

        private void FillText(CellModel model, XlsxCell cell, XlsxCellFormat xf, RunFormat font, ConversionStats stats, CellFormatOverlay conditional)
        {
            if (cell == null) return;
            switch (cell.Type)
            {
                case XlsxCellType.Number:
                    {
                        var format = _format.Get(NumberFormatCode(xf.NumberFormatId));
                        var display = format.FormatNumber(cell.Number, _format);
                        if (display.Unrepresentable) stats.Unrepresentable++;
                        var f = font.Clone();
                        if (display.Color.HasValue) f.Color = display.Color;
                        AddRun(model, KeepNumberTogether(display.Text), f);
                        return;
                    }
                case XlsxCellType.Text:
                    {
                        var format = _format.Get(NumberFormatCode(xf.NumberFormatId));
                        if (format.HasTextSection)
                        {
                            var display = format.FormatText(cell.Text.Text, _format);
                            var f = font.Clone();
                            if (display.Color.HasValue) f.Color = display.Color;
                            AddRun(model, display.Text, f);
                            return;
                        }
                        if (cell.Text.IsRich)
                        {
                            foreach (var run in cell.Text.Runs)
                            {
                                var runFont = _styles.ReadRunFont(run.Properties, font);
                                if (conditional != null) ApplyConditionalFont(conditional, runFont);
                                AddRun(model, run.Text, runFont);
                            }
                            return;
                        }
                        AddRun(model, cell.Text.Text, font);
                        return;
                    }
                case XlsxCellType.Boolean:
                    AddRun(model, ExcelLocaleTexts.Boolean(cell.Number != 0, _format.Language), font);
                    return;
                case XlsxCellType.Error:
                    AddRun(model, ExcelLocaleTexts.Error(cell.Text == null ? null : cell.Text.Text, _format.Language), font);
                    return;
            }
        }

        /// <summary>
        /// Un nombre ne doit pas être coupé en fin de ligne : « 1 234,50 € » ou « 12 % » restent d'un seul tenant
        /// (espaces insécables). Une espace entre deux chiffres (« 19/01/2024 17:30 ») et les dates en toutes lettres
        /// (« lundi 8 janvier 2024 ») peuvent, elles, passer à la ligne.
        /// </summary>
        internal static string KeepNumberTogether(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf(' ') < 0) return text;
            foreach (char c in text)
            {
                if (char.IsLetter(c) && c != 'E' && c != 'e') return text;
            }
            var chars = text.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] != ' ') continue;
                bool digitBefore = i > 0 && char.IsDigit(chars[i - 1]);
                bool digitAfter = i + 1 < chars.Length && char.IsDigit(chars[i + 1]);
                if (!(digitBefore && digitAfter)) chars[i] = NoBreakSpace;
            }
            return new string(chars);
        }

        private static void AddRun(CellModel model, string text, RunFormat format)
        {
            string clean = CleanText(text);
            if (clean.Length == 0) return;
            if (model.Runs.Count > 0 && model.Runs[model.Runs.Count - 1].Format.Equals(format))
            {
                model.Runs[model.Runs.Count - 1].Text += clean;
                return;
            }
            model.Runs.Add(new TextRun(clean, format, null));
        }

        /// <summary>Sauts de ligne normalisés (« \n »), caractères de contrôle et caractères interdits en XML supprimés.</summary>
        internal static string CleanText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r')
                {
                    sb.Append('\n');
                    if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                    continue;
                }
                if (c == '\n' || c == '\t')
                {
                    sb.Append(c);
                    continue;
                }
                if (c < 0x20 || c == '\uFFFE' || c == '\uFFFF') continue;
                if (char.IsHighSurrogate(c))
                {
                    if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                    {
                        sb.Append(c).Append(text[i + 1]);
                        i++;
                    }
                    continue;
                }
                if (char.IsLowSurrogate(c)) continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        internal string NumberFormatCode(int id)
        {
            return _styles.NumberFormatCode(id) ?? ExcelNumberFormat.BuiltInCode(id, _format.Culture);
        }

        private static HorizontalAlignment MapHorizontal(string horizontal, XlsxCellType type)
        {
            switch ((horizontal ?? "general").ToLowerInvariant())
            {
                case "left":
                case "fill":
                    return HorizontalAlignment.Left;
                case "center":
                case "centercontinuous":
                    return HorizontalAlignment.Center;
                case "right":
                    return HorizontalAlignment.Right;
                case "justify":
                    return HorizontalAlignment.Justify;
                case "distributed":
                    return HorizontalAlignment.Distributed;
                default:
                    // Alignement « Standard » d'Excel : nombres et dates à droite, valeurs logiques et erreurs centrées, texte à gauche.
                    switch (type)
                    {
                        case XlsxCellType.Number: return HorizontalAlignment.Right;
                        case XlsxCellType.Boolean:
                        case XlsxCellType.Error: return HorizontalAlignment.Center;
                        default: return HorizontalAlignment.Left;
                    }
            }
        }

        private static VerticalAlignment MapVertical(string vertical)
        {
            switch ((vertical ?? "bottom").ToLowerInvariant())
            {
                case "top":
                case "justify":
                    return VerticalAlignment.Top;
                case "center":
                case "distributed":
                    return VerticalAlignment.Center;
                default:
                    return VerticalAlignment.Bottom;
            }
        }

        private static void MapRotation(int rotation, CellModel model, ConversionStats stats)
        {
            if (rotation == 90 || rotation == 180) model.TextRotation = rotation;
            else if (rotation == 255)
            {
                model.TextRotation = 180;
                stats.StackedText++;
            }
            else if (rotation != 0)
            {
                // Angle quelconque : Word ne sait faire que 90° ; au-delà de 45°, on s'en rapproche.
                if (rotation >= 45 && rotation <= 90) model.TextRotation = 90;
                else if (rotation >= 135 && rotation <= 180) model.TextRotation = 180;
                stats.SlantedText++;
            }
        }

        /// <summary>Bordures d'une zone fusionnée : celles des cellules du pourtour.</summary>
        private CellBorders MergedBorders(XlsxSheet sheet, MergeInfo info)
        {
            var result = new CellBorders();
            int firstRow = info.SheetRows[0], lastRow = info.SheetRows[info.SheetRows.Count - 1];
            int firstColumn = info.SheetColumns[0], lastColumn = info.SheetColumns[info.SheetColumns.Count - 1];
            foreach (var c in info.SheetColumns)
            {
                result.Top = FirstVisible(result.Top, BorderAt(sheet, firstRow, c).Top);
                result.Bottom = FirstVisible(result.Bottom, BorderAt(sheet, lastRow, c).Bottom);
            }
            foreach (var r in info.SheetRows)
            {
                result.Left = FirstVisible(result.Left, BorderAt(sheet, r, firstColumn).Left);
                result.Right = FirstVisible(result.Right, BorderAt(sheet, r, lastColumn).Right);
            }
            return result;
        }

        private CellBorders BorderAt(XlsxSheet sheet, int row, int column)
        {
            return _styles.Border(_styles.CellFormat(sheet.StyleAt(row, column)).BorderId);
        }

        private static BorderLine FirstVisible(BorderLine current, BorderLine candidate)
        {
            return current.IsVisible ? current : candidate ?? BorderLine.None;
        }

        /// <summary>
        /// Excel dessine une bordure commune si l'une des deux cellules voisines la définit ; Word attend la même
        /// bordure des deux côtés. On retient la plus marquée des deux.
        /// </summary>
        private static void ResolveSharedBorders(CellModel[,] owners, int rowCount, int columnCount)
        {
            for (int pass = 0; pass < 2; pass++)
            {
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < columnCount; c++)
                    {
                        var a = owners[r, c];
                        if (c + 1 < columnCount)
                        {
                            var b = owners[r, c + 1];
                            if (!ReferenceEquals(a, b))
                            {
                                var line = Stronger(a.Borders.Right, b.Borders.Left);
                                a.Borders.Right = line;
                                b.Borders.Left = line;
                            }
                        }
                        if (r + 1 < rowCount)
                        {
                            var b = owners[r + 1, c];
                            if (!ReferenceEquals(a, b))
                            {
                                var line = Stronger(a.Borders.Bottom, b.Borders.Top);
                                a.Borders.Bottom = line;
                                b.Borders.Top = line;
                            }
                        }
                    }
                }
            }
        }

        internal static BorderLine Stronger(BorderLine a, BorderLine b)
        {
            a = a ?? BorderLine.None;
            b = b ?? BorderLine.None;
            return Weight(b.Style) > Weight(a.Style) ? b : a;
        }

        private static int Weight(BorderStyle style)
        {
            switch (style)
            {
                case BorderStyle.None: return 0;
                case BorderStyle.Hair: return 1;
                case BorderStyle.Dotted: return 2;
                case BorderStyle.DashDotDot: return 3;
                case BorderStyle.DashDot: return 4;
                case BorderStyle.Dashed: return 5;
                case BorderStyle.Thin: return 6;
                case BorderStyle.MediumDashDotDot: return 7;
                case BorderStyle.MediumDashDot: return 8;
                case BorderStyle.MediumDashed: return 9;
                case BorderStyle.Medium: return 10;
                case BorderStyle.Double: return 11;
                default: return 12; // Thick
            }
        }

        private static void AddGridlines(TableModel table)
        {
            var grid = new BorderLine(BorderStyle.Thin, GridlineColor);
            foreach (var cell in table.Cells)
            {
                var b = cell.Borders;
                if (!b.Left.IsVisible) b.Left = grid;
                if (!b.Right.IsVisible) b.Right = grid;
                if (!b.Top.IsVisible) b.Top = grid;
                if (!b.Bottom.IsVisible) b.Bottom = grid;
            }
        }

        /// <summary>Lignes d'en-tête répétées : titres d'impression, ou ligne d'en-tête d'un tableau Excel en haut de la plage.</summary>
        private static int HeaderRows(XlsxSheet sheet, List<int> rows)
        {
            int first = rows[0];
            var info = sheet.Info;
            if (info.PrintTitleFirstRow.HasValue && info.PrintTitleLastRow.HasValue
                && info.PrintTitleFirstRow.Value <= first && info.PrintTitleLastRow.Value >= first)
            {
                return rows.Count(r => r >= first && r <= info.PrintTitleLastRow.Value);
            }
            foreach (var t in sheet.Tables)
            {
                if (t.HeaderRowCount > 0 && t.Range.FirstRow == first) return 1;
            }
            return 0;
        }

        /// <summary>
        /// Largeurs minimales de chaque colonne, sans dépasser la largeur Excel :
        /// <paramref name="values"/> reçoit la largeur des valeurs à ne jamais couper (nombres, dates, valeurs logiques),
        /// la valeur renvoyée celle du plus long élément insécable, mots compris.
        /// </summary>
        private double[] MinimumWidths(TableModel table, HashSet<CellModel> valueCells, out double[] values)
        {
            var min = new double[table.ColumnCount];
            values = new double[table.ColumnCount];
            foreach (var cell in table.Cells)
            {
                if (cell.ColumnSpan != 1 || cell.TextRotation != 0) continue;
                bool isValue = valueCells.Contains(cell);
                double longest = 0;
                foreach (var run in cell.Runs)
                {
                    // Les espaces insécables des nombres (« 1 234,50 € ») ne sont pas des points de coupure.
                    var pieces = isValue ? run.Text.Split(new[] { ' ', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                         : run.Text.Split(new[] { ' ', '\n', '\t', '-', '\u2013', '/' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var piece in pieces) longest = Math.Max(longest, Measure(piece, run.Format));
                }
                if (longest <= 0) continue;
                // Une valeur affichée dans Excel tient dans sa colonne : on ne demande pas plus (petite marge pour l'estimation).
                double cap = table.ColumnWidthsPt[cell.Column] * (isValue ? 1.08 : 1.0);
                double need = Math.Min(cap, longest + 2 * CellMarginPt + cell.IndentLevel * IndentStepPt(cell.PrimaryFormat) + 1);
                min[cell.Column] = Math.Max(min[cell.Column], need);
                if (isValue) values[cell.Column] = Math.Max(values[cell.Column], need);
            }
            return min;
        }

        /// <summary>
        /// Réduit un tableau plus large que la page, par étapes : d'abord les colonnes qui ont de la marge (sans
        /// descendre sous le plus long mot ou la plus longue valeur), puis les colonnes de texte (dont les mots
        /// passent à la ligne), et seulement en dernier recours les nombres.
        /// </summary>
        private void FitToPage(TableModel table, HashSet<CellModel> valueCells, SheetImport result)
        {
            var widths = table.ColumnWidthsPt;
            double[] values;
            var min = MinimumWidths(table, valueCells, out values);
            // Une colonne à peine assez large dans Excel est légèrement élargie pour que ses valeurs ne soient pas coupées.
            for (int i = 0; i < widths.Length; i++) widths[i] = Math.Max(widths[i], values[i]);

            double total = widths.Sum();
            double available = _options.AvailableWidthPt;
            if (available <= 0 || total <= available + 0.5) return;

            // Plancher des colonnes de texte : quelques caractères.
            var floor = new double[widths.Length];
            for (int i = 0; i < widths.Length; i++)
            {
                double fewCharacters = 4 * 0.55 * (_styles.DefaultFont.Size ?? 11) + 2 * CellMarginPt;
                floor[i] = Math.Max(values[i], Math.Min(min[i], fewCharacters));
            }

            if (!ShrinkTowards(widths, min, available) && !ShrinkTowards(widths, floor, available))
            {
                double sum = widths.Sum();
                for (int i = 0; i < widths.Length; i++) widths[i] *= available / sum;
                result.Warnings.Add("Le tableau reste plus large que la page même réduit au minimum : certaines valeurs passent à la ligne. "
                                  + "Pensez à une page en orientation Paysage ou à une plage plus étroite.");
            }
            result.Warnings.Add("Largeur du tableau réduite à " + Math.Floor(available / total * 100).ToString(CultureInfo.CurrentCulture) + " % pour tenir dans la page.");
        }

        /// <summary>Réduit les largeurs vers <paramref name="lower"/> pour atteindre <paramref name="available"/> ; false si c'est impossible.</summary>
        private static bool ShrinkTowards(double[] widths, double[] lower, double available)
        {
            double total = widths.Sum();
            double excess = total - available;
            if (excess <= 0) return true;
            double slack = 0;
            for (int i = 0; i < widths.Length; i++) slack += Math.Max(0, widths[i] - lower[i]);
            if (slack < excess)
            {
                for (int i = 0; i < widths.Length; i++) widths[i] = Math.Min(widths[i], Math.Max(lower[i], 1));
                return false;
            }
            double ratio = excess / slack;
            for (int i = 0; i < widths.Length; i++)
            {
                if (widths[i] > lower[i]) widths[i] -= (widths[i] - lower[i]) * ratio;
            }
            return true;
        }

        private void AddWarnings(XlsxSheet sheet, ConversionStats stats, SheetImport result)
        {
            if (stats.FormulasWithoutValue > 0)
            {
                result.Warnings.Add(Plural(stats.FormulasWithoutValue, "formule n'a", "formules n'ont") + " pas de résultat enregistré dans le fichier (cellules laissées vides). "
                                  + "Ouvrez le classeur dans Excel, enregistrez-le, puis recommencez.");
            }
            else if (_workbook.FullCalculationOnLoad && sheet.Cells.Any(c => c.HasFormula))
            {
                result.Warnings.Add("Le classeur demande un recalcul à l'ouverture : les résultats de formules enregistrés peuvent être périmés.");
            }
            if (stats.Unrepresentable > 0)
            {
                result.Warnings.Add(Plural(stats.Unrepresentable, "valeur ne peut", "valeurs ne peuvent") + " pas être affichée(s) avec son format (date négative ou hors limites : « ##### » dans Excel) ; nombre écrit au format Standard.");
            }
            if (stats.SlantedText > 0) result.Warnings.Add("Texte incliné (" + stats.SlantedText + " cellule(s)) : Word ne permet que l'orientation verticale.");
            if (stats.StackedText > 0) result.Warnings.Add("Texte vertical empilé (" + stats.StackedText + " cellule(s)) remplacé par un texte pivoté.");
            int unsupported = stats.Conditional.UnsupportedRules.Count(r => r.Ranges.Any(x => x.Intersects(result.Range)));
            if (unsupported > 0)
            {
                result.Warnings.Add(Plural(unsupported, "règle de mise en forme conditionnelle n'a", "règles de mise en forme conditionnelle n'ont")
                                  + " pas pu être reproduite(s) (barres de données, jeux d'icônes ou formules non prises en charge).");
            }
            if (sheet.HasDrawings) result.Warnings.Add("Les images, graphiques et formes de la feuille ne sont pas importés.");
            if (sheet.HasSparklines) result.Warnings.Add("Les graphiques sparkline ne sont pas importés.");
            if (sheet.Truncated) result.Warnings.Add("La feuille est très volumineuse : seule une partie a été lue.");
            if (sheet.Info.PrintAreaHasSeveralRanges) result.Warnings.Add("La zone d'impression comporte plusieurs plages : seule la première est importée.");
        }

        private static string Plural(int count, string singular, string plural)
        {
            return count.ToString(CultureInfo.CurrentCulture) + " " + (count > 1 ? plural : singular);
        }

        // ================================================================== dimensions

        /// <summary>Largeur d'une colonne en points, selon les règles d'Excel (police Normal du classeur).</summary>
        public double ColumnWidthPt(XlsxSheet sheet, int column)
        {
            var info = sheet.Column(column);
            double mdw = _maxDigitWidthPx;
            double pixels;
            if (info != null && info.Width.HasValue)
            {
                double w = info.Width.Value;
                pixels = Math.Truncate(((256 * w + Math.Truncate(128 / mdw)) / 256) * mdw);
            }
            else if (sheet.DefaultColumnWidth.HasValue)
            {
                double w = sheet.DefaultColumnWidth.Value;
                pixels = Math.Truncate(((256 * w + Math.Truncate(128 / mdw)) / 256) * mdw);
            }
            else
            {
                // Largeur standard : baseColWidth caractères + 5 pixels, arrondie au multiple de 8 pixels supérieur.
                pixels = Math.Ceiling((sheet.BaseColumnWidth * mdw + 5) / 8) * 8;
            }
            return Math.Max(pixels, 1) * 0.75;
        }

        /// <summary>Largeur (pixels à 96 ppp) du chiffre le plus large de la police Normal : unité des largeurs de colonnes d'Excel.</summary>
        internal static double MaxDigitWidth(RunFormat font)
        {
            double size = font != null && font.Size.HasValue ? font.Size.Value : 11;
            string name = font != null && font.FontName != null ? font.FontName.ToLowerInvariant() : "calibri";
            double ratio;
            if (name.StartsWith("calibri", StringComparison.Ordinal)) ratio = 0.49;
            else if (name.StartsWith("aptos narrow", StringComparison.Ordinal) || name.Contains("narrow") || name.Contains("condensed")) ratio = 0.47;
            else if (name.StartsWith("aptos", StringComparison.Ordinal)) ratio = 0.54;
            else if (name.StartsWith("arial", StringComparison.Ordinal) || name.StartsWith("helvetica", StringComparison.Ordinal) || name.StartsWith("liberation sans", StringComparison.Ordinal)) ratio = 0.556;
            else if (name.StartsWith("times", StringComparison.Ordinal) || name.StartsWith("liberation serif", StringComparison.Ordinal)) ratio = 0.5;
            else if (name.StartsWith("verdana", StringComparison.Ordinal)) ratio = 0.636;
            else if (name.StartsWith("courier", StringComparison.Ordinal) || name.StartsWith("consolas", StringComparison.Ordinal)) ratio = 0.6;
            else ratio = 0.53;
            return Math.Max(1, Math.Round(size * 96.0 / 72.0 * ratio));
        }

        private double Measure(string text, RunFormat font)
        {
            if (_options.MeasureText != null)
            {
                try
                {
                    return _options.MeasureText(text, font);
                }
                catch (Exception)
                {
                    // Mesure impossible : estimation.
                }
            }
            return EstimateTextWidth(text, font);
        }

        /// <summary>Estimation de la largeur d'un texte (points), à défaut de pouvoir le mesurer.</summary>
        internal static double EstimateTextWidth(string text, RunFormat font)
        {
            double size = font != null && font.Size.HasValue ? font.Size.Value : 11;
            double widest = 0;
            foreach (var line in (text ?? string.Empty).Split('\n'))
            {
                double em = 0;
                foreach (char c in line)
                {
                    if ("il.,:;'|!ftjrI ".IndexOf(c) >= 0) em += 0.3;
                    else if ("mwMW@%".IndexOf(c) >= 0) em += 0.85;
                    else if (char.IsUpper(c)) em += 0.65;
                    else if (char.IsDigit(c)) em += 0.55;
                    else if (c > 0x2E80) em += 1.0; // idéogrammes
                    else em += 0.5;
                }
                widest = Math.Max(widest, em * size);
            }
            if (font != null && font.Bold) widest *= 1.07;
            return widest;
        }

        /// <summary>Largeur d'un niveau de retrait Excel (environ 3 caractères de la police).</summary>
        internal static double IndentStepPt(RunFormat font)
        {
            double size = font != null && font.Size.HasValue ? font.Size.Value : 11;
            return Math.Round(size * 0.82, 2);
        }
    }
}
