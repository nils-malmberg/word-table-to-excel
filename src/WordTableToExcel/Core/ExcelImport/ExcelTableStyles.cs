using System;
using System.Collections.Generic;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Mise en forme apportée à une cellule par le style d'un tableau Excel ou par une règle de mise en forme
    /// conditionnelle (null / false = non modifiée).
    /// </summary>
    public sealed class CellFormatOverlay
    {
        public bool HasFill;
        public Rgb? Fill;
        public bool? Bold;
        public bool? Italic;
        public bool? Strike;
        public UnderlineKind? Underline;
        public bool HasFontColor;
        public Rgb? FontColor;
        public BorderLine Top, Bottom, Left, Right;
    }

    /// <summary>
    /// Styles des tableaux Excel (« Mettre sous forme de tableau ») : styles personnalisés du classeur, lus
    /// exactement, et styles intégrés (Clair 1 à 21, Moyen 1 à 28, Foncé 1 à 11), reconstitués de façon
    /// approchée à partir des couleurs du thème.
    /// </summary>
    public sealed class ExcelTableStyles
    {
        private readonly XlsxStyles _styles;
        private readonly Dictionary<string, XlsxTableStyleDefinition> _builtIn = new Dictionary<string, XlsxTableStyleDefinition>(StringComparer.OrdinalIgnoreCase);

        public ExcelTableStyles(XlsxStyles styles)
        {
            _styles = styles;
        }

        /// <summary>Mise en forme du style de <paramref name="table"/> pour la cellule (ligne, colonne de la feuille), ou null.</summary>
        public CellFormatOverlay Evaluate(XlsxTableInfo table, int row, int column)
        {
            if (table == null || !table.HasStyleInfo || !table.Range.Contains(row, column)) return null;
            var style = Definition(table.StyleName);
            if (style == null) return null;

            var range = table.Range;
            int header = Math.Min(table.HeaderRowCount, range.RowCount);
            int totals = Math.Min(table.TotalsRowCount, range.RowCount - header);
            int dataFirst = range.FirstRow + header;
            int dataLast = range.LastRow - totals;
            var overlay = new CellFormatOverlay();

            Apply(overlay, style, "wholeTable", range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, row, column);

            bool inData = row >= dataFirst && row <= dataLast;
            if (table.ShowColumnStripes && inData)
            {
                int size1 = StripeSize(style, "firstColumnStripe"), size2 = StripeSize(style, "secondColumnStripe");
                int offset = column - range.FirstColumn;
                int cycle = size1 + size2;
                int position = offset % cycle;
                bool first = position < size1;
                int bandStart = column - (first ? position : position - size1);
                int bandEnd = Math.Min(range.LastColumn, bandStart + (first ? size1 : size2) - 1);
                Apply(overlay, style, first ? "firstColumnStripe" : "secondColumnStripe", dataFirst, dataLast, bandStart, bandEnd, row, column);
            }
            if (table.ShowRowStripes && inData)
            {
                int size1 = StripeSize(style, "firstRowStripe"), size2 = StripeSize(style, "secondRowStripe");
                int offset = row - dataFirst;
                int cycle = size1 + size2;
                int position = offset % cycle;
                bool first = position < size1;
                int bandStart = row - (first ? position : position - size1);
                int bandEnd = Math.Min(dataLast, bandStart + (first ? size1 : size2) - 1);
                Apply(overlay, style, first ? "firstRowStripe" : "secondRowStripe", bandStart, bandEnd, range.FirstColumn, range.LastColumn, row, column);
            }
            if (table.ShowLastColumn) Apply(overlay, style, "lastColumn", range.FirstRow, range.LastRow, range.LastColumn, range.LastColumn, row, column);
            if (table.ShowFirstColumn) Apply(overlay, style, "firstColumn", range.FirstRow, range.LastRow, range.FirstColumn, range.FirstColumn, row, column);
            if (header > 0)
            {
                Apply(overlay, style, "headerRow", range.FirstRow, dataFirst - 1, range.FirstColumn, range.LastColumn, row, column);
                if (table.ShowFirstColumn) Apply(overlay, style, "firstHeaderCell", range.FirstRow, dataFirst - 1, range.FirstColumn, range.FirstColumn, row, column);
                if (table.ShowLastColumn) Apply(overlay, style, "lastHeaderCell", range.FirstRow, dataFirst - 1, range.LastColumn, range.LastColumn, row, column);
            }
            if (totals > 0)
            {
                Apply(overlay, style, "totalRow", dataLast + 1, range.LastRow, range.FirstColumn, range.LastColumn, row, column);
                if (table.ShowFirstColumn) Apply(overlay, style, "firstTotalCell", dataLast + 1, range.LastRow, range.FirstColumn, range.FirstColumn, row, column);
                if (table.ShowLastColumn) Apply(overlay, style, "lastTotalCell", dataLast + 1, range.LastRow, range.LastColumn, range.LastColumn, row, column);
            }
            return overlay;
        }

        private static int StripeSize(XlsxTableStyleDefinition style, string element)
        {
            int size;
            return style.StripeSizes.TryGetValue(element, out size) && size > 0 ? size : 1;
        }

        /// <summary>Applique un élément de style si la cellule se trouve dans sa zone ; les bords extérieurs et intérieurs sont distingués.</summary>
        private static void Apply(CellFormatOverlay overlay, XlsxTableStyleDefinition style, string element,
            int firstRow, int lastRow, int firstColumn, int lastColumn, int row, int column)
        {
            if (row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) return;
            XlsxDifferentialFormat f;
            if (!style.Elements.TryGetValue(element, out f) || f == null) return;
            if (f.HasFill)
            {
                overlay.HasFill = true;
                overlay.Fill = f.Fill;
            }
            if (f.Bold.HasValue) overlay.Bold = f.Bold;
            if (f.Italic.HasValue) overlay.Italic = f.Italic;
            if (f.HasFontColor)
            {
                overlay.HasFontColor = true;
                overlay.FontColor = f.FontColor;
            }
            var top = row == firstRow ? f.Top : f.InsideHorizontal;
            var bottom = row == lastRow ? f.Bottom : f.InsideHorizontal;
            var left = column == firstColumn ? f.Left : f.InsideVertical;
            var right = column == lastColumn ? f.Right : f.InsideVertical;
            if (top != null) overlay.Top = top;
            if (bottom != null) overlay.Bottom = bottom;
            if (left != null) overlay.Left = left;
            if (right != null) overlay.Right = right;
        }

        private XlsxTableStyleDefinition Definition(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            var custom = _styles.TableStyle(name);
            if (custom != null) return custom;
            XlsxTableStyleDefinition builtIn;
            if (!_builtIn.TryGetValue(name, out builtIn))
            {
                builtIn = BuildBuiltIn(name);
                _builtIn[name] = builtIn;
            }
            return builtIn;
        }

        // ------------------------------------------------------------------ styles intégrés (approximation)

        private XlsxTableStyleDefinition BuildBuiltIn(string name)
        {
            string family;
            int number;
            if (!ParseName(name, out family, out number)) return null;
            var colors = _styles.Colors;
            Rgb dark = colors.Text1;
            Rgb white = Rgb.White;
            var def = new XlsxTableStyleDefinition { Name = name };

            switch (family)
            {
                case "light":
                    {
                        if (number < 1 || number > 21) return null;
                        Rgb c = Color((number - 1) % 7, dark);
                        if (number <= 7)
                        {
                            Set(def, "wholeTable", top: Line(BorderStyle.Thin, c), bottom: Line(BorderStyle.Thin, c), fontColor: number == 1 ? (Rgb?)null : ExcelColors.ApplyTint(c, -0.25));
                            Set(def, "headerRow", bold: true, bottom: Line(BorderStyle.Thin, c));
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Thin, c));
                            Set(def, "firstRowStripe", fill: Light(c, number == 1));
                            Set(def, "firstColumnStripe", fill: Light(c, number == 1));
                        }
                        else if (number <= 14)
                        {
                            Set(def, "wholeTable", top: Line(BorderStyle.Thin, c), bottom: Line(BorderStyle.Thin, c), left: Line(BorderStyle.Thin, c), right: Line(BorderStyle.Thin, c));
                            Set(def, "headerRow", bold: true, fill: c, fontColor: white);
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, c));
                            Set(def, "firstRowStripe", top: Line(BorderStyle.Thin, c), bottom: Line(BorderStyle.Thin, c));
                            Set(def, "firstColumnStripe", left: Line(BorderStyle.Thin, c), right: Line(BorderStyle.Thin, c));
                        }
                        else
                        {
                            var line = Line(BorderStyle.Thin, c);
                            Set(def, "wholeTable", top: line, bottom: line, left: line, right: line, insideH: line, insideV: line);
                            Set(def, "headerRow", bold: true, bottom: line);
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, c));
                            Set(def, "firstRowStripe", fill: Light(c, number == 15));
                            Set(def, "firstColumnStripe", fill: Light(c, number == 15));
                        }
                        Set(def, "firstColumn", bold: true);
                        Set(def, "lastColumn", bold: true);
                        break;
                    }
                case "medium":
                    {
                        if (number < 1 || number > 28) return null;
                        Rgb c = Color((number - 1) % 7, dark);
                        bool neutral = (number - 1) % 7 == 0;
                        if (number <= 7)
                        {
                            var line = Line(BorderStyle.Thin, ExcelColors.ApplyTint(c, 0.4));
                            Set(def, "wholeTable", top: line, bottom: line, left: line, right: line, insideH: line);
                            Set(def, "headerRow", bold: true, fill: c, fontColor: white);
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, c));
                            Set(def, "firstRowStripe", fill: Light(c, neutral));
                            Set(def, "firstColumnStripe", fill: Light(c, neutral));
                            Set(def, "firstColumn", bold: true);
                            Set(def, "lastColumn", bold: true);
                        }
                        else if (number <= 14)
                        {
                            var line = Line(BorderStyle.Thin, white);
                            Set(def, "wholeTable", fill: ExcelColors.ApplyTint(c, neutral ? 0.85 : 0.8), insideH: line, insideV: line);
                            Set(def, "headerRow", bold: true, fill: c, fontColor: white, bottom: Line(BorderStyle.Thick, white));
                            Set(def, "totalRow", bold: true, fill: c, fontColor: white, top: Line(BorderStyle.Thick, white));
                            Set(def, "firstColumn", bold: true, fill: c, fontColor: white);
                            Set(def, "lastColumn", bold: true, fill: c, fontColor: white);
                            Set(def, "firstRowStripe", fill: ExcelColors.ApplyTint(c, neutral ? 0.65 : 0.6));
                            Set(def, "firstColumnStripe", fill: ExcelColors.ApplyTint(c, neutral ? 0.65 : 0.6));
                        }
                        else if (number <= 21)
                        {
                            var line = Line(BorderStyle.Thin, dark);
                            Set(def, "wholeTable", top: Line(BorderStyle.Medium, dark), bottom: Line(BorderStyle.Medium, dark), insideH: line);
                            Set(def, "headerRow", bold: true, fill: neutral ? dark : c, fontColor: white, bottom: Line(BorderStyle.Medium, dark));
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, dark));
                            Set(def, "firstRowStripe", fill: ExcelColors.ApplyTint(dark, 0.85));
                            Set(def, "firstColumnStripe", fill: ExcelColors.ApplyTint(dark, 0.85));
                            Set(def, "firstColumn", bold: true);
                            Set(def, "lastColumn", bold: true);
                        }
                        else
                        {
                            var line = Line(BorderStyle.Thin, ExcelColors.ApplyTint(c, 0.4));
                            Set(def, "wholeTable", fill: ExcelColors.ApplyTint(c, neutral ? 0.85 : 0.8), top: line, bottom: line, left: line, right: line, insideH: line, insideV: line);
                            Set(def, "headerRow", bold: true);
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, c));
                            Set(def, "firstRowStripe", fill: ExcelColors.ApplyTint(c, neutral ? 0.65 : 0.6));
                            Set(def, "firstColumnStripe", fill: ExcelColors.ApplyTint(c, neutral ? 0.65 : 0.6));
                            Set(def, "firstColumn", bold: true);
                            Set(def, "lastColumn", bold: true);
                        }
                        break;
                    }
                case "dark":
                    {
                        if (number < 1 || number > 11) return null;
                        if (number <= 7)
                        {
                            Rgb c = Color((number - 1) % 7, ExcelColors.ApplyTint(dark, 0.35));
                            Set(def, "wholeTable", fill: ExcelColors.ApplyTint(c, -0.25), fontColor: white);
                            Set(def, "headerRow", bold: true, fill: dark, fontColor: white, bottom: Line(BorderStyle.Medium, white));
                            Set(def, "totalRow", bold: true, fill: ExcelColors.ApplyTint(c, -0.5), fontColor: white, top: Line(BorderStyle.Medium, white));
                            Set(def, "firstColumn", bold: true, fill: ExcelColors.ApplyTint(c, -0.5), fontColor: white);
                            Set(def, "lastColumn", bold: true, fill: ExcelColors.ApplyTint(c, -0.5), fontColor: white);
                            Set(def, "firstRowStripe", fill: c);
                            Set(def, "firstColumnStripe", fill: c);
                        }
                        else
                        {
                            Rgb c = number == 8 ? ExcelColors.ApplyTint(dark, 0.5) : colors.Accent(number == 9 ? 1 : number == 10 ? 3 : 5);
                            Set(def, "wholeTable", fill: ExcelColors.ApplyTint(c, 0.8));
                            Set(def, "headerRow", bold: true, fill: dark, fontColor: white);
                            Set(def, "totalRow", bold: true, top: Line(BorderStyle.Double, dark));
                            Set(def, "firstRowStripe", fill: ExcelColors.ApplyTint(c, 0.6));
                            Set(def, "firstColumnStripe", fill: ExcelColors.ApplyTint(c, 0.6));
                            Set(def, "firstColumn", bold: true);
                            Set(def, "lastColumn", bold: true);
                        }
                        break;
                    }
                default:
                    return null;
            }
            return def;
        }

        /// <summary>« TableStyleMedium2 » → (medium, 2).</summary>
        internal static bool ParseName(string name, out string family, out int number)
        {
            family = null;
            number = 0;
            if (name == null || !name.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)) return false;
            string rest = name.Substring("TableStyle".Length);
            foreach (var f in new[] { "Light", "Medium", "Dark" })
            {
                if (rest.StartsWith(f, StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(rest.Substring(f.Length), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out number))
                {
                    family = f.ToLowerInvariant();
                    return true;
                }
            }
            return false;
        }

        /// <summary>Variante 0 = neutre (noir / gris), 1 à 6 = Accent 1 à 6.</summary>
        private Rgb Color(int variant, Rgb neutral)
        {
            return variant == 0 ? neutral : _styles.Colors.Accent(variant);
        }

        private static Rgb Light(Rgb c, bool neutral)
        {
            return ExcelColors.ApplyTint(c, neutral ? 0.85 : 0.8);
        }

        private static BorderLine Line(BorderStyle style, Rgb color)
        {
            return new BorderLine(style, color);
        }

        private static void Set(XlsxTableStyleDefinition def, string element, bool? bold = null, Rgb? fill = null, Rgb? fontColor = null,
            BorderLine top = null, BorderLine bottom = null, BorderLine left = null, BorderLine right = null, BorderLine insideH = null, BorderLine insideV = null)
        {
            var f = new XlsxDifferentialFormat
            {
                Bold = bold,
                HasFill = fill.HasValue,
                Fill = fill,
                HasFontColor = fontColor.HasValue,
                FontColor = fontColor,
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right,
                InsideHorizontal = insideH,
                InsideVertical = insideV
            };
            def.Elements[element] = f;
        }
    }
}
