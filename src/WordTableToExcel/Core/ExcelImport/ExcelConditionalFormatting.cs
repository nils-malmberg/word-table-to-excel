using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Mise en forme conditionnelle d'une feuille, évaluée sur les valeurs enregistrées : comparaisons de valeurs,
    /// textes, cellules vides ou en erreur, 10 premiers / derniers, au-dessus ou en dessous de la moyenne,
    /// doublons, nuances de couleurs et formules simples. Les barres de données et jeux d'icônes, qui n'ont pas
    /// d'équivalent dans un tableau Word, sont comptés comme non reproduits.
    /// </summary>
    public sealed class ExcelConditionalFormatting : ExcelFormula.IContext
    {
        private sealed class PreparedRule
        {
            public XlsxConditionalRule Rule;
            public XlsxDifferentialFormat Format;
            public List<ExcelFormula> Formulas = new List<ExcelFormula>();
            public bool Broken;
            public List<double> Numbers;
            public Dictionary<string, int> Occurrences;
            public double Average, StandardDeviation;
            public List<KeyValuePair<double, Rgb>> Scale;
        }

        private readonly XlsxSheet _sheet;
        private readonly XlsxStyles _styles;
        private readonly List<PreparedRule> _rules = new List<PreparedRule>();
        private readonly HashSet<XlsxConditionalRule> _unsupported = new HashSet<XlsxConditionalRule>();

        public ExcelConditionalFormatting(XlsxSheet sheet, XlsxStyles styles)
        {
            _sheet = sheet;
            _styles = styles;
            foreach (var rule in sheet.ConditionalRules.OrderBy(r => r.Priority))
            {
                var prepared = new PreparedRule { Rule = rule, Format = styles.DifferentialFormat(rule.DifferentialFormatId) };
                string type = (rule.Type ?? string.Empty).ToLowerInvariant();
                if (type == "databar" || type == "iconset" || type == "timeperiod")
                {
                    _unsupported.Add(rule);
                    continue;
                }
                if (type != "colorscale" && prepared.Format == null) continue; // règle sans effet visible
                try
                {
                    if (type == "cellis" || type == "expression")
                    {
                        foreach (var f in rule.Formulas) prepared.Formulas.Add(ExcelFormula.Parse(f));
                        if (prepared.Formulas.Count == 0) throw new UnsupportedFormulaException("Règle sans formule.");
                    }
                }
                catch (UnsupportedFormulaException)
                {
                    _unsupported.Add(rule);
                    continue;
                }
                _rules.Add(prepared);
            }
        }

        /// <summary>Règles qui n'ont pas pu être reproduites.</summary>
        public IEnumerable<XlsxConditionalRule> UnsupportedRules
        {
            get { return _unsupported; }
        }

        public bool IsEmpty
        {
            get { return _rules.Count == 0; }
        }

        string ExcelFormula.IContext.SheetName
        {
            get { return _sheet.Name; }
        }

        public object Value(int row, int column)
        {
            var cell = _sheet.Cell(row, column);
            if (cell == null) return null;
            switch (cell.Type)
            {
                case XlsxCellType.Number: return cell.Number;
                case XlsxCellType.Boolean: return cell.Number != 0;
                case XlsxCellType.Error: return new ExcelErrorValue(cell.Text.Text);
                case XlsxCellType.Text: return cell.Text.Text;
                default: return null;
            }
        }

        /// <summary>Mise en forme conditionnelle de la cellule, ou null si aucune règle ne s'applique.</summary>
        public CellFormatOverlay Evaluate(int row, int column)
        {
            CellFormatOverlay overlay = null;
            foreach (var rule in _rules)
            {
                if (rule.Broken || !rule.Rule.Applies(row, column)) continue;
                bool matched;
                Rgb? scaleColor = null;
                try
                {
                    matched = Matches(rule, row, column, out scaleColor);
                }
                catch (UnsupportedFormulaException)
                {
                    rule.Broken = true;
                    _unsupported.Add(rule.Rule);
                    continue;
                }
                if (!matched) continue;
                if (overlay == null) overlay = new CellFormatOverlay();
                if (scaleColor.HasValue)
                {
                    if (!overlay.HasFill)
                    {
                        overlay.HasFill = true;
                        overlay.Fill = scaleColor;
                    }
                }
                else
                {
                    Merge(overlay, rule.Format);
                }
                if (rule.Rule.StopIfTrue) break;
            }
            return overlay;
        }

        /// <summary>Complète la mise en forme sans écraser celle d'une règle plus prioritaire.</summary>
        private static void Merge(CellFormatOverlay o, XlsxDifferentialFormat f)
        {
            if (f == null) return;
            if (f.HasFill && !o.HasFill)
            {
                o.HasFill = true;
                o.Fill = f.Fill;
            }
            if (f.Bold.HasValue && !o.Bold.HasValue) o.Bold = f.Bold;
            if (f.Italic.HasValue && !o.Italic.HasValue) o.Italic = f.Italic;
            if (f.Strike.HasValue && !o.Strike.HasValue) o.Strike = f.Strike;
            if (f.Underline.HasValue && !o.Underline.HasValue) o.Underline = f.Underline;
            if (f.HasFontColor && !o.HasFontColor)
            {
                o.HasFontColor = true;
                o.FontColor = f.FontColor;
            }
            if (f.Top != null && o.Top == null) o.Top = f.Top;
            if (f.Bottom != null && o.Bottom == null) o.Bottom = f.Bottom;
            if (f.Left != null && o.Left == null) o.Left = f.Left;
            if (f.Right != null && o.Right == null) o.Right = f.Right;
        }

        private bool Matches(PreparedRule prepared, int row, int column, out Rgb? scaleColor)
        {
            scaleColor = null;
            var rule = prepared.Rule;
            object value = Value(row, column);
            string type = (rule.Type ?? string.Empty).ToLowerInvariant();
            switch (type)
            {
                case "cellis":
                    {
                        if (value is ExcelErrorValue) return false;
                        var a = prepared.Formulas[0].Evaluate(this, row, column, rule.AnchorRow, rule.AnchorColumn);
                        if (a is ExcelErrorValue) return false;
                        object b = null;
                        if (prepared.Formulas.Count > 1)
                        {
                            b = prepared.Formulas[1].Evaluate(this, row, column, rule.AnchorRow, rule.AnchorColumn);
                            if (b is ExcelErrorValue) return false;
                        }
                        switch ((rule.Operator ?? "equal").ToLowerInvariant())
                        {
                            case "lessthan": return ExcelFormula.Compare(value, a) < 0;
                            case "lessthanorequal": return ExcelFormula.Compare(value, a) <= 0;
                            case "greaterthan": return ExcelFormula.Compare(value, a) > 0;
                            case "greaterthanorequal": return ExcelFormula.Compare(value, a) >= 0;
                            case "notequal": return ExcelFormula.Compare(value, a) != 0;
                            case "between":
                            case "notbetween":
                                {
                                    if (b == null) return false;
                                    var low = ExcelFormula.Compare(a, b) <= 0 ? a : b;
                                    var high = ReferenceEquals(low, a) ? b : a;
                                    bool inside = ExcelFormula.Compare(value, low) >= 0 && ExcelFormula.Compare(value, high) <= 0;
                                    return rule.Operator.Equals("between", StringComparison.OrdinalIgnoreCase) ? inside : !inside;
                                }
                            default: return ExcelFormula.Compare(value, a) == 0;
                        }
                    }
                case "expression":
                    return ExcelFormula.IsTrue(prepared.Formulas[0].Evaluate(this, row, column, rule.AnchorRow, rule.AnchorColumn));
                case "containstext":
                case "notcontainstext":
                case "beginswith":
                case "endswith":
                    {
                        if (value is ExcelErrorValue) return false;
                        string text = ExcelFormula.Text(value);
                        string search = rule.Text ?? string.Empty;
                        bool result;
                        if (type == "beginswith") result = text.StartsWith(search, StringComparison.OrdinalIgnoreCase);
                        else if (type == "endswith") result = text.EndsWith(search, StringComparison.OrdinalIgnoreCase);
                        else result = text.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
                        return type == "notcontainstext" ? !result : result;
                    }
                case "containsblanks":
                    return value == null || (value is string && ((string)value).Trim().Length == 0);
                case "notcontainsblanks":
                    return !(value == null || (value is string && ((string)value).Trim().Length == 0));
                case "containserrors":
                    return value is ExcelErrorValue;
                case "notcontainserrors":
                    return !(value is ExcelErrorValue);
                case "top10":
                    {
                        if (!(value is double)) return false;
                        var numbers = Numbers(prepared);
                        if (numbers.Count == 0) return false;
                        int n = rule.Percent ? (int)Math.Max(1, Math.Floor(numbers.Count * rule.Rank / 100.0)) : Math.Max(1, rule.Rank);
                        n = Math.Min(n, numbers.Count);
                        double v = (double)value;
                        return rule.Bottom ? v <= numbers[n - 1] : v >= numbers[numbers.Count - n];
                    }
                case "aboveaverage":
                    {
                        if (!(value is double)) return false;
                        var numbers = Numbers(prepared);
                        if (numbers.Count == 0) return false;
                        double v = (double)value;
                        double limit = prepared.Average;
                        if (rule.StandardDeviations > 0)
                        {
                            limit += (rule.AboveAverage ? 1 : -1) * rule.StandardDeviations * prepared.StandardDeviation;
                        }
                        if (rule.AboveAverage) return rule.EqualAverage ? v >= limit : v > limit;
                        return rule.EqualAverage ? v <= limit : v < limit;
                    }
                case "duplicatevalues":
                case "uniquevalues":
                    {
                        string key = Key(value);
                        if (key == null) return false;
                        var occurrences = Occurrences(prepared);
                        int count;
                        occurrences.TryGetValue(key, out count);
                        return type == "duplicatevalues" ? count > 1 : count == 1;
                    }
                case "colorscale":
                    {
                        if (!(value is double)) return false;
                        var scale = Scale(prepared, row, column);
                        if (scale == null || scale.Count < 2) return false;
                        scaleColor = Interpolate(scale, (double)value);
                        return true;
                    }
                default:
                    _unsupported.Add(rule);
                    prepared.Broken = true;
                    return false;
            }
        }

        private static string Key(object value)
        {
            if (value == null) return null;
            if (value is double) return "n" + ((double)value).ToString("R", CultureInfo.InvariantCulture);
            if (value is bool) return "b" + value;
            if (value is string) return ((string)value).Length == 0 ? null : "t" + ((string)value).ToLowerInvariant();
            return "e" + value;
        }

        private IEnumerable<object> RangeValues(XlsxConditionalRule rule)
        {
            foreach (var range in rule.Ranges)
            {
                // Les cellules absentes sont vides : seules les cellules présentes sont parcourues.
                foreach (var cell in _sheet.Cells)
                {
                    if (range.Contains(cell.Row, cell.Column)) yield return Value(cell.Row, cell.Column);
                }
            }
        }

        private List<double> Numbers(PreparedRule prepared)
        {
            if (prepared.Numbers != null) return prepared.Numbers;
            var numbers = RangeValues(prepared.Rule).OfType<double>().ToList();
            numbers.Sort();
            prepared.Numbers = numbers;
            if (numbers.Count > 0)
            {
                prepared.Average = numbers.Average();
                if (numbers.Count > 1)
                {
                    double avg = prepared.Average;
                    prepared.StandardDeviation = Math.Sqrt(numbers.Sum(x => (x - avg) * (x - avg)) / (numbers.Count - 1));
                }
            }
            return numbers;
        }

        private Dictionary<string, int> Occurrences(PreparedRule prepared)
        {
            if (prepared.Occurrences != null) return prepared.Occurrences;
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var v in RangeValues(prepared.Rule))
            {
                string key = Key(v);
                if (key == null) continue;
                int count;
                result.TryGetValue(key, out count);
                result[key] = count + 1;
            }
            prepared.Occurrences = result;
            return result;
        }

        private List<KeyValuePair<double, Rgb>> Scale(PreparedRule prepared, int row, int column)
        {
            if (prepared.Scale != null) return prepared.Scale;
            var rule = prepared.Rule;
            var numbers = Numbers(prepared);
            var scale = new List<KeyValuePair<double, Rgb>>();
            if (numbers.Count == 0 || rule.ScaleThresholds.Count != rule.ScaleColors.Count)
            {
                prepared.Scale = scale;
                return scale;
            }
            double min = numbers[0], max = numbers[numbers.Count - 1];
            for (int i = 0; i < rule.ScaleThresholds.Count; i++)
            {
                var color = _styles.Colors.Resolve(rule.ScaleColors[i]);
                if (!color.HasValue) continue;
                string kind = (rule.ScaleThresholds[i].Key ?? "min").ToLowerInvariant();
                string text = rule.ScaleThresholds[i].Value;
                double p;
                double threshold;
                switch (kind)
                {
                    case "min": threshold = min; break;
                    case "max": threshold = max; break;
                    case "percent":
                        p = Parse(text, row, column, rule);
                        threshold = min + (max - min) * p / 100;
                        break;
                    case "percentile":
                        p = Parse(text, row, column, rule);
                        threshold = Percentile(numbers, p / 100);
                        break;
                    default:
                        threshold = Parse(text, row, column, rule);
                        break;
                }
                scale.Add(new KeyValuePair<double, Rgb>(threshold, color.Value));
            }
            prepared.Scale = scale;
            return scale;
        }

        private double Parse(string text, int row, int column, XlsxConditionalRule rule)
        {
            double d;
            if (text != null && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            var value = ExcelFormula.Parse(text).Evaluate(this, row, column, rule.AnchorRow, rule.AnchorColumn);
            if (!ExcelFormula.ToNumber(value, out d)) throw new UnsupportedFormulaException("Seuil non numérique.");
            return d;
        }

        /// <summary>Centile inclusif (CENTILE.INCLURE).</summary>
        private static double Percentile(List<double> sorted, double p)
        {
            p = Math.Max(0, Math.Min(1, p));
            double position = (sorted.Count - 1) * p;
            int lower = (int)Math.Floor(position);
            int upper = Math.Min(sorted.Count - 1, lower + 1);
            return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
        }

        private static Rgb Interpolate(List<KeyValuePair<double, Rgb>> scale, double v)
        {
            if (v <= scale[0].Key) return scale[0].Value;
            for (int i = 1; i < scale.Count; i++)
            {
                if (v <= scale[i].Key)
                {
                    double span = scale[i].Key - scale[i - 1].Key;
                    double ratio = span <= 0 ? 1 : (v - scale[i - 1].Key) / span;
                    return scale[i - 1].Value.Blend(scale[i].Value, ratio);
                }
            }
            return scale[scale.Count - 1].Value;
        }
    }
}
