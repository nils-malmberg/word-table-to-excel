using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Erreur Excel (#DIV/0!, #VALUE!…) produite par une formule.</summary>
    public sealed class ExcelErrorValue
    {
        public ExcelErrorValue(string code)
        {
            Code = code;
        }

        public readonly string Code;

        public override string ToString()
        {
            return Code;
        }
    }

    /// <summary>La formule utilise une fonction ou une syntaxe non prise en charge.</summary>
    public sealed class UnsupportedFormulaException : Exception
    {
        public UnsupportedFormulaException(string message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// Évaluateur de formules Excel simples, pour les règles de mise en forme conditionnelle :
    /// références relatives et absolues à la même feuille, opérateurs arithmétiques, de comparaison et « &amp; »,
    /// fonctions usuelles (LIGNE/ROW, MOD, ET/AND, OU/OR, SI/IF, ESTVIDE/ISBLANK, SOMME/SUM, MOYENNE/AVERAGE…).
    /// Les formules sont enregistrées en anglais dans le fichier, quelle que soit la langue d'Excel.
    /// </summary>
    public sealed class ExcelFormula
    {
        /// <summary>Accès aux valeurs des cellules de la feuille.</summary>
        public interface IContext
        {
            /// <summary>Valeur d'une cellule : double, string, bool, <see cref="ExcelErrorValue"/> ou null (vide).</summary>
            object Value(int row, int column);
            /// <summary>Nom de la feuille (les références à une autre feuille ne sont pas prises en charge).</summary>
            string SheetName { get; }
        }

        private abstract class Node
        {
        }

        private sealed class Constant : Node
        {
            public object Value;
        }

        private sealed class Reference : Node
        {
            public int Row, Column;
            public bool RowAbsolute, ColumnAbsolute;
        }

        private sealed class RangeNode : Node
        {
            public Reference First, Last;
        }

        private sealed class Unary : Node
        {
            public char Operator;
            public Node Operand;
        }

        private sealed class Binary : Node
        {
            public string Operator;
            public Node Left, Right;
        }

        private sealed class Call : Node
        {
            public string Name;
            public List<Node> Arguments = new List<Node>();
        }

        private sealed class RangeValue
        {
            public int FirstRow, FirstColumn, LastRow, LastColumn;
        }

        private readonly Node _root;

        private ExcelFormula(Node root)
        {
            _root = root;
        }

        /// <summary>Analyse la formule (sans « = » initial) ; lève <see cref="UnsupportedFormulaException"/> si elle n'est pas comprise.</summary>
        public static ExcelFormula Parse(string formula)
        {
            if (formula == null) throw new UnsupportedFormulaException("Formule vide.");
            var parser = new Parser(formula.Trim().TrimStart('='));
            var node = parser.ParseComparison();
            parser.ExpectEnd();
            return new ExcelFormula(node);
        }

        /// <summary>
        /// Évalue la formule pour la cellule (<paramref name="row"/>, <paramref name="column"/>) ; les références relatives
        /// sont décalées par rapport à la cellule d'ancrage de la règle (<paramref name="anchorRow"/>, <paramref name="anchorColumn"/>).
        /// </summary>
        public object Evaluate(IContext context, int row, int column, int anchorRow, int anchorColumn)
        {
            var evaluation = new Evaluation { Context = context, Row = row, Column = column, RowOffset = row - anchorRow, ColumnOffset = column - anchorColumn };
            var value = evaluation.Eval(_root);
            var range = value as RangeValue;
            if (range != null) value = evaluation.Context.Value(range.FirstRow, range.FirstColumn);
            return value;
        }

        /// <summary>Valeur « vraie » au sens d'une règle conditionnelle.</summary>
        public static bool IsTrue(object value)
        {
            if (value is bool) return (bool)value;
            if (value is double) return (double)value != 0;
            return false;
        }

        // ------------------------------------------------------------------ évaluation

        private sealed class Evaluation
        {
            public IContext Context;
            public int Row, Column, RowOffset, ColumnOffset;

            public object Eval(Node node)
            {
                var constant = node as Constant;
                if (constant != null) return constant.Value;
                var reference = node as Reference;
                if (reference != null)
                {
                    int r, c;
                    Resolve(reference, out r, out c);
                    if (r < 0 || c < 0 || r >= CellReference.MaxRows || c >= CellReference.MaxColumns) return new ExcelErrorValue("#REF!");
                    return Context.Value(r, c);
                }
                var range = node as RangeNode;
                if (range != null)
                {
                    int r1, c1, r2, c2;
                    Resolve(range.First, out r1, out c1);
                    Resolve(range.Last, out r2, out c2);
                    return new RangeValue { FirstRow = Math.Min(r1, r2), FirstColumn = Math.Min(c1, c2), LastRow = Math.Max(r1, r2), LastColumn = Math.Max(c1, c2) };
                }
                var unary = node as Unary;
                if (unary != null)
                {
                    var v = Scalar(Eval(unary.Operand));
                    if (v is ExcelErrorValue) return v;
                    double d;
                    if (!ToNumber(v, out d)) return new ExcelErrorValue("#VALUE!");
                    if (unary.Operator == '%') return d / 100;
                    return unary.Operator == '-' ? -d : d;
                }
                var binary = node as Binary;
                if (binary != null) return EvalBinary(binary);
                return EvalCall((Call)node);
            }

            private void Resolve(Reference reference, out int row, out int column)
            {
                row = reference.RowAbsolute ? reference.Row : reference.Row + RowOffset;
                column = reference.ColumnAbsolute ? reference.Column : reference.Column + ColumnOffset;
            }

            private object Scalar(object value)
            {
                var range = value as RangeValue;
                if (range == null) return value;
                // Intersection implicite : même ligne ou même colonne que la cellule évaluée.
                if (range.FirstColumn == range.LastColumn && Row >= range.FirstRow && Row <= range.LastRow) return Context.Value(Row, range.FirstColumn);
                if (range.FirstRow == range.LastRow && Column >= range.FirstColumn && Column <= range.LastColumn) return Context.Value(range.FirstRow, Column);
                return Context.Value(range.FirstRow, range.FirstColumn);
            }

            private object EvalBinary(Binary b)
            {
                var left = Scalar(Eval(b.Left));
                var right = Scalar(Eval(b.Right));
                if (left is ExcelErrorValue) return left;
                if (right is ExcelErrorValue) return right;
                switch (b.Operator)
                {
                    case "&":
                        return Text(left) + Text(right);
                    case "=": return Compare(left, right) == 0;
                    case "<>": return Compare(left, right) != 0;
                    case "<": return Compare(left, right) < 0;
                    case ">": return Compare(left, right) > 0;
                    case "<=": return Compare(left, right) <= 0;
                    case ">=": return Compare(left, right) >= 0;
                }
                double x, y;
                if (!ToNumber(left, out x) || !ToNumber(right, out y)) return new ExcelErrorValue("#VALUE!");
                switch (b.Operator)
                {
                    case "+": return x + y;
                    case "-": return x - y;
                    case "*": return x * y;
                    case "/": return y == 0 ? (object)new ExcelErrorValue("#DIV/0!") : x / y;
                    case "^": return Math.Pow(x, y);
                }
                throw new UnsupportedFormulaException("Opérateur " + b.Operator);
            }

            private object EvalCall(Call call)
            {
                var args = call.Arguments;
                switch (call.Name)
                {
                    case "ROW":
                        if (args.Count == 0) return (double)(Row + 1);
                        return (double)(FirstCell(args[0], true) + 1);
                    case "COLUMN":
                        if (args.Count == 0) return (double)(Column + 1);
                        return (double)(FirstCell(args[0], false) + 1);
                    case "TRUE": return true;
                    case "FALSE": return false;
                    case "NOT":
                        {
                            var v = Scalar(Eval(Arg(args, 0)));
                            if (v is ExcelErrorValue) return v;
                            return !Truthy(v);
                        }
                    case "AND":
                    case "OR":
                        {
                            bool and = call.Name == "AND";
                            bool result = and;
                            foreach (var a in args)
                            {
                                foreach (var v in Values(Eval(a)))
                                {
                                    if (v is ExcelErrorValue) return v;
                                    if (v == null || v is string) continue;
                                    bool t = Truthy(v);
                                    result = and ? result && t : result || t;
                                }
                            }
                            return result;
                        }
                    case "IF":
                        {
                            var condition = Scalar(Eval(Arg(args, 0)));
                            if (condition is ExcelErrorValue) return condition;
                            if (Truthy(condition)) return args.Count > 1 ? Scalar(Eval(args[1])) : (object)true;
                            return args.Count > 2 ? Scalar(Eval(args[2])) : (object)false;
                        }
                    case "IFERROR":
                        {
                            var v = Scalar(Eval(Arg(args, 0)));
                            return v is ExcelErrorValue ? Scalar(Eval(Arg(args, 1))) : v;
                        }
                    case "ISBLANK": return Scalar(Eval(Arg(args, 0))) == null;
                    case "ISNUMBER": return Scalar(Eval(Arg(args, 0))) is double;
                    case "ISTEXT": return Scalar(Eval(Arg(args, 0))) is string;
                    case "ISNONTEXT": return !(Scalar(Eval(Arg(args, 0))) is string);
                    case "ISLOGICAL": return Scalar(Eval(Arg(args, 0))) is bool;
                    case "ISERROR": return Scalar(Eval(Arg(args, 0))) is ExcelErrorValue;
                    case "ISERR":
                        {
                            var e = Scalar(Eval(Arg(args, 0))) as ExcelErrorValue;
                            return e != null && e.Code != "#N/A";
                        }
                    case "ISNA":
                        {
                            var e = Scalar(Eval(Arg(args, 0))) as ExcelErrorValue;
                            return e != null && e.Code == "#N/A";
                        }
                    case "ISEVEN":
                    case "ISODD":
                        {
                            double d;
                            if (!Number(args, 0, out d)) return new ExcelErrorValue("#VALUE!");
                            bool even = Math.Truncate(d) % 2 == 0;
                            return call.Name == "ISEVEN" ? even : !even;
                        }
                    case "MOD":
                        {
                            double n, d;
                            if (!Number(args, 0, out n) || !Number(args, 1, out d)) return new ExcelErrorValue("#VALUE!");
                            if (d == 0) return new ExcelErrorValue("#DIV/0!");
                            return n - d * Math.Floor(n / d);
                        }
                    case "ABS":
                    case "INT":
                    case "TRUNC":
                    case "SIGN":
                        {
                            double n;
                            if (!Number(args, 0, out n)) return new ExcelErrorValue("#VALUE!");
                            if (call.Name == "ABS") return Math.Abs(n);
                            if (call.Name == "INT") return Math.Floor(n);
                            if (call.Name == "SIGN") return (double)Math.Sign(n);
                            return Math.Truncate(n);
                        }
                    case "ROUND":
                    case "ROUNDUP":
                    case "ROUNDDOWN":
                        {
                            double n, digits = 0;
                            if (!Number(args, 0, out n) || (args.Count > 1 && !Number(args, 1, out digits))) return new ExcelErrorValue("#VALUE!");
                            double factor = Math.Pow(10, (int)digits);
                            double scaled = n * factor;
                            if (call.Name == "ROUND") scaled = Math.Round(scaled, MidpointRounding.AwayFromZero);
                            else if (call.Name == "ROUNDUP") scaled = Math.Sign(scaled) * Math.Ceiling(Math.Abs(scaled) - 1e-9);
                            else scaled = Math.Truncate(scaled);
                            return scaled / factor;
                        }
                    case "LEN": return (double)Text(Scalar(Eval(Arg(args, 0)))).Length;
                    case "UPPER": return Text(Scalar(Eval(Arg(args, 0)))).ToUpperInvariant();
                    case "LOWER": return Text(Scalar(Eval(Arg(args, 0)))).ToLowerInvariant();
                    case "TRIM": return string.Join(" ", Text(Scalar(Eval(Arg(args, 0)))).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
                    case "LEFT":
                    case "RIGHT":
                        {
                            string s = Text(Scalar(Eval(Arg(args, 0))));
                            double n = 1;
                            if (args.Count > 1 && !Number(args, 1, out n)) return new ExcelErrorValue("#VALUE!");
                            int count = Math.Max(0, Math.Min(s.Length, (int)n));
                            return call.Name == "LEFT" ? s.Substring(0, count) : s.Substring(s.Length - count);
                        }
                    case "SEARCH":
                    case "FIND":
                        {
                            string needle = Text(Scalar(Eval(Arg(args, 0))));
                            string haystack = Text(Scalar(Eval(Arg(args, 1))));
                            int index = haystack.IndexOf(needle, call.Name == "SEARCH" ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
                            return index < 0 ? (object)new ExcelErrorValue("#VALUE!") : (double)(index + 1);
                        }
                    case "EXACT": return string.Equals(Text(Scalar(Eval(Arg(args, 0)))), Text(Scalar(Eval(Arg(args, 1)))), StringComparison.Ordinal);
                    case "TODAY": return Math.Floor(ExcelDates.ToSerial(DateTime.Today, false));
                    case "NOW": return ExcelDates.ToSerial(DateTime.Now, false);
                    case "WEEKDAY":
                        {
                            double serial, type = 1;
                            if (!Number(args, 0, out serial) || (args.Count > 1 && !Number(args, 1, out type))) return new ExcelErrorValue("#VALUE!");
                            int dow = (int)((Math.Floor(serial) + 6) % 7); // 0 = dimanche
                            if ((int)type == 2) return (double)(dow == 0 ? 7 : dow);
                            if ((int)type == 3) return (double)(dow == 0 ? 6 : dow - 1);
                            return (double)(dow + 1);
                        }
                    case "SUM":
                    case "AVERAGE":
                    case "MIN":
                    case "MAX":
                    case "COUNT":
                    case "COUNTA":
                        return Aggregate(call.Name, args);
                    case "COUNTIF":
                        return CountIf(args);
                }
                throw new UnsupportedFormulaException("Fonction " + call.Name);
            }

            private static Node Arg(List<Node> args, int index)
            {
                if (index >= args.Count) throw new UnsupportedFormulaException("Argument manquant.");
                return args[index];
            }

            private bool Number(List<Node> args, int index, out double value)
            {
                var v = Scalar(Eval(Arg(args, index)));
                return ToNumber(v, out value);
            }

            private int FirstCell(Node node, bool row)
            {
                var reference = node as Reference;
                if (reference != null)
                {
                    int r, c;
                    Resolve(reference, out r, out c);
                    return row ? r : c;
                }
                var range = node as RangeNode;
                if (range != null)
                {
                    int r, c;
                    Resolve(range.First, out r, out c);
                    return row ? r : c;
                }
                throw new UnsupportedFormulaException("LIGNE/COLONNE attend une référence.");
            }

            private IEnumerable<object> Values(object value)
            {
                var range = value as RangeValue;
                if (range == null)
                {
                    yield return value;
                    yield break;
                }
                long count = (long)(range.LastRow - range.FirstRow + 1) * (range.LastColumn - range.FirstColumn + 1);
                if (count > 100000) throw new UnsupportedFormulaException("Plage trop grande.");
                for (int r = range.FirstRow; r <= range.LastRow; r++)
                {
                    for (int c = range.FirstColumn; c <= range.LastColumn; c++) yield return Context.Value(r, c);
                }
            }

            private object Aggregate(string name, List<Node> args)
            {
                var numbers = new List<double>();
                int nonEmpty = 0;
                foreach (var a in args)
                {
                    var raw = Eval(a);
                    bool isRange = raw is RangeValue;
                    foreach (var v in Values(raw))
                    {
                        if (v is ExcelErrorValue)
                        {
                            if (name == "COUNT" || name == "COUNTA")
                            {
                                if (name == "COUNTA") nonEmpty++;
                                continue;
                            }
                            return v;
                        }
                        if (v != null) nonEmpty++;
                        if (v is double) numbers.Add((double)v);
                        else if (!isRange && v is bool) numbers.Add((bool)v ? 1 : 0);
                    }
                }
                switch (name)
                {
                    case "SUM": return numbers.Sum();
                    case "AVERAGE": return numbers.Count == 0 ? (object)new ExcelErrorValue("#DIV/0!") : numbers.Average();
                    case "MIN": return numbers.Count == 0 ? 0.0 : numbers.Min();
                    case "MAX": return numbers.Count == 0 ? 0.0 : numbers.Max();
                    case "COUNT": return (double)numbers.Count;
                    default: return (double)nonEmpty;
                }
            }

            private object CountIf(List<Node> args)
            {
                var range = Eval(Arg(args, 0));
                var criterion = Scalar(Eval(Arg(args, 1)));
                string op = "=";
                object target = criterion;
                var s = criterion as string;
                if (s != null)
                {
                    foreach (var candidate in new[] { "<=", ">=", "<>", "<", ">", "=" })
                    {
                        if (s.StartsWith(candidate, StringComparison.Ordinal))
                        {
                            op = candidate;
                            s = s.Substring(candidate.Length);
                            break;
                        }
                    }
                    double d;
                    target = double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? (object)d : s;
                }
                double count = 0;
                foreach (var v in Values(range))
                {
                    if (v == null && !(target is string && ((string)target).Length == 0)) continue;
                    if ((v is double) != (target is double) && op != "<>") continue;
                    int cmp = Compare(v, target);
                    bool match = op == "=" ? cmp == 0 : op == "<>" ? cmp != 0 : op == "<" ? cmp < 0 : op == ">" ? cmp > 0 : op == "<=" ? cmp <= 0 : cmp >= 0;
                    if (match) count++;
                }
                return count;
            }
        }

        // ------------------------------------------------------------------ conversions

        private static bool Truthy(object v)
        {
            if (v is bool) return (bool)v;
            if (v is double) return (double)v != 0;
            return false;
        }

        internal static bool ToNumber(object v, out double d)
        {
            d = 0;
            if (v == null) return true;
            if (v is double)
            {
                d = (double)v;
                return true;
            }
            if (v is bool)
            {
                d = (bool)v ? 1 : 0;
                return true;
            }
            var s = v as string;
            return s != null && double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }

        internal static string Text(object v)
        {
            if (v == null) return string.Empty;
            if (v is double) return ((double)v).ToString("R", CultureInfo.InvariantCulture);
            if (v is bool) return (bool)v ? "TRUE" : "FALSE";
            return v.ToString();
        }

        /// <summary>Comparaison à la manière d'Excel : nombres &lt; textes &lt; valeurs logiques ; textes sans tenir compte de la casse.</summary>
        internal static int Compare(object a, object b)
        {
            if (a == null) a = b is string ? (object)string.Empty : b is bool ? (object)false : 0.0;
            if (b == null) b = a is string ? (object)string.Empty : a is bool ? (object)false : 0.0;
            int ra = Rank(a), rb = Rank(b);
            if (ra != rb) return ra.CompareTo(rb);
            if (a is double) return ((double)a).CompareTo((double)b);
            if (a is bool) return ((bool)a).CompareTo((bool)b);
            return string.Compare((string)a, (string)b, StringComparison.OrdinalIgnoreCase);
        }

        private static int Rank(object v)
        {
            if (v is double) return 0;
            if (v is string) return 1;
            return 2;
        }

        // ------------------------------------------------------------------ analyse syntaxique

        private sealed class Parser
        {
            private readonly string _s;
            private int _i;

            public Parser(string s)
            {
                _s = s;
            }

            public void ExpectEnd()
            {
                SkipSpaces();
                if (_i < _s.Length) throw new UnsupportedFormulaException("Syntaxe non reconnue : " + _s);
            }

            public Node ParseComparison()
            {
                var left = ParseConcat();
                while (true)
                {
                    SkipSpaces();
                    string op = null;
                    foreach (var candidate in new[] { "<>", "<=", ">=", "=", "<", ">" })
                    {
                        if (string.CompareOrdinal(_s, _i, candidate, 0, candidate.Length) == 0)
                        {
                            op = candidate;
                            break;
                        }
                    }
                    if (op == null) return left;
                    _i += op.Length;
                    left = new Binary { Operator = op, Left = left, Right = ParseConcat() };
                }
            }

            private Node ParseConcat()
            {
                var left = ParseAdditive();
                while (Accept('&')) left = new Binary { Operator = "&", Left = left, Right = ParseAdditive() };
                return left;
            }

            private Node ParseAdditive()
            {
                var left = ParseTerm();
                while (true)
                {
                    if (Accept('+')) left = new Binary { Operator = "+", Left = left, Right = ParseTerm() };
                    else if (Accept('-')) left = new Binary { Operator = "-", Left = left, Right = ParseTerm() };
                    else return left;
                }
            }

            private Node ParseTerm()
            {
                var left = ParsePower();
                while (true)
                {
                    if (Accept('*')) left = new Binary { Operator = "*", Left = left, Right = ParsePower() };
                    else if (Accept('/')) left = new Binary { Operator = "/", Left = left, Right = ParsePower() };
                    else return left;
                }
            }

            private Node ParsePower()
            {
                var left = ParseUnary();
                while (Accept('^')) left = new Binary { Operator = "^", Left = left, Right = ParseUnary() };
                return left;
            }

            private Node ParseUnary()
            {
                if (Accept('-')) return new Unary { Operator = '-', Operand = ParseUnary() };
                if (Accept('+')) return ParseUnary();
                var node = ParsePrimary();
                while (Accept('%')) node = new Unary { Operator = '%', Operand = node };
                return node;
            }

            private Node ParsePrimary()
            {
                SkipSpaces();
                if (_i >= _s.Length) throw new UnsupportedFormulaException("Formule incomplète.");
                char c = _s[_i];
                if (c == '(')
                {
                    _i++;
                    var inner = ParseComparison();
                    Expect(')');
                    return inner;
                }
                if (c == '"')
                {
                    var sb = new StringBuilder();
                    _i++;
                    while (true)
                    {
                        if (_i >= _s.Length) throw new UnsupportedFormulaException("Chaîne non terminée.");
                        if (_s[_i] == '"')
                        {
                            if (_i + 1 < _s.Length && _s[_i + 1] == '"')
                            {
                                sb.Append('"');
                                _i += 2;
                                continue;
                            }
                            _i++;
                            break;
                        }
                        sb.Append(_s[_i++]);
                    }
                    return new Constant { Value = sb.ToString() };
                }
                if (char.IsDigit(c) || c == '.')
                {
                    int start = _i;
                    while (_i < _s.Length && (char.IsDigit(_s[_i]) || _s[_i] == '.')) _i++;
                    if (_i < _s.Length && (_s[_i] == 'E' || _s[_i] == 'e'))
                    {
                        int save = _i;
                        _i++;
                        if (_i < _s.Length && (_s[_i] == '+' || _s[_i] == '-')) _i++;
                        if (_i < _s.Length && char.IsDigit(_s[_i]))
                        {
                            while (_i < _s.Length && char.IsDigit(_s[_i])) _i++;
                        }
                        else _i = save;
                    }
                    double d;
                    if (!double.TryParse(_s.Substring(start, _i - start), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) throw new UnsupportedFormulaException("Nombre invalide.");
                    return new Constant { Value = d };
                }
                if (c == '#')
                {
                    int start = _i;
                    _i++;
                    while (_i < _s.Length && (char.IsLetterOrDigit(_s[_i]) || _s[_i] == '/' || _s[_i] == '!' || _s[_i] == '?' || _s[_i] == '_')) _i++;
                    return new Constant { Value = new ExcelErrorValue(_s.Substring(start, _i - start).ToUpperInvariant()) };
                }

                string word = ReadWord();
                if (word.Length == 0) throw new UnsupportedFormulaException("Syntaxe non reconnue : " + _s);
                SkipSpaces();
                if (_i < _s.Length && _s[_i] == '(')
                {
                    _i++;
                    string name = word.ToUpperInvariant();
                    if (name.StartsWith("_XLFN.", StringComparison.Ordinal)) name = name.Substring(6);
                    var call = new Call { Name = name };
                    SkipSpaces();
                    if (!Accept(')'))
                    {
                        do
                        {
                            call.Arguments.Add(ParseComparison());
                        }
                        while (Accept(',') || Accept(';'));
                        Expect(')');
                    }
                    return call;
                }
                string upper = word.ToUpperInvariant();
                if (upper == "TRUE") return new Constant { Value = true };
                if (upper == "FALSE") return new Constant { Value = false };

                var first = ParseReference(word);
                SkipSpaces();
                if (Accept(':'))
                {
                    var last = ParseReference(ReadWord());
                    return new RangeNode { First = first, Last = last };
                }
                return first;
            }

            private string ReadWord()
            {
                SkipSpaces();
                int start = _i;
                if (_i < _s.Length && _s[_i] == '\'')
                {
                    // 'Nom de feuille'!A1
                    _i++;
                    while (_i < _s.Length)
                    {
                        if (_s[_i] == '\'')
                        {
                            if (_i + 1 < _s.Length && _s[_i + 1] == '\'')
                            {
                                _i += 2;
                                continue;
                            }
                            _i++;
                            break;
                        }
                        _i++;
                    }
                }
                while (_i < _s.Length && (char.IsLetterOrDigit(_s[_i]) || _s[_i] == '$' || _s[_i] == '_' || _s[_i] == '.' || _s[_i] == '!')) _i++;
                return _s.Substring(start, _i - start);
            }

            private static Reference ParseReference(string text)
            {
                int bang = text.LastIndexOf('!');
                if (bang >= 0) throw new UnsupportedFormulaException("Référence à une autre feuille.");
                int i = 0;
                bool colAbs = false, rowAbs = false;
                if (i < text.Length && text[i] == '$')
                {
                    colAbs = true;
                    i++;
                }
                int letters = i;
                while (i < text.Length && char.IsLetter(text[i])) i++;
                string col = text.Substring(letters, i - letters);
                if (i < text.Length && text[i] == '$')
                {
                    rowAbs = true;
                    i++;
                }
                string rowText = text.Substring(i);
                int row, column;
                if (col.Length == 0 || !CellReference.TryParse(col + rowText, out row, out column))
                {
                    throw new UnsupportedFormulaException("Nom ou référence non pris en charge : " + text);
                }
                return new Reference { Row = row, Column = column, RowAbsolute = rowAbs, ColumnAbsolute = colAbs };
            }

            private bool Accept(char c)
            {
                SkipSpaces();
                if (_i < _s.Length && _s[_i] == c)
                {
                    _i++;
                    return true;
                }
                return false;
            }

            private void Expect(char c)
            {
                if (!Accept(c)) throw new UnsupportedFormulaException("« " + c + " » attendu.");
            }

            private void SkipSpaces()
            {
                while (_i < _s.Length && char.IsWhiteSpace(_s[_i])) _i++;
            }
        }
    }
}
