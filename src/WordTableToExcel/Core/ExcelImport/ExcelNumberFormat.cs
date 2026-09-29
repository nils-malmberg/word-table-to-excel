using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Paramètres régionaux utilisés pour reproduire l'affichage d'Excel.</summary>
    public sealed class ExcelFormatSettings
    {
        private readonly Dictionary<string, ExcelNumberFormat> _cache = new Dictionary<string, ExcelNumberFormat>(StringComparer.Ordinal);

        public ExcelFormatSettings(CultureInfo culture, CultureInfo language, bool date1904, ExcelColors colors)
        {
            Culture = culture ?? CultureInfo.CurrentCulture;
            Language = language ?? CultureInfo.CurrentUICulture;
            Date1904 = date1904;
            Colors = colors ?? ExcelColors.Default;
        }

        /// <summary>Paramètres régionaux de Windows : séparateurs décimal et de milliers, noms des mois et des jours, date courte.</summary>
        public CultureInfo Culture { get; private set; }
        /// <summary>Langue d'Office : textes VRAI/FAUX et codes d'erreur (#VALEUR!…).</summary>
        public CultureInfo Language { get; private set; }
        public bool Date1904 { get; private set; }
        public ExcelColors Colors { get; private set; }

        /// <summary>Format analysé (mis en cache : un classeur réutilise quelques formats pour des milliers de cellules).</summary>
        public ExcelNumberFormat Get(string code)
        {
            code = code ?? "General";
            ExcelNumberFormat format;
            if (!_cache.TryGetValue(code, out format))
            {
                format = ExcelNumberFormat.Parse(code, Culture);
                _cache[code] = format;
            }
            return format;
        }
    }

    /// <summary>Texte affiché par Excel pour une valeur, avec la couleur imposée par le format ([Rouge]…).</summary>
    public struct ExcelDisplayValue
    {
        public string Text;
        public Rgb? Color;
        /// <summary>
        /// La valeur ne peut pas être affichée par ce format (date négative ou trop grande, aucune section
        /// applicable) : Excel afficherait « ##### » ; le nombre est alors écrit au format Standard.
        /// </summary>
        public bool Unrepresentable;
    }

    /// <summary>
    /// Format de nombre Excel (« # ##0,00 € », « jj/mm/aaaa », « 0,0% », « [Rouge]-0,00 »…) :
    /// analyse du code et production du texte exactement tel qu'Excel l'affiche.
    /// Prend en charge les sections (positif ; négatif ; zéro ; texte), les conditions, les couleurs,
    /// les séparateurs de milliers et divisions par mille, pourcentages, notation scientifique, fractions,
    /// dates et heures (y compris durées [h]:mm), balises régionales [$€-40C] et le format Standard.
    /// </summary>
    public sealed class ExcelNumberFormat
    {
        internal const char Pad = '\uE000';

        private readonly List<Section> _sections;

        private ExcelNumberFormat(string code, List<Section> sections)
        {
            Code = code;
            _sections = sections;
        }

        public string Code { get; private set; }

        /// <summary>true si la première section affiche une date ou une heure.</summary>
        public bool IsDateTime
        {
            get { return _sections.Count > 0 && _sections[0].IsDate; }
        }

        /// <summary>true si le format a une section texte (4e section, ou dernière section contenant @) qui modifie l'affichage des textes.</summary>
        public bool HasTextSection
        {
            get
            {
                Section text;
                NumericSections(out text);
                return text != null && !(text.HasText && text.IsOnlyTextPlaceholder);
            }
        }

        /// <summary>true si le format est « Texte » (@).</summary>
        public bool IsTextFormat
        {
            get { return _sections.Count == 1 && _sections[0].HasText && !_sections[0].HasDigits; }
        }

        public static ExcelNumberFormat Parse(string code)
        {
            return Parse(code, CultureInfo.CurrentCulture);
        }

        /// <param name="culture">Paramètres régionaux des formats « système » ([$-F800] date longue, [$-F400] heure).</param>
        public static ExcelNumberFormat Parse(string code, CultureInfo culture)
        {
            if (code == null) code = "General";
            culture = culture ?? CultureInfo.CurrentCulture;
            var sections = new List<Section>();
            foreach (var part in SplitSections(code))
            {
                sections.Add(Section.Parse(part, culture));
                if (sections.Count == 4) break;
            }
            if (sections.Count == 0) sections.Add(Section.Parse("General", culture));
            return new ExcelNumberFormat(code, sections);
        }

        // ------------------------------------------------------------------ valeurs

        public ExcelDisplayValue FormatNumber(double value, ExcelFormatSettings settings)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return new ExcelDisplayValue { Text = "#NUM!" };

            Section textSection;
            List<Section> numeric = NumericSections(out textSection);
            Section section;
            bool absolute;
            if (!ChooseSection(numeric, value, out section, out absolute))
            {
                return new ExcelDisplayValue { Text = General(value, settings.Culture), Unrepresentable = true };
            }
            if (section == null)
            {
                // Seule une section texte (format « @ ») : le nombre s'affiche au format Standard.
                return new ExcelDisplayValue { Text = General(value, settings.Culture) };
            }

            var result = new ExcelDisplayValue { Color = section.ResolveColor(settings.Colors) };
            bool negative = value < 0 && !absolute;
            string text;
            if (section.IsDate)
            {
                if (!section.FormatDate(value, settings, out text))
                {
                    return new ExcelDisplayValue { Text = General(value, settings.Culture), Unrepresentable = true, Color = result.Color };
                }
            }
            else
            {
                text = section.FormatNumber(Math.Abs(value), negative, settings);
            }
            result.Text = FinishNumber(text);
            return result;
        }

        public ExcelDisplayValue FormatText(string text, ExcelFormatSettings settings)
        {
            Section textSection;
            NumericSections(out textSection);
            if (textSection == null) return new ExcelDisplayValue { Text = text ?? string.Empty };
            return new ExcelDisplayValue
            {
                Text = textSection.FormatText(text ?? string.Empty, settings),
                Color = textSection.ResolveColor(settings.Colors)
            };
        }

        private List<Section> NumericSections(out Section textSection)
        {
            textSection = null;
            var numeric = new List<Section>(_sections);
            if (numeric.Count == 4)
            {
                textSection = numeric[3];
                numeric.RemoveAt(3);
            }
            else if (numeric[numeric.Count - 1].HasText)
            {
                textSection = numeric[numeric.Count - 1];
                numeric.RemoveAt(numeric.Count - 1);
            }
            return numeric;
        }

        /// <summary>
        /// Choisit la section d'un nombre. <paramref name="absolute"/> = la section affiche la valeur absolue
        /// (section « négatif » : son signe est porté par le format, par ex. des parenthèses).
        /// </summary>
        private static bool ChooseSection(List<Section> numeric, double value, out Section section, out bool absolute)
        {
            section = null;
            absolute = false;
            if (numeric.Count == 0) return true;

            bool conditional = false;
            foreach (var s in numeric)
            {
                if (s.Condition != null) conditional = true;
            }

            if (!conditional)
            {
                switch (numeric.Count)
                {
                    case 1:
                        section = numeric[0];
                        return true;
                    case 2:
                        section = value >= 0 ? numeric[0] : numeric[1];
                        absolute = value < 0;
                        return true;
                    default:
                        section = value > 0 ? numeric[0] : value < 0 ? numeric[1] : numeric[2];
                        absolute = value < 0;
                        return true;
                }
            }

            // Formats conditionnels : [cond1]…;[cond2]…;autre.
            var first = numeric[0];
            var second = numeric.Count > 1 ? numeric[1] : null;
            if (first.Condition != null && first.Condition.Matches(value)) section = first;
            else if (second != null && second.Condition != null && second.Condition.Matches(value)) section = second;
            else if (first.Condition != null && second != null && second.Condition != null)
            {
                section = numeric.Count > 2 ? numeric[2] : null;
            }
            else if (first.Condition == null)
            {
                section = first;
            }
            else
            {
                section = second;
            }
            if (section == null) return false;
            // Section « sinon » (sans condition) : valeur absolue ; section conditionnelle : signe affiché,
            // sauf si sa condition ne retient que des nombres négatifs ([<0]).
            absolute = value < 0 && (section.Condition == null || section.Condition.ExcludesNonNegative);
            return true;
        }

        /// <summary>Supprime les espacements de début et de fin (caractères « _x », « ? » inutilisés) et les espaces superflus.</summary>
        private static string FinishNumber(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return StripPadding(text).Trim(' ');
        }

        internal static string StripPadding(string text)
        {
            if (text.IndexOf(Pad) < 0) return text;
            int start = 0, end = text.Length;
            while (start < end && text[start] == Pad) start++;
            while (end > start && text[end - 1] == Pad) end--;
            // Les espacements d'alignement d'Excel n'ont pas d'équivalent dans Word : une série devient une espace.
            var sb = new StringBuilder(end - start);
            for (int i = start; i < end; i++)
            {
                if (text[i] == Pad)
                {
                    if (sb.Length == 0 || sb[sb.Length - 1] != ' ') sb.Append(' ');
                    continue;
                }
                sb.Append(text[i]);
            }
            return sb.ToString();
        }

        // ------------------------------------------------------------------ format Standard

        /// <summary>
        /// Format « Standard » d'Excel : au plus 11 caractères (hors signe), notation scientifique pour les très
        /// grands et très petits nombres, zéros non significatifs supprimés.
        /// </summary>
        public static string General(double value, CultureInfo culture)
        {
            if (value == 0 || double.IsNaN(value) || double.IsInfinity(value)) return "0";
            var d = ExcelDecimal.FromDouble(value);
            string body = GeneralBody(d, culture.NumberFormat.NumberDecimalSeparator);
            return (value < 0 && body != "0" ? culture.NumberFormat.NegativeSign : string.Empty) + body;
        }

        private static string GeneralBody(ExcelDecimal d, string decimalSeparator)
        {
            if (d.IsZero) return "0";
            int v = d.Log10Floor;
            string s;
            if (v >= -4 && v <= -1)
            {
                s = Fixed(d.RoundSignificant(10 + v), decimalSeparator);
            }
            else if (v >= 0 && v <= 9)
            {
                s = Fixed(d.RoundSignificant(10), decimalSeparator);
                if (Length(s) > 11) s = Exponential(d, 5, decimalSeparator);
            }
            else if (v >= -9 && v <= -5)
            {
                s = Fixed(d.RoundDecimals(12), decimalSeparator);
                if (Length(s) > 11) s = Exponential(d, 5, decimalSeparator);
            }
            else if (v == 10)
            {
                var rounded = d.RoundDecimals(0);
                s = rounded.Log10Floor > 10 ? Exponential(d, 5, decimalSeparator) : Fixed(rounded, decimalSeparator);
            }
            else
            {
                s = Exponential(d, 5, decimalSeparator);
            }
            return s;
        }

        private static int Length(string s)
        {
            return s.Length;
        }

        /// <summary>Écriture fixe sans zéros inutiles (« 1234.5 », « 0.001 »).</summary>
        private static string Fixed(ExcelDecimal d, string decimalSeparator)
        {
            if (d.IsZero) return "0";
            string integer = d.IntegerDigits();
            string fraction = d.FractionDigits(d.DecimalCount);
            return (integer.Length == 0 ? "0" : integer) + (fraction.Length > 0 ? decimalSeparator + fraction : string.Empty);
        }

        /// <summary>Notation scientifique « 1.23457E+11 » (mantisse sans zéros de fin, exposant sur 2 chiffres au moins).</summary>
        private static string Exponential(ExcelDecimal d, int decimals, string decimalSeparator)
        {
            int exponent = d.Log10Floor;
            var mantissa = d.Shift(-exponent).RoundDecimals(decimals);
            if (mantissa.Log10Floor >= 1)
            {
                exponent++;
                mantissa = d.Shift(-exponent).RoundDecimals(decimals);
            }
            string m = Fixed(mantissa, decimalSeparator);
            return m + "E" + (exponent < 0 ? "-" : "+") + Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture);
        }

        // ------------------------------------------------------------------ formats intégrés

        /// <summary>
        /// Code des formats intégrés d'Excel (numFmtId 0 à 81), adaptés aux paramètres régionaux pour ceux
        /// qui en dépendent (date courte, monnaie). Renvoie « General » pour un identifiant inconnu.
        /// </summary>
        public static string BuiltInCode(int id, CultureInfo culture)
        {
            culture = culture ?? CultureInfo.CurrentCulture;
            switch (id)
            {
                case 0: return "General";
                case 1: case 59: return "0";
                case 2: case 60: return "0.00";
                case 3: case 61: return "#,##0";
                case 4: case 62: return "#,##0.00";
                case 5: return CurrencyCode(culture, 0, false);
                case 6: return CurrencyCode(culture, 0, true);
                case 7: return CurrencyCode(culture, 2, false);
                case 8: return CurrencyCode(culture, 2, true);
                case 9: case 67: return "0%";
                case 10: case 68: return "0.00%";
                case 11: return "0.00E+00";
                case 12: case 69: return "# ?/?";
                case 13: case 70: return "# ??/??";
                case 14: case 71: case 72:
                case 27: case 28: case 29: case 30: case 31: case 32: case 33: case 34: case 35: case 36:
                case 50: case 51: case 52: case 53: case 54: case 55: case 56: case 57: case 58:
                    return ShortDateCode(culture);
                case 15: case 73: return "d-mmm-yy";
                case 16: case 74: return "d-mmm";
                case 17: case 75: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: case 76: return "h:mm";
                case 21: case 77: return "h:mm:ss";
                case 22: case 78: return ShortDateCode(culture) + " " + ShortTimeCode(culture);
                case 37: return "#,##0 ;(#,##0)";
                case 38: return "#,##0 ;[Red](#,##0)";
                case 39: return "#,##0.00;(#,##0.00)";
                case 40: return "#,##0.00;[Red](#,##0.00)";
                case 41: return "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)";
                case 42: return "_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* #,##0_);_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* \\(#,##0\\);_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* \"-\"_);_(@_)";
                case 43: return "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
                case 44: return "_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* #,##0.00_);_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* \\(#,##0.00\\);_(" + Quote(culture.NumberFormat.CurrencySymbol) + "* \"-\"??_);_(@_)";
                case 45: case 79: return "mm:ss";
                case 46: case 80: return "[h]:mm:ss";
                case 47: case 81: return "mm:ss.0";
                case 48: return "##0.0E+0";
                case 49: return "@";
                default: return "General";
            }
        }

        private static string CurrencyCode(CultureInfo culture, int decimals, bool red)
        {
            var nfi = culture.NumberFormat;
            string number = decimals > 0 ? "#,##0." + new string('0', decimals) : "#,##0";
            string symbol = Quote(nfi.CurrencySymbol);
            string positive;
            switch (nfi.CurrencyPositivePattern)
            {
                case 0: positive = symbol + number; break;
                case 1: positive = number + symbol; break;
                case 2: positive = symbol + "\\ " + number; break;
                default: positive = number + "\\ " + symbol; break;
            }
            return positive + ";" + (red ? "[Red]" : string.Empty) + "\\-" + positive;
        }

        private static string Quote(string literal)
        {
            return "\"" + (literal ?? string.Empty).Replace("\"", string.Empty) + "\"";
        }

        internal static string ShortDateCode(CultureInfo culture)
        {
            return DotNetPatternToExcel(culture.DateTimeFormat.ShortDatePattern, culture);
        }

        private static string ShortTimeCode(CultureInfo culture)
        {
            string pattern = culture.DateTimeFormat.ShortTimePattern ?? "H:mm";
            return pattern.Contains("HH") ? "hh:mm" : "h:mm";
        }

        /// <summary>Convertit un modèle de date .NET (« dddd d MMMM yyyy ») en code de format Excel.</summary>
        internal static string DotNetPatternToExcel(string pattern, CultureInfo culture)
        {
            if (string.IsNullOrEmpty(pattern)) return "General";
            var sb = new StringBuilder();
            var literal = new StringBuilder();
            Action flush = () =>
            {
                if (literal.Length == 0) return;
                sb.Append('"').Append(literal.ToString().Replace("\"", string.Empty)).Append('"');
                literal.Length = 0;
            };
            bool twelveHour = pattern.IndexOf('t') >= 0 && pattern.IndexOf('h') >= 0;
            int i = 0;
            while (i < pattern.Length)
            {
                char c = pattern[i];
                int run = 1;
                while (i + run < pattern.Length && pattern[i + run] == c) run++;
                switch (c)
                {
                    case 'd':
                        flush();
                        sb.Append(new string('d', Math.Min(run, 4)));
                        break;
                    case 'M':
                        flush();
                        sb.Append(new string('m', Math.Min(run, 4)));
                        break;
                    case 'y':
                        flush();
                        sb.Append(run <= 2 ? "yy" : "yyyy");
                        break;
                    case 'H':
                    case 'h':
                        flush();
                        sb.Append(run >= 2 ? "hh" : "h");
                        break;
                    case 'm':
                        flush();
                        sb.Append(run >= 2 ? "mm" : "m");
                        break;
                    case 's':
                        flush();
                        sb.Append(run >= 2 ? "ss" : "s");
                        break;
                    case 't':
                        flush();
                        if (twelveHour) sb.Append(run >= 2 ? "AM/PM" : "A/P");
                        break;
                    case 'f':
                    case 'F':
                    case 'g':
                    case 'z':
                    case 'K':
                        break;
                    case '/':
                        literal.Append(culture.DateTimeFormat.DateSeparator);
                        break;
                    case ':':
                        literal.Append(culture.DateTimeFormat.TimeSeparator);
                        break;
                    case '\'':
                    case '"':
                        {
                            int end = pattern.IndexOf(c, i + 1);
                            if (end < 0) end = pattern.Length;
                            literal.Append(pattern.Substring(i + 1, end - i - 1));
                            i = end + 1;
                            continue;
                        }
                    case '\\':
                        if (i + 1 < pattern.Length) literal.Append(pattern[i + 1]);
                        i += 2;
                        continue;
                    default:
                        literal.Append(c, run);
                        break;
                }
                i += run;
            }
            flush();
            return sb.ToString();
        }

        /// <summary>Découpe le code en sections (« ; » hors guillemets, crochets et caractères échappés).</summary>
        internal static List<string> SplitSections(string code)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool quoted = false, bracket = false;
            for (int i = 0; i < code.Length; i++)
            {
                char c = code[i];
                if (quoted)
                {
                    sb.Append(c);
                    if (c == '"') quoted = false;
                    continue;
                }
                if (bracket)
                {
                    sb.Append(c);
                    if (c == ']') bracket = false;
                    continue;
                }
                if (c == '\\' || c == '_' || c == '*')
                {
                    sb.Append(c);
                    if (i + 1 < code.Length) sb.Append(code[++i]);
                    continue;
                }
                if (c == '"') quoted = true;
                else if (c == '[') bracket = true;
                else if (c == ';')
                {
                    result.Add(sb.ToString());
                    sb.Length = 0;
                    continue;
                }
                sb.Append(c);
            }
            result.Add(sb.ToString());
            return result;
        }

        // ================================================================== sections

        private enum TokenKind
        {
            Literal,
            Pad,
            Digit,
            DecimalPoint,
            Comma,
            Percent,
            Exponent,
            Slash,
            Text,
            General,
            Date,
            AmPm,
            Elapsed,
            SubSecond
        }

        private sealed class Token
        {
            public TokenKind Kind;
            /// <summary>Texte littéral, caractère de chiffre (0 # ?), lettre de date (y m d h s e g).</summary>
            public string Value;
            public int Count;
            public bool Minute;
            public bool ExponentPlus;
            public string Am;
            public string Pm;

            public char Char
            {
                get { return Value[0]; }
            }

            public override string ToString()
            {
                return Kind + ":" + Value + (Count > 1 ? "×" + Count : string.Empty);
            }
        }

        private sealed class Condition
        {
            public string Operator;
            public double Value;

            public bool Matches(double v)
            {
                switch (Operator)
                {
                    case "<": return v < Value;
                    case "<=": return v <= Value;
                    case ">": return v > Value;
                    case ">=": return v >= Value;
                    case "<>": return v != Value;
                    default: return v == Value;
                }
            }

            /// <summary>La condition ne retient que des nombres négatifs : la section porte elle-même le signe.</summary>
            public bool ExcludesNonNegative
            {
                get { return (Operator == "<" && Value <= 0) || (Operator == "<=" && Value < 0); }
            }
        }

        private sealed class Section
        {
            private readonly List<Token> _tokens = new List<Token>();
            private string _colorName;
            private int _colorIndex = -1;
            private CultureInfo _locale;
            private bool _systemLongDate;
            private bool _systemLongTime;

            public Condition Condition;
            public bool IsDate;
            public bool IsGeneral;
            public bool HasText;
            public bool HasDigits;
            private bool _hasAmPm;
            private bool _showsTime;
            private int _secondDecimals;

            // Nombres
            private int _percentCount;
            private int _scaleCommas;
            private bool _thousands;
            private bool _scientific;
            private bool _fraction;

            public static Section Parse(string text, CultureInfo culture)
            {
                var s = new Section();
                s.Tokenize(text);
                s.Classify(culture);
                return s;
            }

            /// <summary>La section se limite à « @ » (sans littéral ni couleur) : le texte s'affiche tel quel.</summary>
            public bool IsOnlyTextPlaceholder
            {
                get { return _colorName == null && _colorIndex < 0 && _tokens.Count == 1 && _tokens[0].Kind == TokenKind.Text; }
            }

            public Rgb? ResolveColor(ExcelColors colors)
            {
                if (_colorIndex >= 1) return colors.Indexed(_colorIndex + 7);
                switch (_colorName)
                {
                    case "black": return Rgb.Black;
                    case "white": return Rgb.White;
                    case "red": return new Rgb(255, 0, 0);
                    case "green": return new Rgb(0, 255, 0);
                    case "blue": return new Rgb(0, 0, 255);
                    case "yellow": return new Rgb(255, 255, 0);
                    case "magenta": return new Rgb(255, 0, 255);
                    case "cyan": return new Rgb(0, 255, 255);
                    default: return null;
                }
            }

            // -------------------------------------------------------------- analyse

            private void Tokenize(string text)
            {
                int i = 0;
                while (i < text.Length)
                {
                    char c = text[i];
                    char lower = char.ToLowerInvariant(c);
                    switch (c)
                    {
                        case '"':
                            {
                                int end = text.IndexOf('"', i + 1);
                                if (end < 0) end = text.Length;
                                AddLiteral(text.Substring(i + 1, end - i - 1));
                                i = end + 1;
                                continue;
                            }
                        case '\\':
                            if (i + 1 < text.Length) AddLiteral(text[i + 1].ToString());
                            i += 2;
                            continue;
                        case '_':
                            _tokens.Add(new Token { Kind = TokenKind.Pad, Value = " " });
                            i += 2;
                            continue;
                        case '*':
                            i += 2; // caractère de remplissage : sans équivalent dans un tableau Word
                            continue;
                        case '[':
                            {
                                int end = text.IndexOf(']', i + 1);
                                if (end < 0) end = text.Length;
                                ParseBracket(text.Substring(i + 1, end - i - 1));
                                i = end + 1;
                                continue;
                            }
                        case '0':
                        case '#':
                        case '?':
                            _tokens.Add(new Token { Kind = TokenKind.Digit, Value = c.ToString() });
                            i++;
                            continue;
                        case '.':
                            _tokens.Add(new Token { Kind = TokenKind.DecimalPoint, Value = "." });
                            i++;
                            continue;
                        case ',':
                            _tokens.Add(new Token { Kind = TokenKind.Comma, Value = "," });
                            i++;
                            continue;
                        case '%':
                            _tokens.Add(new Token { Kind = TokenKind.Percent, Value = "%" });
                            i++;
                            continue;
                        case '/':
                            _tokens.Add(new Token { Kind = TokenKind.Slash, Value = "/" });
                            i++;
                            continue;
                        case '@':
                            _tokens.Add(new Token { Kind = TokenKind.Text, Value = "@" });
                            i++;
                            continue;
                    }

                    if ((c == 'E' || c == 'e') && i + 1 < text.Length && (text[i + 1] == '+' || text[i + 1] == '-'))
                    {
                        _tokens.Add(new Token { Kind = TokenKind.Exponent, Value = c.ToString(), ExponentPlus = text[i + 1] == '+' });
                        i += 2;
                        continue;
                    }
                    if (lower == 'g' && string.Compare(text, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        _tokens.Add(new Token { Kind = TokenKind.General, Value = "General" });
                        i += 7;
                        continue;
                    }
                    if (lower == 'a' && string.Compare(text, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        string am = text.Substring(i, 2), pm = text.Substring(i + 3, 2);
                        _tokens.Add(new Token { Kind = TokenKind.AmPm, Value = "AM/PM", Am = am, Pm = pm });
                        i += 5;
                        continue;
                    }
                    if (lower == 'a' && string.Compare(text, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        _tokens.Add(new Token { Kind = TokenKind.AmPm, Value = "A/P", Am = text.Substring(i, 1), Pm = text.Substring(i + 2, 1) });
                        i += 3;
                        continue;
                    }
                    if (lower == 'b' && i + 1 < text.Length && (text[i + 1] == '1' || text[i + 1] == '2'))
                    {
                        i += 2; // choix du calendrier (B1 / B2) : ignoré
                        continue;
                    }
                    if (lower == 'y' || lower == 'm' || lower == 'd' || lower == 'h' || lower == 's' || lower == 'e' || lower == 'g')
                    {
                        int run = 1;
                        while (i + run < text.Length && char.ToLowerInvariant(text[i + run]) == lower) run++;
                        _tokens.Add(new Token { Kind = TokenKind.Date, Value = lower.ToString(), Count = run });
                        i += run;
                        continue;
                    }
                    AddLiteral(c.ToString());
                    i++;
                }
            }

            private void AddLiteral(string text)
            {
                if (text.Length == 0) return;
                if (_tokens.Count > 0 && _tokens[_tokens.Count - 1].Kind == TokenKind.Literal)
                {
                    _tokens[_tokens.Count - 1].Value += text;
                    return;
                }
                _tokens.Add(new Token { Kind = TokenKind.Literal, Value = text });
            }

            private void ParseBracket(string content)
            {
                string b = content.Trim();
                if (b.Length == 0) return;
                string lower = b.ToLowerInvariant();

                if (b[0] == '<' || b[0] == '>' || b[0] == '=')
                {
                    int opLength = 1;
                    if (b.Length > 1 && (b[1] == '=' || b[1] == '>')) opLength = 2;
                    double value;
                    if (double.TryParse(b.Substring(opLength).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                    {
                        Condition = new Condition { Operator = b.Substring(0, opLength), Value = value };
                    }
                    return;
                }
                if (b[0] == '$')
                {
                    string rest = b.Substring(1);
                    int dash = rest.LastIndexOf('-');
                    string currency = dash >= 0 ? rest.Substring(0, dash) : rest;
                    string locale = dash >= 0 ? rest.Substring(dash + 1) : null;
                    // « [$-x-sysdate] » : le tiret appartient au code régional.
                    if (lower.Contains("x-sysdate")) { _systemLongDate = true; return; }
                    if (lower.Contains("x-systime")) { _systemLongTime = true; return; }
                    if (currency.Length > 0) AddLiteral(currency);
                    if (!string.IsNullOrEmpty(locale)) ApplyLocale(locale);
                    return;
                }
                if (lower == "black" || lower == "white" || lower == "red" || lower == "green" || lower == "blue"
                    || lower == "yellow" || lower == "magenta" || lower == "cyan")
                {
                    _colorName = lower;
                    return;
                }
                if (lower.StartsWith("color", StringComparison.Ordinal))
                {
                    int index;
                    if (int.TryParse(lower.Substring(5), NumberStyles.Integer, CultureInfo.InvariantCulture, out index) && index >= 1 && index <= 56) _colorIndex = index;
                    return;
                }
                char first = lower[0];
                if ((first == 'h' || first == 'm' || first == 's') && lower.Trim(first).Length == 0)
                {
                    _tokens.Add(new Token { Kind = TokenKind.Elapsed, Value = first.ToString(), Count = lower.Length });
                    return;
                }
                // [DBNum1], [natnum], etc. : sans effet sur les chiffres affichés ici.
            }

            private void ApplyLocale(string locale)
            {
                int lcid;
                if (int.TryParse(locale, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out lcid))
                {
                    int language = lcid & 0xFFFF;
                    if (language == 0xF800) { _systemLongDate = true; return; }
                    if (language == 0xF400) { _systemLongTime = true; return; }
                    if (language == 0) return;
                    try
                    {
                        _locale = CultureInfo.GetCultureInfo(language);
                    }
                    catch (ArgumentException)
                    {
                        _locale = null;
                    }
                    return;
                }
                try
                {
                    _locale = CultureInfo.GetCultureInfo(locale);
                }
                catch (ArgumentException)
                {
                    _locale = null;
                }
            }

            private void Classify(CultureInfo culture)
            {
                foreach (var t in _tokens)
                {
                    if (t.Kind == TokenKind.General) IsGeneral = true;
                    else if (t.Kind == TokenKind.Text) HasText = true;
                    else if (t.Kind == TokenKind.AmPm || t.Kind == TokenKind.Elapsed) IsDate = true;
                    else if (t.Kind == TokenKind.Date && t.Value != "e" && t.Value != "g") IsDate = true;
                }
                if (_systemLongDate || _systemLongTime) IsDate = true;

                if (IsDate)
                {
                    PrepareDate(culture);
                    return;
                }

                // Hors date, les lettres e / g isolées sont du texte.
                for (int i = 0; i < _tokens.Count; i++)
                {
                    var t = _tokens[i];
                    if (t.Kind == TokenKind.Date) _tokens[i] = new Token { Kind = TokenKind.Literal, Value = new string(t.Value[0], t.Count) };
                }
                if (IsGeneral) return;
                PrepareNumber();
            }

            private void PrepareDate(CultureInfo settingsCulture)
            {
                if (_systemLongDate || _systemLongTime)
                {
                    var culture = _locale ?? settingsCulture;
                    string pattern = _systemLongDate ? culture.DateTimeFormat.LongDatePattern : culture.DateTimeFormat.LongTimePattern;
                    var replacement = Parse(DotNetPatternToExcel(pattern, culture), culture);
                    _tokens.Clear();
                    _tokens.AddRange(replacement._tokens);
                    _hasAmPm = replacement._hasAmPm;
                    _showsTime = replacement._showsTime;
                    _secondDecimals = replacement._secondDecimals;
                    return;
                }

                var result = new List<Token>();
                for (int i = 0; i < _tokens.Count; i++)
                {
                    var t = _tokens[i];
                    if (t.Kind == TokenKind.AmPm) _hasAmPm = true;
                    if (t.Kind == TokenKind.DecimalPoint && i + 1 < _tokens.Count && IsZeroDigit(_tokens[i + 1]) && PreviousIsSeconds(result))
                    {
                        int count = 0;
                        int j = i + 1;
                        while (j < _tokens.Count && IsZeroDigit(_tokens[j]))
                        {
                            count++;
                            j++;
                        }
                        count = Math.Min(count, 3);
                        _secondDecimals = Math.Max(_secondDecimals, count);
                        result.Add(new Token { Kind = TokenKind.SubSecond, Value = ".", Count = count });
                        i = j - 1;
                        continue;
                    }
                    switch (t.Kind)
                    {
                        case TokenKind.Digit:
                        case TokenKind.DecimalPoint:
                        case TokenKind.Comma:
                        case TokenKind.Percent:
                        case TokenKind.Slash:
                        case TokenKind.Exponent:
                            result.Add(new Token { Kind = TokenKind.Literal, Value = t.Kind == TokenKind.Exponent ? t.Value + (t.ExponentPlus ? "+" : "-") : t.Value });
                            continue;
                        case TokenKind.Text:
                        case TokenKind.General:
                            continue;
                    }
                    result.Add(t);
                }

                // « m » : minutes après une heure ou avant des secondes, mois sinon.
                for (int i = 0; i < result.Count; i++)
                {
                    var t = result[i];
                    if (t.Kind != TokenKind.Date || t.Value != "m" || t.Count > 2) continue;
                    var previous = NeighbourTimeToken(result, i, -1);
                    var next = NeighbourTimeToken(result, i, 1);
                    bool afterHour = previous != null && ((previous.Kind == TokenKind.Date && previous.Value == "h") || (previous.Kind == TokenKind.Elapsed && previous.Value == "h"));
                    bool beforeSecond = next != null && ((next.Kind == TokenKind.Date && next.Value == "s") || (next.Kind == TokenKind.Elapsed && next.Value == "s"));
                    t.Minute = afterHour || beforeSecond;
                }
                foreach (var t in result)
                {
                    if (t.Kind == TokenKind.AmPm || t.Kind == TokenKind.Elapsed || t.Kind == TokenKind.SubSecond
                        || (t.Kind == TokenKind.Date && (t.Value == "h" || t.Value == "s" || (t.Value == "m" && t.Minute))))
                    {
                        _showsTime = true;
                    }
                }
                _tokens.Clear();
                _tokens.AddRange(result);
            }

            private static bool IsZeroDigit(Token t)
            {
                return t.Kind == TokenKind.Digit && t.Value == "0";
            }

            private static bool PreviousIsSeconds(List<Token> tokens)
            {
                if (tokens.Count == 0) return false;
                var t = tokens[tokens.Count - 1];
                return (t.Kind == TokenKind.Date && t.Value == "s") || (t.Kind == TokenKind.Elapsed && t.Value == "s");
            }

            private static Token NeighbourTimeToken(List<Token> tokens, int index, int direction)
            {
                for (int i = index + direction; i >= 0 && i < tokens.Count; i += direction)
                {
                    var t = tokens[i];
                    if (t.Kind == TokenKind.Date || t.Kind == TokenKind.Elapsed) return t;
                }
                return null;
            }

            private void PrepareNumber()
            {
                foreach (var t in _tokens)
                {
                    if (t.Kind == TokenKind.Digit) HasDigits = true;
                    else if (t.Kind == TokenKind.Percent) _percentCount++;
                    else if (t.Kind == TokenKind.Exponent) _scientific = true;
                }

                // Fractions : barre oblique précédée d'un chiffre et suivie d'un chiffre ou d'un nombre fixe.
                int slash = _tokens.FindIndex(t => t.Kind == TokenKind.Slash);
                if (!_scientific && slash > 0 && _tokens[slash - 1].Kind == TokenKind.Digit && slash + 1 < _tokens.Count
                    && (_tokens[slash + 1].Kind == TokenKind.Digit || (_tokens[slash + 1].Kind == TokenKind.Literal && _tokens[slash + 1].Value.Length > 0 && char.IsDigit(_tokens[slash + 1].Value[0]))))
                {
                    _fraction = true;
                }
                else
                {
                    for (int i = 0; i < _tokens.Count; i++)
                    {
                        if (_tokens[i].Kind == TokenKind.Slash) _tokens[i] = new Token { Kind = TokenKind.Literal, Value = "/" };
                    }
                }

                // Virgules : séparateur de milliers (entre deux chiffres de la partie entière) ou division par 1000.
                int decimalPoint = _tokens.FindIndex(t => t.Kind == TokenKind.DecimalPoint);
                var cleaned = new List<Token>();
                for (int i = 0; i < _tokens.Count; i++)
                {
                    var t = _tokens[i];
                    if (t.Kind != TokenKind.Comma)
                    {
                        cleaned.Add(t);
                        continue;
                    }
                    int p = i - 1;
                    while (p >= 0 && _tokens[p].Kind == TokenKind.Comma) p--;
                    if (p < 0 || _tokens[p].Kind != TokenKind.Digit)
                    {
                        cleaned.Add(new Token { Kind = TokenKind.Literal, Value = "," });
                        continue;
                    }
                    int n = i + 1;
                    while (n < _tokens.Count && _tokens[n].Kind == TokenKind.Comma) n++;
                    bool digitFollows = n < _tokens.Count && _tokens[n].Kind == TokenKind.Digit;
                    bool inIntegerPart = decimalPoint < 0 || i < decimalPoint;
                    if (digitFollows && inIntegerPart) _thousands = true;
                    else if (!digitFollows) _scaleCommas++;
                    // sinon : virgule entre deux chiffres décimaux, ignorée
                }
                _tokens.Clear();
                _tokens.AddRange(cleaned);
            }

            // -------------------------------------------------------------- nombres

            public string FormatNumber(double abs, bool negative, ExcelFormatSettings settings)
            {
                var nfi = settings.Culture.NumberFormat;
                string body;
                if (_tokens.Count == 0) return string.Empty;
                if (IsGeneral) body = Render(GeneralBody(ExcelDecimal.FromDouble(abs), nfi.NumberDecimalSeparator), null, null, settings);
                else if (!HasDigits) body = Render(null, null, null, settings);
                else if (_scientific) body = FormatScientific(abs, settings);
                else if (_fraction) body = FormatFraction(abs, settings);
                else body = FormatFixed(abs, settings);
                if (negative) body = InsertMinus(body, nfi.NegativeSign);
                return body;
            }

            private static string InsertMinus(string text, string minus)
            {
                int i = 0;
                while (i < text.Length && text[i] == Pad) i++;
                return text.Substring(0, i) + minus + text.Substring(i);
            }

            private ExcelDecimal Scaled(double abs)
            {
                return ExcelDecimal.FromDouble(abs).Shift(2 * _percentCount - 3 * _scaleCommas);
            }

            private string FormatFixed(double abs, ExcelFormatSettings settings)
            {
                int decimalPoint = _tokens.FindIndex(t => t.Kind == TokenKind.DecimalPoint);
                var integerTokens = new List<int>();
                var decimalTokens = new List<int>();
                for (int i = 0; i < _tokens.Count; i++)
                {
                    if (_tokens[i].Kind != TokenKind.Digit) continue;
                    if (decimalPoint >= 0 && i > decimalPoint) decimalTokens.Add(i);
                    else integerTokens.Add(i);
                }
                var value = Scaled(abs).RoundDecimals(decimalTokens.Count);
                var output = new Dictionary<int, string>();
                FillInteger(value.IntegerDigits(), integerTokens, _thousands, settings, output);
                FillDecimals(value.FractionDigits(decimalTokens.Count), decimalTokens, output);
                return Render(null, output, decimalPoint, settings);
            }

            /// <summary>Répartit les chiffres entiers dans les emplacements (de droite à gauche) ; le premier reçoit les chiffres en trop.</summary>
            private void FillInteger(string digits, List<int> placeholders, bool thousands, ExcelFormatSettings settings, Dictionary<int, string> output)
            {
                if (placeholders.Count == 0) return;
                int n = placeholders.Count;
                var parts = new string[n];
                for (int k = 0; k < n; k++)
                {
                    int fromRight = n - 1 - k;
                    char kind = _tokens[placeholders[k]].Char;
                    if (fromRight < digits.Length) parts[k] = digits[digits.Length - 1 - fromRight].ToString();
                    else parts[k] = kind == '0' ? "0" : kind == '?' ? Pad.ToString() : string.Empty;
                }
                if (digits.Length > n) parts[0] = digits.Substring(0, digits.Length - n) + parts[0];

                if (thousands)
                {
                    var sb = new StringBuilder();
                    int pads = 0;
                    foreach (var p in parts)
                    {
                        if (p == Pad.ToString()) pads++;
                        else sb.Append(p);
                    }
                    string grouped = Group(sb.ToString(), settings.Culture.NumberFormat);
                    output[placeholders[0]] = new string(Pad, pads) + grouped;
                    for (int k = 1; k < n; k++) output[placeholders[k]] = string.Empty;
                    return;
                }
                for (int k = 0; k < n; k++) output[placeholders[k]] = parts[k];
            }

            private static string Group(string digits, NumberFormatInfo nfi)
            {
                if (digits.Length <= 3) return digits;
                string separator = nfi.NumberGroupSeparator;
                int[] sizes = nfi.NumberGroupSizes;
                int size = sizes != null && sizes.Length > 0 && sizes[0] > 0 ? sizes[0] : 3;
                int next = sizes != null && sizes.Length > 1 && sizes[1] > 0 ? sizes[1] : size; // Inde : 3 puis 2
                var groups = new List<string>();
                int end = digits.Length;
                int current = size;
                while (end > 0)
                {
                    int start = Math.Max(0, end - current);
                    groups.Insert(0, digits.Substring(start, end - start));
                    end = start;
                    current = next;
                }
                return string.Join(separator, groups.ToArray());
            }

            private void FillDecimals(string digits, List<int> placeholders, Dictionary<int, string> output)
            {
                int last = placeholders.Count - 1;
                while (last >= 0 && digits[last] == '0' && _tokens[placeholders[last]].Char != '0') last--;
                for (int k = 0; k < placeholders.Count; k++)
                {
                    char kind = _tokens[placeholders[k]].Char;
                    if (k <= last) output[placeholders[k]] = digits[k].ToString();
                    else output[placeholders[k]] = kind == '?' ? Pad.ToString() : kind == '0' ? "0" : string.Empty;
                }
            }

            private string FormatScientific(double abs, ExcelFormatSettings settings)
            {
                int exponentToken = _tokens.FindIndex(t => t.Kind == TokenKind.Exponent);
                int decimalPoint = _tokens.FindIndex(t => t.Kind == TokenKind.DecimalPoint);
                if (decimalPoint > exponentToken) decimalPoint = -1;
                var integerTokens = new List<int>();
                var decimalTokens = new List<int>();
                var exponentTokens = new List<int>();
                for (int i = 0; i < _tokens.Count; i++)
                {
                    if (_tokens[i].Kind != TokenKind.Digit) continue;
                    if (i > exponentToken) exponentTokens.Add(i);
                    else if (decimalPoint >= 0 && i > decimalPoint) decimalTokens.Add(i);
                    else integerTokens.Add(i);
                }

                var value = Scaled(abs);
                int step = Math.Max(1, integerTokens.Count);
                int exponent = 0;
                ExcelDecimal mantissa = ExcelDecimal.Zero;
                if (!value.IsZero)
                {
                    int e = value.Log10Floor;
                    exponent = FloorDiv(e, step) * step;
                    mantissa = value.Shift(-exponent).RoundDecimals(decimalTokens.Count);
                    if (mantissa.Log10Floor >= step)
                    {
                        exponent += step;
                        mantissa = value.Shift(-exponent).RoundDecimals(decimalTokens.Count);
                    }
                }

                var output = new Dictionary<int, string>();
                FillInteger(mantissa.IntegerDigits(), integerTokens, false, settings, output);
                FillDecimals(mantissa.FractionDigits(decimalTokens.Count), decimalTokens, output);

                var e0 = _tokens[exponentToken];
                int minDigits = 0;
                foreach (var i in exponentTokens)
                {
                    if (_tokens[i].Char == '0') minDigits++;
                }
                string sign = exponent < 0 ? "-" : e0.ExponentPlus ? "+" : string.Empty;
                output[exponentToken] = e0.Value + sign + Math.Abs(exponent).ToString(new string('0', Math.Max(1, minDigits)), CultureInfo.InvariantCulture);
                foreach (var i in exponentTokens) output[i] = string.Empty;
                return Render(null, output, decimalPoint, settings);
            }

            private static int FloorDiv(int a, int b)
            {
                int q = a / b;
                if ((a % b != 0) && ((a < 0) != (b < 0))) q--;
                return q;
            }

            private string FormatFraction(double abs, ExcelFormatSettings settings)
            {
                int slash = _tokens.FindIndex(t => t.Kind == TokenKind.Slash);

                // Numérateur : chiffres contigus avant la barre ; partie entière : chiffres qui précèdent.
                var numeratorTokens = new List<int>();
                int i = slash - 1;
                while (i >= 0 && _tokens[i].Kind == TokenKind.Digit)
                {
                    numeratorTokens.Insert(0, i);
                    i--;
                }
                var integerTokens = new List<int>();
                for (int k = 0; k <= i; k++)
                {
                    if (_tokens[k].Kind == TokenKind.Digit) integerTokens.Add(k);
                }

                // Dénominateur : nombre fixe (« /8 », « /100 ») ou emplacements (« /?? »).
                var denominatorTokens = new List<int>();
                int partialLiteral = -1, partialDigits = 0;
                string fixedDenominator = null;
                int j = slash + 1;
                if (j < _tokens.Count && _tokens[j].Kind == TokenKind.Literal && char.IsDigit(_tokens[j].Value[0]) && _tokens[j].Value[0] != '0')
                {
                    var sb = new StringBuilder();
                    for (; j < _tokens.Count; j++)
                    {
                        var t = _tokens[j];
                        if (t.Kind == TokenKind.Digit && t.Value == "0")
                        {
                            sb.Append('0');
                            denominatorTokens.Add(j);
                            continue;
                        }
                        if (t.Kind != TokenKind.Literal) break;
                        int d = 0;
                        while (d < t.Value.Length && char.IsDigit(t.Value[d])) d++;
                        if (d == 0) break;
                        sb.Append(t.Value.Substring(0, d));
                        if (d < t.Value.Length)
                        {
                            partialLiteral = j; // « 8 kg » : la suite du littéral reste affichée
                            partialDigits = d;
                            break;
                        }
                        denominatorTokens.Add(j);
                    }
                    fixedDenominator = sb.ToString();
                }
                else
                {
                    while (j < _tokens.Count && _tokens[j].Kind == TokenKind.Digit)
                    {
                        denominatorTokens.Add(j);
                        j++;
                    }
                }

                double scaled = Scaled(abs).ToDouble();
                bool mixed = integerTokens.Count > 0;
                long whole = mixed ? (long)Math.Floor(scaled) : 0;
                double fraction = mixed ? scaled - whole : scaled;
                long numerator, denominator, fixedValue;
                if (fixedDenominator != null && long.TryParse(fixedDenominator, NumberStyles.None, CultureInfo.InvariantCulture, out fixedValue) && fixedValue > 0)
                {
                    denominator = fixedValue;
                    numerator = (long)Math.Round(fraction * denominator, MidpointRounding.AwayFromZero);
                }
                else
                {
                    fixedDenominator = null;
                    int places = Math.Max(1, Math.Min(4, denominatorTokens.Count));
                    long max = (long)Math.Pow(10, places) - 1;
                    Approximate(fraction, max, out numerator, out denominator);
                }
                if (mixed && denominator > 0 && numerator >= denominator)
                {
                    whole += numerator / denominator;
                    numerator %= denominator;
                }

                var output = new Dictionary<int, string>();
                if (mixed)
                {
                    string wholeDigits = whole == 0 ? (numerator == 0 ? "0" : string.Empty) : whole.ToString(CultureInfo.InvariantCulture);
                    FillInteger(wholeDigits, integerTokens, _thousands, settings, output);
                }

                string pad = Pad.ToString();
                if (mixed && numerator == 0)
                {
                    // Fraction nulle : Excel la remplace par des espaces, seule la partie entière est visible.
                    foreach (var k in numeratorTokens) output[k] = pad;
                    output[slash] = pad;
                    foreach (var k in denominatorTokens) output[k] = pad;
                    if (partialLiteral >= 0) output[partialLiteral] = new string(Pad, partialDigits) + _tokens[partialLiteral].Value.Substring(partialDigits);
                    return Render(null, output, -1, settings);
                }

                FillInteger(numerator.ToString(CultureInfo.InvariantCulture), numeratorTokens, false, settings, output);
                output[slash] = "/";
                if (fixedDenominator != null)
                {
                    foreach (var k in denominatorTokens) output[k] = _tokens[k].Kind == TokenKind.Digit ? "0" : _tokens[k].Value;
                }
                else
                {
                    // Dénominateur aligné à gauche : les emplacements « ? » inutilisés deviennent des espaces.
                    string den = denominator.ToString(CultureInfo.InvariantCulture);
                    int n = denominatorTokens.Count;
                    for (int k = 0; k < n; k++)
                    {
                        char kind = _tokens[denominatorTokens[k]].Char;
                        if (k < den.Length) output[denominatorTokens[k]] = k == n - 1 ? den.Substring(k) : den[k].ToString();
                        else output[denominatorTokens[k]] = kind == '?' ? pad : kind == '0' ? "0" : string.Empty;
                    }
                }
                return Render(null, output, -1, settings);
            }

            /// <summary>
            /// Fraction la plus proche de <paramref name="x"/> dont le dénominateur ne dépasse pas
            /// <paramref name="maxDenominator"/> (à égalité, le plus petit dénominateur) : π au format « # ??/?? » → 3 14/99.
            /// </summary>
            internal static void Approximate(double x, long maxDenominator, out long numerator, out long denominator)
            {
                numerator = (long)Math.Round(x, MidpointRounding.AwayFromZero);
                denominator = 1;
                double best = Math.Abs(x - numerator);
                for (long q = 2; q <= maxDenominator && best > 0; q++)
                {
                    long p = (long)Math.Round(x * q, MidpointRounding.AwayFromZero);
                    double error = Math.Abs(x - (double)p / q);
                    if (error < best - 1e-12)
                    {
                        best = error;
                        numerator = p;
                        denominator = q;
                    }
                }
            }

            /// <summary>Assemble la section : littéraux, valeurs calculées des emplacements, séparateur décimal.</summary>
            private string Render(string general, Dictionary<int, string> output, int? decimalPoint, ExcelFormatSettings settings)
            {
                var nfi = settings.Culture.NumberFormat;
                var sb = new StringBuilder();
                for (int i = 0; i < _tokens.Count; i++)
                {
                    var t = _tokens[i];
                    string computed;
                    if (output != null && output.TryGetValue(i, out computed))
                    {
                        sb.Append(computed);
                        continue;
                    }
                    switch (t.Kind)
                    {
                        case TokenKind.Literal:
                            sb.Append(t.Value);
                            break;
                        case TokenKind.Pad:
                            sb.Append(Pad);
                            break;
                        case TokenKind.Percent:
                            sb.Append(nfi.PercentSymbol);
                            break;
                        case TokenKind.DecimalPoint:
                            if (decimalPoint.HasValue && i == decimalPoint.Value) sb.Append(nfi.NumberDecimalSeparator);
                            else sb.Append('.');
                            break;
                        case TokenKind.General:
                            sb.Append(general ?? string.Empty);
                            break;
                        case TokenKind.Slash:
                            sb.Append('/');
                            break;
                        case TokenKind.Digit:
                        case TokenKind.Text:
                        case TokenKind.Comma:
                            break;
                        default:
                            sb.Append(t.Value);
                            break;
                    }
                }
                return sb.ToString();
            }

            // -------------------------------------------------------------- texte

            public string FormatText(string text, ExcelFormatSettings settings)
            {
                var sb = new StringBuilder();
                foreach (var t in _tokens)
                {
                    switch (t.Kind)
                    {
                        case TokenKind.Text:
                            sb.Append(text);
                            break;
                        case TokenKind.Literal:
                            sb.Append(t.Value);
                            break;
                        case TokenKind.Pad:
                            sb.Append(Pad);
                            break;
                        case TokenKind.General:
                            sb.Append(text);
                            break;
                    }
                }
                return StripPadding(sb.ToString());
            }

            // -------------------------------------------------------------- dates

            public bool FormatDate(double serial, ExcelFormatSettings settings, out string text)
            {
                text = null;
                ExcelDateParts parts;
                // L'heure n'est arrondie à la seconde (ou à la précision affichée) que si le format l'affiche.
                if (!ExcelDates.TryGetParts(serial, settings.Date1904, _showsTime ? _secondDecimals : -1, out parts)) return false;

                var culture = _locale ?? settings.Culture;
                var dtf = GregorianFormat(culture);
                var sb = new StringBuilder();
                foreach (var t in _tokens)
                {
                    switch (t.Kind)
                    {
                        case TokenKind.Literal:
                            sb.Append(t.Value);
                            break;
                        case TokenKind.Pad:
                            sb.Append(Pad);
                            break;
                        case TokenKind.AmPm:
                            sb.Append(parts.Hour < 12 ? t.Am : t.Pm);
                            break;
                        case TokenKind.SubSecond:
                            {
                                string ms = parts.Millisecond.ToString("000", CultureInfo.InvariantCulture);
                                sb.Append(settings.Culture.NumberFormat.NumberDecimalSeparator).Append(ms.Substring(0, t.Count));
                                break;
                            }
                        case TokenKind.Elapsed:
                            {
                                long total = t.Value == "h" ? parts.TotalMilliseconds / 3600000L
                                           : t.Value == "m" ? parts.TotalMilliseconds / 60000L
                                           : parts.TotalMilliseconds / 1000L;
                                sb.Append(total.ToString(new string('0', Math.Max(1, t.Count)), CultureInfo.InvariantCulture));
                                break;
                            }
                        case TokenKind.Date:
                            sb.Append(DatePart(t, parts, dtf));
                            break;
                    }
                }
                text = sb.ToString();
                return true;
            }

            private string DatePart(Token t, ExcelDateParts p, DateTimeFormatInfo dtf)
            {
                var inv = CultureInfo.InvariantCulture;
                switch (t.Value)
                {
                    case "y":
                        return t.Count <= 2 ? (p.Year % 100).ToString("00", inv) : p.Year.ToString("0000", inv);
                    case "e":
                        return p.Year.ToString(inv);
                    case "g":
                        return string.Empty;
                    case "d":
                        if (t.Count == 1) return p.Day.ToString(inv);
                        if (t.Count == 2) return p.Day.ToString("00", inv);
                        if (t.Count == 3) return dtf.GetAbbreviatedDayName((DayOfWeek)p.DayOfWeek);
                        return dtf.GetDayName((DayOfWeek)p.DayOfWeek);
                    case "h":
                        {
                            int hour = p.Hour;
                            if (_hasAmPm)
                            {
                                hour %= 12;
                                if (hour == 0) hour = 12;
                            }
                            return t.Count >= 2 ? hour.ToString("00", inv) : hour.ToString(inv);
                        }
                    case "s":
                        return t.Count >= 2 ? p.Second.ToString("00", inv) : p.Second.ToString(inv);
                    case "m":
                        if (t.Minute) return t.Count >= 2 ? p.Minute.ToString("00", inv) : p.Minute.ToString(inv);
                        if (t.Count == 1) return p.Month.ToString(inv);
                        if (t.Count == 2) return p.Month.ToString("00", inv);
                        if (t.Count == 3) return dtf.GetAbbreviatedMonthName(p.Month);
                        string full = dtf.GetMonthName(p.Month);
                        if (t.Count == 5) return full.Length > 0 ? full.Substring(0, 1) : string.Empty;
                        return full;
                    default:
                        return string.Empty;
                }
            }

            private static DateTimeFormatInfo GregorianFormat(CultureInfo culture)
            {
                var dtf = culture.DateTimeFormat;
                if (dtf.Calendar is GregorianCalendar) return dtf;
                foreach (var calendar in culture.OptionalCalendars)
                {
                    if (calendar is GregorianCalendar)
                    {
                        var clone = (DateTimeFormatInfo)dtf.Clone();
                        clone.Calendar = calendar;
                        return clone;
                    }
                }
                return CultureInfo.InvariantCulture.DateTimeFormat;
            }
        }
    }
}
