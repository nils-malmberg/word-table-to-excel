using System;
using System.Globalization;
using System.Text;

namespace WordTableToExcel.Core.Text
{
    /// <summary>
    /// Reconnaît les textes qui représentent sans ambiguïté un nombre (« 1 234,50 », « 12,5 % »,
    /// « 1,234.5 », « 42 € »…) et produit la valeur et un format numérique Excel qui
    /// conserve l'apparence (séparateur de milliers, nombre de décimales, pourcentage, devise).
    /// Les cas ambigus ou risqués (zéros de tête, codes, dates…) restent du texte.
    /// </summary>
    public static class NumberParser
    {
        private const string CurrencySymbols = "€$£¥₹₽₩₺₫₪₱฿₴₦";

        public static bool TryParse(string text, CultureInfo culture, out double value, out string numberFormat)
        {
            value = 0;
            numberFormat = null;
            if (string.IsNullOrEmpty(text)) return false;
            if (culture == null) culture = CultureInfo.CurrentCulture;

            string s = TrimAll(text);
            if (s.Length == 0 || s.Length > 40) return false;

            bool negative = false;
            string currency = null;
            bool currencyPrefix = false;
            bool currencySpaced = false;
            bool percent = false;

            if (IsMinus(s[0]) || s[0] == '+')
            {
                negative = IsMinus(s[0]);
                s = TrimAll(s.Substring(1));
            }
            if (s.Length > 0 && CurrencySymbols.IndexOf(s[0]) >= 0)
            {
                currency = s.Substring(0, 1);
                currencyPrefix = true;
                string rest = s.Substring(1);
                currencySpaced = rest.Length > 0 && IsSpace(rest[0]);
                s = TrimAll(rest);
                if (!negative && s.Length > 0 && IsMinus(s[0]))
                {
                    negative = true;
                    s = TrimAll(s.Substring(1));
                }
            }
            if (s.Length > 0 && s[s.Length - 1] == '%')
            {
                percent = true;
                s = TrimAll(s.Substring(0, s.Length - 1));
            }
            else if (currency == null && s.Length > 0 && CurrencySymbols.IndexOf(s[s.Length - 1]) >= 0)
            {
                currency = s.Substring(s.Length - 1);
                string rest = s.Substring(0, s.Length - 1);
                currencySpaced = rest.Length > 0 && IsSpace(rest[rest.Length - 1]);
                s = TrimAll(rest);
            }

            if (s.Length == 0 || !char.IsDigit(s[0]) || !char.IsDigit(s[s.Length - 1])) return false;

            int dots = 0, commas = 0, lastDot = -1, lastComma = -1;
            bool hasSpaceGroup = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c >= '0' && c <= '9') continue;
                if (c == '.') { dots++; lastDot = i; continue; }
                if (c == ',') { commas++; lastComma = i; continue; }
                if (IsSpace(c) || c == '\'' || c == '\u2019') { hasSpaceGroup = true; continue; }
                return false;
            }

            int decimalPos = -1;
            char groupChar = '\0';

            if (dots > 0 && commas > 0)
            {
                bool dotIsDecimal = lastDot > lastComma;
                if ((dotIsDecimal ? dots : commas) != 1) return false;
                decimalPos = dotIsDecimal ? lastDot : lastComma;
                groupChar = dotIsDecimal ? ',' : '.';
            }
            else if (dots + commas > 0)
            {
                char sep = dots > 0 ? '.' : ',';
                int count = dots > 0 ? dots : commas;
                int pos = dots > 0 ? lastDot : lastComma;
                if (count > 1 || hasSpaceGroup)
                {
                    if (count == 1 && hasSpaceGroup)
                    {
                        // « 1 234,5 » : espaces = milliers, séparateur unique = décimales.
                        decimalPos = pos;
                    }
                    else
                    {
                        groupChar = sep;
                    }
                }
                else
                {
                    int digitsAfter = s.Length - pos - 1;
                    int digitsBefore = pos;
                    if (digitsAfter != 3 || digitsBefore > 3)
                    {
                        decimalPos = pos;
                    }
                    else
                    {
                        // « 1,234 » ou « 1.234 » : ambigu, on tranche avec la culture.
                        string dec = culture.NumberFormat.NumberDecimalSeparator;
                        string grp = culture.NumberFormat.NumberGroupSeparator;
                        if (dec == sep.ToString()) decimalPos = pos;
                        else if (grp == sep.ToString()) groupChar = sep;
                        else return false;
                    }
                }
            }

            string integerPart = decimalPos >= 0 ? s.Substring(0, decimalPos) : s;
            string fractionPart = decimalPos >= 0 ? s.Substring(decimalPos + 1) : string.Empty;

            foreach (char c in fractionPart)
            {
                if (c < '0' || c > '9') return false;
            }

            bool grouped;
            string integerDigits;
            if (!ParseIntegerPart(integerPart, groupChar, out integerDigits, out grouped)) return false;

            if (integerDigits.Length > 1 && integerDigits[0] == '0') return false; // code, n° de téléphone…
            if (integerDigits.Length + fractionPart.Length > 15) return false;    // précision Excel

            string invariant = integerDigits + (fractionPart.Length > 0 ? "." + fractionPart : string.Empty);
            double parsed;
            if (!double.TryParse(invariant, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out parsed)) return false;

            if (negative) parsed = -parsed;
            if (percent) parsed /= 100.0;
            value = parsed;

            bool needsFormat = grouped || fractionPart.Length > 0 || percent || currency != null;
            if (!needsFormat) return true;

            var fmt = new StringBuilder();
            fmt.Append(grouped ? "#,##0" : "0");
            if (fractionPart.Length > 0) fmt.Append('.').Append('0', fractionPart.Length);
            if (percent) fmt.Append('%');
            if (currency != null)
            {
                string literal = "\"" + (currencyPrefix
                    ? currency + (currencySpaced ? " " : string.Empty)
                    : (currencySpaced ? " " : string.Empty) + currency) + "\"";
                if (currencyPrefix) fmt.Insert(0, literal);
                else fmt.Append(literal);
            }
            numberFormat = fmt.ToString();
            return true;
        }

        private static bool ParseIntegerPart(string part, char groupChar, out string digits, out bool grouped)
        {
            digits = null;
            grouped = false;
            var groups = new System.Collections.Generic.List<string>();
            var current = new StringBuilder();
            foreach (char c in part)
            {
                if (c >= '0' && c <= '9')
                {
                    current.Append(c);
                    continue;
                }
                bool isGroup = IsSpace(c) || c == '\'' || c == '\u2019' || (groupChar != '\0' && c == groupChar);
                if (!isGroup || current.Length == 0) return false;
                groups.Add(current.ToString());
                current.Length = 0;
            }
            if (current.Length == 0) return false;
            groups.Add(current.ToString());

            if (groups.Count > 1)
            {
                if (groups[0].Length > 3) return false;
                for (int i = 1; i < groups.Count; i++)
                {
                    if (groups[i].Length != 3) return false;
                }
                grouped = true;
            }
            digits = string.Concat(groups.ToArray());
            return true;
        }

        private static bool IsMinus(char c)
        {
            return c == '-' || c == '\u2212' || c == '\u2013';
        }

        private static bool IsSpace(char c)
        {
            return c == ' ' || c == '\u00A0' || c == '\u202F' || c == '\u2009';
        }

        private static string TrimAll(string s)
        {
            return s.Trim(' ', '\u00A0', '\u202F', '\u2009', '\t', '\n', '\r');
        }
    }
}
