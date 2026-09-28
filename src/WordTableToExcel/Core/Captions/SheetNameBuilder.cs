using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace WordTableToExcel.Core.Captions
{
    /// <summary>
    /// Produit des noms de feuille Excel valides et uniques : 31 caractères maximum,
    /// sans <c>\ / ? * [ ] :</c>, sans apostrophe en début/fin, différent de « History ».
    /// </summary>
    public sealed class SheetNameBuilder
    {
        public const int MaxLength = 31;
        private readonly HashSet<string> _used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Nom par défaut d'un tableau sans légende : « Tableau_1 », « Tableau_2 »…</summary>
        public static string DefaultName(int tableNumber)
        {
            return "Tableau_" + tableNumber.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Réserve un nom unique à partir du texte souhaité (légende) ou, à défaut, du nom de repli.</summary>
        public string Reserve(string preferred, string fallback)
        {
            string name = Sanitize(preferred);
            if (name.Length == 0) name = Sanitize(fallback);
            if (name.Length == 0) name = "Tableau";

            if (_used.Add(name)) return name;

            for (int n = 2; n < 10000; n++)
            {
                string suffix = " (" + n.ToString(CultureInfo.InvariantCulture) + ")";
                string candidate = Truncate(name, MaxLength - suffix.Length).TrimEnd() + suffix;
                if (_used.Add(candidate)) return candidate;
            }
            throw new InvalidOperationException("Impossible de générer un nom de feuille unique.");
        }

        public static string Sanitize(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            var sb = new StringBuilder(raw.Length);
            bool pendingSpace = false;
            foreach (char c in raw)
            {
                char mapped = c;
                switch (c)
                {
                    case '\\':
                    case '/':
                    case '?':
                    case '*':
                    case '[':
                    case ']':
                    case ':':
                        mapped = '-';
                        break;
                }
                if (char.IsWhiteSpace(mapped) || char.IsControl(mapped) || mapped == '\u200B')
                {
                    pendingSpace = sb.Length > 0;
                    continue;
                }
                if (pendingSpace) sb.Append(' ');
                pendingSpace = false;
                sb.Append(mapped);
            }

            string name = CollapseDashes(sb.ToString()).Trim().Trim('\'').Trim();
            name = Truncate(name, MaxLength).Trim().Trim('\'').Trim();
            bool meaningful = false;
            foreach (char c in name) meaningful |= char.IsLetterOrDigit(c);
            if (!meaningful) return string.Empty; // « :: » ou « [] » : on préfère le nom par défaut
            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase)) name = "History_";
            return name;
        }

        private static string CollapseDashes(string s)
        {
            // « Tableau 1 : Titre » devient « Tableau 1 - Titre » et non « Tableau 1 -- Titre ».
            while (s.Contains("--")) s = s.Replace("--", "-");
            return s;
        }

        private static string Truncate(string s, int max)
        {
            if (s.Length <= max) return s;
            int cut = max;
            if (char.IsHighSurrogate(s[cut - 1])) cut--;
            return s.Substring(0, cut);
        }
    }
}
