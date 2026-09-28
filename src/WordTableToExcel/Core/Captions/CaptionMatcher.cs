using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using WordTableToExcel.Core.Text;

namespace WordTableToExcel.Core.Captions
{
    /// <summary>
    /// Reconnaît les légendes de tableau, quelle que soit la langue du document :
    /// <list type="bullet">
    /// <item>champ <c>SEQ</c> dont l'identificateur désigne un tableau (SEQ Table, SEQ Tableau, SEQ Tabla, SEQ Tabelle…) ;</item>
    /// <item>paragraphe dont le texte commence par « Tabl… » ou par le libellé « Tableau » d'une autre langue.</item>
    /// </list>
    /// </summary>
    public sealed class CaptionMatcher
    {
        /// <summary>Libellés complets reconnus (forme normalisée : minuscules, sans accents).</summary>
        private static readonly string[] BuiltInLabelWords =
        {
            // Latin
            "table", "tables", "tableau", "tableaux", "tabla", "tablas", "tabl", "tab",
            "tabelle", "tabellen", "tabella", "tabel", "tabela", "tabele", "tabell",
            "tabulka", "tabula", "tablica", "tablice", "taulukko", "taula", "tablo", "tablazat",
            "lentele", "tafla", "tabelo", "jadual", "bang",
            // Cyrillique / grec
            "таблица", "таблиця", "табела", "табліца", "табл", "πινακας",
            // Autres écritures
            "جدول", "טבלה", "ตาราง", "तालिका", "表", "표"
        };

        /// <summary>Débuts de mot reconnus (identificateurs SEQ et textes « Tabl… »).</summary>
        private static readonly string[] BuiltInPrefixes =
        {
            "tabl", "tabel", "tabul", "taulu", "taula", "lentel", "tafl",
            "табл", "табел", "πινακ", "جدول", "טבלה", "ตาราง", "तालिका", "表", "표"
        };

        /// <summary>Écritures sans espace entre le libellé et le numéro (« 表1 »).</summary>
        private static readonly string[] AttachedLabels = { "表", "표", "ตาราง" };

        private static readonly Regex NumberingAfterLabel = new Regex(
            @"^[\s\u00A0\u202F.:\-\u2013\u2014]*(?:(?:n°|no\.?|nr\.?|num\.?|№)\s*)?(?:\d|[IVXLC]+(?![\p{L}])|[A-Z](?:[.\-\u2013]?\d|(?![\p{L}])))",
            RegexOptions.CultureInvariant);

        private static readonly Regex SeqCode = new Regex(
            @"^\s*SEQ\s+(?:""(?<id>[^""]+)""|(?<id>[^\s\\]+))",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private readonly HashSet<string> _labelWords = new HashSet<string>(StringComparer.Ordinal);
        private readonly List<string> _prefixes = new List<string>();

        public CaptionMatcher()
            : this(null)
        {
        }

        /// <param name="extraLabels">Libellés supplémentaires (ex. libellé de légende localisé de Word, libellés personnalisés).</param>
        public CaptionMatcher(IEnumerable<string> extraLabels)
        {
            foreach (var w in BuiltInLabelWords) _labelWords.Add(w);
            _prefixes.AddRange(BuiltInPrefixes);

            if (extraLabels != null)
            {
                foreach (var raw in extraLabels)
                {
                    string label = Normalize(raw).Trim().TrimEnd('.');
                    if (label.Length == 0) continue;
                    _labelWords.Add(label);
                    if (!_prefixes.Contains(label)) _prefixes.Add(label);
                }
            }
        }

        /// <summary>Extrait l'identificateur d'un code de champ SEQ (« SEQ Tableau \* ARABIC » → « Tableau »).</summary>
        public static string ParseSequenceIdentifier(string fieldCode)
        {
            if (string.IsNullOrEmpty(fieldCode)) return null;
            var m = SeqCode.Match(fieldCode.Replace('\u00A0', ' '));
            return m.Success ? m.Groups["id"].Value : null;
        }

        /// <summary>true si l'identificateur de séquence désigne des tableaux (Table, Tableau, Tabla, Tabelle…).</summary>
        public bool IsTableSequenceIdentifier(string identifier)
        {
            string id = Normalize(identifier);
            if (id.Length == 0) return false;
            if (_labelWords.Contains(id)) return true;
            foreach (var prefix in _prefixes)
            {
                if (id.StartsWith(prefix, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        /// <summary>true si le code de champ est un champ SEQ de tableau.</summary>
        public bool IsTableSequenceField(string fieldCode)
        {
            string id = ParseSequenceIdentifier(fieldCode);
            return id != null && IsTableSequenceIdentifier(id);
        }

        /// <summary>
        /// true si le texte du paragraphe ressemble à une légende de tableau :
        /// « Tableau 3 : … », « Table A-1 », « Tabla 2 », « 表1 », etc.
        /// </summary>
        /// <param name="hasCaptionStyle">Le paragraphe utilise le style « Légende » de Word.</param>
        public bool LooksLikeCaptionText(string paragraphText, bool hasCaptionStyle)
        {
            string text = CleanCaptionText(paragraphText);
            if (text.Length == 0 || text.Length > 400) return false;

            string normalized = Normalize(text);

            foreach (var label in AttachedLabels)
            {
                if (normalized.StartsWith(label, StringComparison.Ordinal))
                {
                    string rest = normalized.Substring(label.Length);
                    return rest.Length == 0 || !char.IsLetter(rest[0]) || hasCaptionStyle;
                }
            }

            int end = 0;
            while (end < text.Length && (char.IsLetter(text[end]) || CharUnicodeInfo.GetUnicodeCategory(text[end]) == UnicodeCategory.NonSpacingMark)) end++;
            if (end == 0) return false;

            string word = Normalize(text.Substring(0, end));
            string after = text.Substring(end);

            if (_labelWords.Contains(word))
            {
                // « Tab » seul est trop court pour être fiable sans numéro.
                if (word == "tab" || word == "tabl" || word == "bang") return NumberingAfterLabel.IsMatch(after) || hasCaptionStyle;
                return true;
            }

            foreach (var prefix in _prefixes)
            {
                if (word.StartsWith(prefix, StringComparison.Ordinal))
                {
                    return hasCaptionStyle || NumberingAfterLabel.IsMatch(after);
                }
            }
            return false;
        }

        /// <summary>Texte de légende lisible : sans codes de champ, caractères de contrôle ni espaces superflus.</summary>
        public static string CleanCaptionText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            string text = CellTextSanitizer.CleanPlain(raw);
            var sb = new StringBuilder(text.Length);
            bool space = false;
            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c) || c == '\u200B')
                {
                    space = sb.Length > 0;
                    continue;
                }
                if (space) sb.Append(' ');
                space = false;
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>Minuscules invariantes, sans diacritiques, sans espaces de tête.</summary>
        public static string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            string decomposed = s.Trim().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(decomposed.Length);
            foreach (char c in decomposed)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
                sb.Append(c);
            }
            return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
        }
    }
}
