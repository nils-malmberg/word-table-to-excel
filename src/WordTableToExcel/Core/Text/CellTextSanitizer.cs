using System.Collections.Generic;
using System.Text;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Text
{
    /// <summary>
    /// Nettoie le texte lu dans Word pour qu'il soit valide et lisible dans Excel :
    /// marques de fin de cellule, sauts de ligne, codes de champ, caractères de contrôle
    /// interdits en XML, limite de 32 767 caractères par cellule.
    /// </summary>
    public static class CellTextSanitizer
    {
        public const int ExcelMaxCellLength = 32767;

        private const char FieldBegin = '\u0013';
        private const char FieldSeparator = '\u0014';
        private const char FieldEnd = '\u0015';

        /// <summary>
        /// Nettoie une suite de segments : conserve leur mise en forme, fusionne les segments
        /// identiques consécutifs et supprime les sauts de ligne finaux.
        /// </summary>
        public static List<TextRun> Clean(IEnumerable<TextRun> runs)
        {
            var result = new List<TextRun>();
            // Pile des champs ouverts : true = on est dans le code du champ (invisible).
            var fieldStack = new List<bool>();
            int total = 0;

            foreach (var run in runs)
            {
                if (run == null || string.IsNullOrEmpty(run.Text)) continue;
                var sb = new StringBuilder(run.Text.Length);
                string text = run.Text;

                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (c == FieldBegin)
                    {
                        fieldStack.Add(true);
                        continue;
                    }
                    if (c == FieldSeparator)
                    {
                        if (fieldStack.Count > 0) fieldStack[fieldStack.Count - 1] = false;
                        continue;
                    }
                    if (c == FieldEnd)
                    {
                        if (fieldStack.Count > 0) fieldStack.RemoveAt(fieldStack.Count - 1);
                        continue;
                    }
                    if (fieldStack.Contains(true)) continue; // code de champ masqué

                    if (char.IsHighSurrogate(c))
                    {
                        if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                        {
                            sb.Append(c).Append(text[i + 1]);
                            i++;
                        }
                        continue; // surrogate orphelin : invalide en XML
                    }
                    if (char.IsLowSurrogate(c)) continue;

                    char mapped;
                    if (TryMap(c, out mapped)) sb.Append(mapped);
                }

                if (sb.Length == 0) continue;

                string cleaned = sb.ToString();
                if (total + cleaned.Length > ExcelMaxCellLength)
                {
                    int keep = ExcelMaxCellLength - total;
                    if (keep <= 0) break;
                    if (keep < cleaned.Length && char.IsHighSurrogate(cleaned[keep - 1])) keep--;
                    cleaned = cleaned.Substring(0, keep);
                }
                total += cleaned.Length;

                var last = result.Count > 0 ? result[result.Count - 1] : null;
                if (last != null && last.HasSameFormatting(run))
                {
                    last.Text += cleaned;
                }
                else
                {
                    result.Add(new TextRun(cleaned, run.Format, run.Highlight));
                }
            }

            TrimTrailingLineBreaks(result);
            return result;
        }

        /// <summary>Nettoie une chaîne isolée (légende, nom de document…).</summary>
        public static string CleanPlain(string text)
        {
            var runs = Clean(new[] { new TextRun(text, null, null) });
            return runs.Count == 0 ? string.Empty : runs[0].Text;
        }

        /// <summary>
        /// Texte réduit à ses caractères visibles, pour comparer deux lectures d'un même contenu : codes de champ,
        /// caractères de contrôle, espaces et sauts de ligne sont ignorés (ils diffèrent selon la façon de lire),
        /// sans limite de longueur.
        /// </summary>
        public static string Comparable(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            var fieldStack = new List<bool>();
            foreach (char c in text)
            {
                if (c == FieldBegin)
                {
                    fieldStack.Add(true);
                    continue;
                }
                if (c == FieldSeparator)
                {
                    if (fieldStack.Count > 0) fieldStack[fieldStack.Count - 1] = false;
                    continue;
                }
                if (c == FieldEnd)
                {
                    if (fieldStack.Count > 0) fieldStack.RemoveAt(fieldStack.Count - 1);
                    continue;
                }
                if (fieldStack.Contains(true)) continue;
                if (c == '\u001E')
                {
                    sb.Append('-');
                    continue;
                }
                if (c < ' ' || char.IsWhiteSpace(c) || c == '﻿' || c == '​') continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static bool TryMap(char c, out char mapped)
        {
            mapped = c;
            switch (c)
            {
                case '\r':      // fin de paragraphe
                case '\n':
                case '\u000B':  // saut de ligne manuel (Maj+Entrée)
                case '\u000C':  // saut de page / section
                case '\u000E':  // saut de colonne
                case '\u2028':  // séparateur de ligne
                case '\u2029':  // séparateur de paragraphe
                    mapped = '\n';
                    return true;
                case '\t':
                    mapped = ' ';
                    return true;
                case '\u001E':  // trait d'union insécable Word
                    mapped = '-';
                    return true;
                case '\u001F':  // trait d'union conditionnel
                case '\u0007':  // marque de fin de cellule
                case '\uFEFF':
                case '\uFFFE':
                case '\uFFFF':
                    return false;
            }
            if (c < ' ') return false; // autres caractères de contrôle (objets, notes, ancres…)
            return true;
        }

        private static void TrimTrailingLineBreaks(List<TextRun> runs)
        {
            while (runs.Count > 0)
            {
                var last = runs[runs.Count - 1];
                string trimmed = last.Text.TrimEnd('\n');
                if (trimmed.Length == last.Text.Length) return;
                if (trimmed.Length == 0)
                {
                    runs.RemoveAt(runs.Count - 1);
                    continue;
                }
                last.Text = trimmed;
                return;
            }
        }
    }
}
