using System;
using System.Collections.Generic;
using System.Globalization;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Textes affichés par Excel pour les valeurs logiques et les erreurs, dans la langue d'Office
    /// (le fichier enregistre toujours la forme anglaise : TRUE, #DIV/0!, #VALUE!…).
    /// </summary>
    public static class ExcelLocaleTexts
    {
        private static readonly string[] ErrorCodes = { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" };

        private static readonly Dictionary<string, string[]> Booleans = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "fr", new[] { "VRAI", "FAUX" } },
            { "de", new[] { "WAHR", "FALSCH" } },
            { "es", new[] { "VERDADERO", "FALSO" } },
            { "it", new[] { "VERO", "FALSO" } },
            { "pt", new[] { "VERDADEIRO", "FALSO" } },
            { "nl", new[] { "WAAR", "ONWAAR" } },
        };

        // Même ordre que ErrorCodes ; les erreurs plus récentes (#SPILL!, #CALC!…) restent en anglais.
        private static readonly Dictionary<string, string[]> Errors = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "fr", new[] { "#NUL!", "#DIV/0!", "#VALEUR!", "#REF!", "#NOM?", "#NOMBRE!", "#N/A" } },
            { "de", new[] { "#NULL!", "#DIV/0!", "#WERT!", "#BEZUG!", "#NAME?", "#ZAHL!", "#NV" } },
            { "es", new[] { "#¡NULO!", "#¡DIV/0!", "#¡VALOR!", "#¡REF!", "#¿NOMBRE?", "#¡NUM!", "#N/A" } },
            { "it", new[] { "#NULLO!", "#DIV/0!", "#VALORE!", "#RIF!", "#NOME?", "#NUM!", "#N/D" } },
            { "pt", new[] { "#NULO!", "#DIV/0!", "#VALOR!", "#REF!", "#NOME?", "#NÚM!", "#N/D" } },
            { "nl", new[] { "#LEEG!", "#DEEL/0!", "#WAARDE!", "#VERW!", "#NAAM?", "#GETAL!", "#N/B" } },
        };

        public static string Boolean(bool value, CultureInfo language)
        {
            string[] texts;
            if (language != null && Booleans.TryGetValue(language.TwoLetterISOLanguageName, out texts)) return value ? texts[0] : texts[1];
            return value ? "TRUE" : "FALSE";
        }

        public static string Error(string code, CultureInfo language)
        {
            if (string.IsNullOrEmpty(code)) return string.Empty;
            string normalized = code.Trim().ToUpperInvariant();
            int index = Array.IndexOf(ErrorCodes, normalized);
            string[] texts;
            if (index >= 0 && language != null && Errors.TryGetValue(language.TwoLetterISOLanguageName, out texts)) return texts[index];
            return code.Trim();
        }
    }
}
