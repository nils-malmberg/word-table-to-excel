using System.Globalization;
using System.Text;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Point d'insertion lu dans Word (document actif, position du curseur), pour montrer à l'utilisateur
    /// où les tableaux importés seront placés avant qu'il ne lance l'import.
    /// </summary>
    internal sealed class InsertionInfo
    {
        public string DocumentName;
        /// <summary>Page du curseur ; 0 si inconnue.</summary>
        public int Page;
        public bool InTable;
        public bool HasSelection;
        /// <summary>Faux si le curseur est dans un en-tête, un pied de page, une note ou une zone de texte.</summary>
        public bool MainStory = true;
        /// <summary>Texte du paragraphe où se trouve le curseur (null : inconnu).</summary>
        public string ParagraphText;
        /// <summary>Curseur au tout début du paragraphe.</summary>
        public bool AtParagraphStart;
        /// <summary>Raison qui empêche l'import (mode protégé, document protégé…) ; null si l'import est possible.</summary>
        public string Problem;

        public bool CanImport
        {
            get { return Problem == null; }
        }

        private const int ExcerptLength = 60;

        /// <summary>Texte affiché dans la fenêtre ; <paramref name="info"/> null : aucun document ouvert dans Word.</summary>
        public static string Describe(InsertionInfo info, bool wordRunning)
        {
            if (info == null)
            {
                return wordRunning
                    ? "Aucun document n'est ouvert dans Word. Ouvrez ou créez un document, puis cliquez à l'endroit où insérer le tableau."
                    : "Word n'est pas ouvert. Ouvrez votre document dans Word, puis cliquez à l'endroit où insérer le tableau.";
            }
            if (info.Problem != null) return info.Problem;

            var sb = new StringBuilder();
            sb.Append("Insertion dans « ").Append(info.DocumentName).Append(" »");
            if (info.Page > 0) sb.Append(", page ").Append(info.Page.ToString(CultureInfo.CurrentCulture));
            sb.Append(" : ");

            string paragraph = Excerpt(info.ParagraphText);
            if (info.InTable)
            {
                sb.Append("juste après le tableau où se trouve le curseur.");
            }
            else if (info.ParagraphText == null)
            {
                sb.Append("à l'emplacement du curseur.");
            }
            else if (paragraph.Length == 0)
            {
                sb.Append("sur la ligne vide où se trouve le curseur.");
            }
            else if (info.AtParagraphStart && !info.HasSelection)
            {
                sb.Append("juste avant le paragraphe « ").Append(paragraph).Append(" ».");
            }
            else
            {
                sb.Append("juste après le paragraphe « ").Append(paragraph).Append(" ».");
            }

            if (info.HasSelection) sb.Append("\nLe texte sélectionné dans Word est conservé.");
            if (!info.MainStory) sb.Append("\nAttention : le curseur est dans un en-tête, un pied de page, une note ou une zone de texte.");
            return sb.ToString();
        }

        /// <summary>Début du paragraphe, sur une ligne, sans caractères de contrôle de Word.</summary>
        public static string Excerpt(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            bool space = false;
            foreach (char c in text)
            {
                if (char.IsControl(c) || char.IsWhiteSpace(c))
                {
                    space = sb.Length > 0;
                    continue;
                }
                if (space) sb.Append(' ');
                space = false;
                sb.Append(c);
            }
            string result = sb.ToString();
            if (result.Length <= ExcerptLength) return result;
            int cut = ExcerptLength - 1;
            if (char.IsHighSurrogate(result[cut - 1])) cut--;
            int lastSpace = result.LastIndexOf(' ', cut - 1, cut - 1);
            if (lastSpace > ExcerptLength / 2) cut = lastSpace;
            return result.Substring(0, cut).TrimEnd() + "…";
        }
    }
}
