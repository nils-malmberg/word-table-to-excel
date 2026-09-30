using System;
using WordTableToExcel.Core.Captions;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Recherche, pour un tableau, le paragraphe de légende situé juste au-dessus et juste
    /// au-dessous (en ignorant les paragraphes vides). Un paragraphe est une légende de tableau
    /// s'il contient un champ SEQ « Tableau / Table / Tabla… », ou si son texte commence par
    /// « Tabl… » (ou l'équivalent dans une autre langue).
    /// </summary>
    public sealed class WordCaptionScanner
    {
        /// <summary>wdFieldSequence.</summary>
        private const int FieldSequence = 12;
        /// <summary>wdStyleCaption.</summary>
        private const int StyleCaption = -35;
        /// <summary>Nombre maximal de paragraphes vides ignorés entre le tableau et sa légende.</summary>
        private const int MaxEmptyParagraphs = 3;

        private readonly dynamic _document;
        private readonly CaptionMatcher _matcher;
        private readonly Action<string> _log;
        private string _captionStyleName;
        private bool _captionStyleResolved;

        public WordCaptionScanner(object document, CaptionMatcher matcher, Action<string> log)
        {
            if (document == null) throw new ArgumentNullException("document");
            _document = document;
            _matcher = matcher ?? new CaptionMatcher();
            _log = log ?? (s => { });
        }

        public TableCaptionContext Scan(object table, int tableIndex)
        {
            dynamic t = table;
            dynamic range = t.Range;
            int start = WordCom.AsInt(range.Start);
            int end = WordCom.AsInt(range.End);

            var context = new TableCaptionContext { TableIndex = tableIndex };
            context.Above = SafeFind(start, true, tableIndex);
            context.Below = SafeFind(end, false, tableIndex);
            return context;
        }

        private CaptionCandidate SafeFind(int boundary, bool above, int tableIndex)
        {
            try
            {
                return Find(boundary, above);
            }
            catch (Exception ex)
            {
                _log(string.Format("Légende {0} du tableau {1} : {2}", above ? "au-dessus" : "au-dessous", tableIndex, ex.Message));
                return null;
            }
        }

        private CaptionCandidate Find(int boundary, bool above)
        {
            dynamic paragraph;
            if (above)
            {
                if (boundary <= 0) return null;
                paragraph = _document.Range(boundary - 1, boundary - 1).Paragraphs.Item(1);
            }
            else
            {
                int documentEnd = WordCom.AsInt(_document.Content.End);
                if (boundary >= documentEnd) return null;
                paragraph = _document.Range(boundary, boundary).Paragraphs.Item(1);
            }

            for (int step = 0; step <= MaxEmptyParagraphs && paragraph != null; step++)
            {
                dynamic range = paragraph.Range;
                if (WordCom.AsInt(range.Tables.Count) > 0) return null; // on est arrivé dans un autre tableau

                // Texte sans les passages supprimés en suivi des modifications : une légende supprimée n'en est plus une.
                string text = WordRevisions.VisibleText((object)range);
                string clean = CaptionMatcher.CleanCaptionText(text);
                if (clean.Length == 0)
                {
                    paragraph = above ? paragraph.Previous() : paragraph.Next();
                    continue;
                }
                return Evaluate((object)paragraph, (object)range, clean);
            }
            return null;
        }

        private CaptionCandidate Evaluate(object paragraphObject, object rangeObject, string text)
        {
            dynamic paragraph = paragraphObject;
            dynamic range = rangeObject;
            int key = WordCom.AsInt(range.Start);

            bool tableSequence = false;
            bool otherSequence = false;
            if (WordCom.AsInt(range.Fields.Count) > 0)
            {
                foreach (dynamic field in range.Fields)
                {
                    if (WordCom.AsInt(field.Type) != FieldSequence) continue;
                    string code = WordCom.AsString(field.Code.Text);
                    if (_matcher.IsTableSequenceField(code)) tableSequence = true;
                    else otherSequence = true;
                }
            }

            if (tableSequence) return new CaptionCandidate(key, text, true);
            if (otherSequence) return null; // légende de figure, d'équation…

            return _matcher.LooksLikeCaptionText(text, HasCaptionStyle(paragraph)) ? new CaptionCandidate(key, text, false) : null;
        }

        private bool HasCaptionStyle(dynamic paragraph)
        {
            try
            {
                if (!_captionStyleResolved)
                {
                    _captionStyleResolved = true;
                    _captionStyleName = WordCom.AsString(_document.Styles.Item(StyleCaption).NameLocal);
                }
                string name = WordCom.AsString(paragraph.Style.NameLocal);
                return name.Length > 0 && (string.Equals(name, _captionStyleName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "Caption", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "Légende", StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
