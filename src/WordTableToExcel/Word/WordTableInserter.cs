using System;
using System.Globalization;
using System.IO;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.Word
{
    /// <summary>Légende à créer pour un tableau importé.</summary>
    internal sealed class CaptionRequest
    {
        /// <summary>Texte après le numéro (« Ventes 2024 ») ; vide = texte à compléter.</summary>
        public string Title;
        public bool Below;
    }

    /// <summary>
    /// Insère des tableaux dans le document Word, au point d'insertion, sans jamais écraser de contenu :
    /// insertion XML en un seul bloc (Word 2007 et suivants), avec repli par insertion d'un document .docx
    /// temporaire ; légende « Tableau N » créée par la fonction de Word (champ SEQ, style Légende).
    /// </summary>
    internal sealed class WordTableInserter
    {
        private const int WdCollapseEnd = 0;
        private const int WdWithInTable = 12;
        private const int WdCaptionTable = -2;
        private const int WdCaptionPositionAbove = 0;
        private const int WdCaptionPositionBelow = 1;
        private const int WdYellow = 7;
        private const int WdFieldSequence = 12;

        public const string Placeholder = "[Titre du tableau]";

        private readonly dynamic _application;
        private readonly dynamic _document;
        private readonly int _wordVersion;
        /// <summary>Plage de référence dans l'histoire (corps, en-tête, zone de texte…) où se trouve le curseur.</summary>
        private dynamic _anchor;
        private int _position = -1;
        private string _captionLabel;

        public WordTableInserter(object application, object document, int wordMajorVersion)
        {
            _application = application;
            _document = document;
            _wordVersion = wordMajorVersion;
        }

        /// <summary>Nom du libellé de légende des tableaux dans la langue de Word (« Tableau », « Table »…).</summary>
        public string CaptionLabel
        {
            get
            {
                if (_captionLabel == null)
                {
                    try
                    {
                        _captionLabel = WordCom.AsString(_application.CaptionLabels.Item(WdCaptionTable).Name);
                    }
                    catch (Exception ex)
                    {
                        Log.Info("Libellé de légende introuvable : " + ex.Message);
                    }
                    if (string.IsNullOrEmpty(_captionLabel)) _captionLabel = "Tableau";
                }
                return _captionLabel;
            }
        }

        /// <summary>Largeur utile (points) de la page à l'endroit de l'insertion : marges et colonnes de texte déduites.</summary>
        public double AvailableWidth()
        {
            try
            {
                dynamic setup = _application.Selection.Sections.Item(1).PageSetup;
                double width = WordCom.AsDouble(setup.PageWidth) - WordCom.AsDouble(setup.LeftMargin) - WordCom.AsDouble(setup.RightMargin);
                try
                {
                    if (!WordCom.IsTrue(setup.GutterPos)) width -= WordCom.AsDouble(setup.Gutter); // gutter à gauche (wdGutterPosLeft = 0)
                }
                catch (Exception)
                {
                    // Propriété absente des anciennes versions.
                }
                try
                {
                    dynamic columns = setup.TextColumns;
                    if (WordCom.AsInt(columns.Count) > 1) width = WordCom.AsDouble(columns.Width);
                }
                catch (Exception)
                {
                    // Une seule colonne.
                }
                if (width > 36 && width < 5000 && !WordCom.IsUndefined(width)) return width;
            }
            catch (Exception ex)
            {
                Log.Info("Largeur de page inconnue : " + ex.Message);
            }
            return 453.5; // A4, marges de 2,5 cm
        }

        /// <summary>
        /// Insère le tableau et sa légende éventuelle ; les tableaux suivants sont placés à la suite, séparés
        /// par un paragraphe vide (deux tableaux accolés seraient fusionnés par Word).
        /// </summary>
        public void Insert(TableModel table, CaptionRequest caption)
        {
            dynamic range = NextInsertionRange();
            int start = WordCom.AsInt(range.Start);

            string method = InsertContent(range, table);
            dynamic inserted = FindInsertedTable(start);
            if (inserted == null) throw new InvalidOperationException("Le tableau inséré est introuvable dans le document.");
            Log.Info("Tableau inséré (" + method + ") : " + table.RowCount + " × " + table.ColumnCount + ".");

            int end = WordCom.AsInt(inserted.Range.End);
            if (caption != null)
            {
                end = AddCaption(inserted, caption, end);
            }
            _position = end;
        }

        /// <summary>Place le curseur après le dernier tableau inséré.</summary>
        public void SelectEnd()
        {
            if (_position < 0) return;
            try
            {
                At(_position).Select();
            }
            catch (Exception ex)
            {
                Log.Info("Sélection finale impossible : " + ex.Message);
            }
        }

        // ------------------------------------------------------------------ point d'insertion

        /// <summary>Plage vide à la position donnée, dans l'histoire du point d'insertion.</summary>
        private dynamic At(int position)
        {
            dynamic r = _anchor.Duplicate;
            r.SetRange(position, position);
            return r;
        }

        private bool InTable(int position)
        {
            try
            {
                return position >= 0 && WordCom.IsTrue(At(position).Information(WdWithInTable));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private dynamic NextInsertionRange()
        {
            dynamic range;
            if (_anchor == null)
            {
                range = _application.Selection.Range.Duplicate;
                range.Collapse(WdCollapseEnd); // ne jamais remplacer le texte sélectionné
                _anchor = range.Duplicate;
            }
            else
            {
                range = At(_position);
            }

            // Dans un tableau : on se place après le tableau (le plus extérieur).
            for (int guard = 0; guard < 16 && WordCom.IsTrue(range.Information(WdWithInTable)); guard++)
            {
                int after = WordCom.AsInt(range.Tables.Item(1).Range.End);
                range = At(after);
            }

            dynamic paragraph = range.Paragraphs.Item(1).Range;
            int paragraphStart = WordCom.AsInt(paragraph.Start);
            int paragraphEnd = WordCom.AsInt(paragraph.End);
            bool empty = WordCom.AsString(paragraph.Text).TrimEnd('\r', '\a', '\n').Length == 0;

            int position;
            if (empty || WordCom.AsInt(range.Start) == paragraphStart)
            {
                // Paragraphe vide : le tableau s'y place et le paragraphe reste après lui.
                // Début d'un paragraphe non vide : le tableau se place juste avant.
                position = paragraphStart;
            }
            else
            {
                // Milieu ou fin d'un paragraphe : nouveau paragraphe vide après lui.
                paragraph.InsertParagraphAfter();
                position = paragraphEnd;
            }

            // Deux tableaux accolés seraient fusionnés par Word : un paragraphe vide les sépare
            // (et sépare aussi deux tableaux importés à la suite).
            if (InTable(position - 1) || (_position >= 0 && position == _position))
            {
                At(position).InsertBefore("\r");
                position++;
            }
            return At(position);
        }

        // ------------------------------------------------------------------ insertion

        private string InsertContent(dynamic range, TableModel table)
        {
            Exception xmlError = null;
            if (_wordVersion >= 12)
            {
                try
                {
                    range.InsertXML(WordTableXmlWriter.BuildFlatOpc(table));
                    return "XML";
                }
                catch (Exception ex)
                {
                    xmlError = ex;
                    Log.Error("Insertion XML impossible, essai par fichier temporaire", ex);
                }
            }

            string temp = Path.Combine(Path.GetTempPath(), "WordTableToExcel-" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                File.WriteAllBytes(temp, WordTableXmlWriter.BuildDocx(table));
                range.InsertFile(FileName: temp, ConfirmConversions: false, Link: false, Attachment: false);
                return "fichier";
            }
            catch (Exception ex)
            {
                Log.Error("Insertion par fichier temporaire impossible", ex);
                if (_wordVersion > 0 && _wordVersion < 12)
                {
                    throw new InvalidOperationException("Cette version de Word ne sait pas lire les tableaux au format Word 2007. "
                        + "Installez le « Pack de compatibilité Microsoft Office » ou utilisez Word 2007 ou une version ultérieure.", ex);
                }
                throw new InvalidOperationException("Word a refusé l'insertion du tableau : " + (xmlError ?? ex).Message, xmlError ?? ex);
            }
            finally
            {
                try
                {
                    if (File.Exists(temp)) File.Delete(temp);
                }
                catch (Exception)
                {
                    // Fichier temporaire : sans conséquence.
                }
            }
        }

        private dynamic FindInsertedTable(int start)
        {
            try
            {
                dynamic probe = At(start);
                if (WordCom.IsTrue(probe.Information(WdWithInTable))) return probe.Tables.Item(1);
            }
            catch (Exception ex)
            {
                Log.Info("Recherche directe du tableau : " + ex.Message);
            }

            // Premier tableau de l'histoire qui commence à partir de la position d'insertion.
            dynamic story = _anchor.Duplicate;
            story.WholeStory();
            dynamic best = null;
            int bestStart = int.MaxValue;
            foreach (dynamic table in story.Tables)
            {
                int s = WordCom.AsInt(table.Range.Start);
                if (s >= start && s < bestStart)
                {
                    best = table;
                    bestStart = s;
                }
            }
            return best;
        }

        // ------------------------------------------------------------------ légende

        /// <summary>Crée la légende (fonction « Insérer une légende » de Word) ; renvoie la position de fin du bloc tableau + légende.</summary>
        private int AddCaption(dynamic table, CaptionRequest caption, int tableEnd)
        {
            string title = string.IsNullOrEmpty(caption.Title) ? null : caption.Title.Trim();
            string separator = CaptionMatcherIsFrench(CaptionLabel) ? "\u00A0: " : ": ";
            string text = separator + (title ?? Placeholder);
            try
            {
                table.Range.InsertCaption(Label: WdCaptionTable, Title: text, TitleAutoText: "",
                    Position: caption.Below ? WdCaptionPositionBelow : WdCaptionPositionAbove, ExcludeLabel: 0);
            }
            catch (Exception ex)
            {
                // Libellé de légende indisponible : légende écrite à la main (même résultat, style Légende et champ SEQ).
                Log.Error("InsertCaption", ex);
                InsertCaptionManually(table, caption.Below, text);
            }

            dynamic paragraph = CaptionParagraph(table, caption.Below);
            if (paragraph != null && title == null) Highlight(paragraph, Placeholder);
            int end = WordCom.AsInt(table.Range.End);
            if (caption.Below && paragraph != null) end = WordCom.AsInt(paragraph.End);
            return Math.Max(end, tableEnd);
        }

        private static bool CaptionMatcherIsFrench(string label)
        {
            return Core.Captions.CaptionMatcher.Normalize(label).StartsWith("tableau", StringComparison.Ordinal);
        }

        private dynamic CaptionParagraph(dynamic table, bool below)
        {
            try
            {
                if (below)
                {
                    int end = WordCom.AsInt(table.Range.End);
                    return At(end).Paragraphs.Item(1).Range;
                }
                int start = WordCom.AsInt(table.Range.Start);
                if (start == 0) return null;
                return At(start - 1).Paragraphs.Item(1).Range;
            }
            catch (Exception ex)
            {
                Log.Info("Paragraphe de légende introuvable : " + ex.Message);
                return null;
            }
        }

        private void InsertCaptionManually(dynamic table, bool below, string text)
        {
            const int WdStyleCaption = -35;
            const int WdFieldEmpty = -1;
            int position;
            if (below)
            {
                position = WordCom.AsInt(table.Range.End);
                At(position).InsertParagraphBefore();
            }
            else
            {
                position = WordCom.AsInt(table.Range.Start);
                At(position).InsertBefore("\r");
            }
            dynamic paragraph = At(position).Paragraphs.Item(1).Range;
            paragraph.Style = WdStyleCaption;
            // Libellé, champ SEQ, puis texte, écrits de la fin vers le début du paragraphe.
            dynamic start = At(WordCom.AsInt(paragraph.Start));
            start.InsertAfter(text);
            dynamic fieldRange = At(WordCom.AsInt(paragraph.Start));
            dynamic field = fieldRange.Fields.Add(fieldRange, WdFieldEmpty, "SEQ " + QuoteLabel(CaptionLabel) + " \\* ARABIC", false);
            At(WordCom.AsInt(paragraph.Start)).InsertBefore(CaptionLabel + " ");
            field.Update();
        }

        private static string QuoteLabel(string label)
        {
            return label.IndexOf(' ') >= 0 ? "\"" + label + "\"" : label;
        }

        private static void Highlight(dynamic paragraph, string text)
        {
            try
            {
                dynamic r = paragraph.Duplicate;
                dynamic find = r.Find;
                find.ClearFormatting();
                bool found = WordCom.IsTrue(find.Execute(FindText: text, MatchCase: true, MatchWholeWord: false, MatchWildcards: false, Forward: true, Wrap: 0));
                if (found) r.HighlightColorIndex = WdYellow;
            }
            catch (Exception ex)
            {
                Log.Info("Surlignage de la légende impossible : " + ex.Message);
            }
        }

        /// <summary>Met à jour la numérotation des légendes de tableau (champs SEQ) après l'insertion.</summary>
        public void UpdateCaptionNumbers()
        {
            try
            {
                string label = CaptionLabel;
                foreach (dynamic field in _document.Fields)
                {
                    if (WordCom.AsInt(field.Type) != WdFieldSequence) continue;
                    string code = WordCom.AsString(field.Code.Text);
                    string id = Core.Captions.CaptionMatcher.ParseSequenceIdentifier(code);
                    if (id != null && string.Equals(id, label, StringComparison.CurrentCultureIgnoreCase)) field.Update();
                }
            }
            catch (Exception ex)
            {
                Log.Info("Mise à jour des numéros de légende : " + ex.Message);
            }
        }

        internal static string Describe(int count)
        {
            return count.ToString(CultureInfo.CurrentCulture) + (count > 1 ? " tableaux" : " tableau");
        }
    }
}
