using System;
using System.Collections.Generic;
using System.Globalization;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Lit le texte d'une plage Word et sa mise en forme caractère par caractère, sans parcourir
    /// chaque caractère : on interroge la police de la plage entière ; si une propriété est
    /// hétérogène (wdUndefined), la plage est coupée en deux et chaque moitié est examinée
    /// (recherche dichotomique). Le coût est proportionnel au nombre de changements de
    /// mise en forme, pas à la longueur du texte. Lecture seule : aucune modification du document.
    /// </summary>
    public sealed class WordRunReader
    {
        private const int MaxDepth = 48;

        private readonly dynamic _document;
        private readonly Func<int, Rgb?> _themeLookup;

        /// <param name="document">Objet Word.Document (liaison tardive).</param>
        /// <param name="themeLookup">Résolution des couleurs de thème (peut être null).</param>
        public WordRunReader(object document, Func<int, Rgb?> themeLookup)
        {
            if (document == null) throw new ArgumentNullException("document");
            _document = document;
            _themeLookup = themeLookup;
        }

        /// <summary>Segments bruts (non nettoyés) de la plage [start, end[.</summary>
        public List<TextRun> Read(int start, int end)
        {
            var runs = new List<TextRun>();
            if (end > start) ReadSegment(start, end, runs, true);
            return runs;
        }

        private void ReadSegment(int start, int end, List<TextRun> output, bool allowParagraphSplit)
        {
            dynamic range = _document.Range(start, end);
            string text = WordCom.AsString(range.Text);

            if (text.Length == end - start)
            {
                // Correspondance exacte positions ↔ caractères : on peut découper librement.
                Bisect(start, end, text, start, output, 0);
                return;
            }

            // Champs, texte masqué, tableaux imbriqués… : les positions ne correspondent plus au texte.
            // On descend au niveau du paragraphe, puis on lit chaque paragraphe d'un bloc.
            if (allowParagraphSplit)
            {
                var parts = new List<int[]>();
                foreach (dynamic paragraph in range.Paragraphs)
                {
                    dynamic pr = paragraph.Range;
                    int ps = Math.Max(start, WordCom.AsInt(pr.Start));
                    int pe = Math.Min(end, WordCom.AsInt(pr.End));
                    if (pe > ps) parts.Add(new[] { ps, pe });
                }
                if (parts.Count > 1)
                {
                    foreach (var part in parts) ReadSegment(part[0], part[1], output, false);
                    return;
                }
            }

            FormatProbe probe = FormatProbe.Read((object)range, false);
            if (probe.HasUndefined)
            {
                // Propriétés hétérogènes non découpables : on complète avec celles du premier caractère.
                FormatProbe first = FormatProbe.Read((object)_document.Range(start, start + 1), false);
                probe.FillUndefinedFrom(first);
            }
            Emit(text, probe, (object)range, output);
        }

        private void Bisect(int start, int end, string text, int textOffset, List<TextRun> output, int depth)
        {
            dynamic range = _document.Range(start, end);
            bool leaf = end - start <= 1 || depth >= MaxDepth;
            FormatProbe probe = FormatProbe.Read((object)range, !leaf);

            if (!probe.HasUndefined || leaf)
            {
                Emit(text.Substring(start - textOffset, end - start), probe, (object)range, output);
                return;
            }

            int mid = start + (end - start) / 2;
            // Ne pas séparer les deux moitiés d'un caractère hors BMP (émoji, idéogramme rare…).
            if (char.IsLowSurrogate(text[mid - textOffset]) && mid - 1 > start) mid--;
            Bisect(start, mid, text, textOffset, output, depth + 1);
            Bisect(mid, end, text, textOffset, output, depth + 1);
        }

        private void Emit(string text, FormatProbe probe, object range, List<TextRun> output)
        {
            if (string.IsNullOrEmpty(text) || probe.Hidden) return;
            if (probe.AllCaps) text = text.ToUpper(CultureInfo.CurrentCulture);

            var format = new RunFormat
            {
                FontName = string.IsNullOrEmpty(probe.Name) ? null : probe.Name,
                Size = probe.Size > 0 && !WordCom.IsUndefined(probe.Size) ? probe.Size : (double?)null,
                Bold = probe.Bold,
                Italic = probe.Italic,
                Strike = probe.Strike,
                Underline = probe.Underline,
                Position = probe.Position,
                Color = ResolveColor(probe.Color, range)
            };

            var run = new TextRun(text, format, WordColor.FromHighlightIndex(probe.Highlight));
            var last = output.Count > 0 ? output[output.Count - 1] : null;
            if (last != null && last.HasSameFormatting(run)) last.Text += text;
            else output.Add(run);
        }

        private Rgb? ResolveColor(int color, object rangeObject)
        {
            dynamic range = rangeObject;
            if (color == WordColor.Automatic || WordCom.IsUndefined(color)) return null;
            if (WordColor.IsPlainRgb(color)) return Rgb.FromBgr(color);
            // Couleur de thème : Word 2007+ sait donner la couleur RVB effective.
            try
            {
                int rgb = WordCom.AsInt(range.Font.TextColor.RGB);
                if (WordColor.IsPlainRgb(rgb)) return Rgb.FromBgr(rgb);
            }
            catch (Exception)
            {
                // TextColor n'existe pas avant Word 2007.
            }
            return WordColor.Decode(color, _themeLookup);
        }

        /// <summary>Valeurs de police lues sur une plage ; wdUndefined signale une plage hétérogène.</summary>
        private sealed class FormatProbe
        {
            public string Name;
            public double Size;
            public bool Bold, Italic, Strike, AllCaps, Hidden;
            public UnderlineKind Underline;
            public VerticalPosition Position;
            public int Color;
            public int Highlight;
            public bool HasUndefined;

            private readonly HashSet<string> _undefined = new HashSet<string>(StringComparer.Ordinal);

            /// <param name="stopAtFirstUndefined">Arrête la lecture dès qu'une propriété est hétérogène (économise des appels COM).</param>
            public static FormatProbe Read(object rangeObject, bool stopAtFirstUndefined)
            {
                dynamic range = rangeObject;
                var p = new FormatProbe();
                dynamic font = range.Font;

                int v = WordCom.AsInt(font.Bold);
                if (p.Check("Bold", v) && stopAtFirstUndefined) return p;
                p.Bold = WordCom.IsTrue(v);

                v = WordCom.AsInt(font.Italic);
                if (p.Check("Italic", v) && stopAtFirstUndefined) return p;
                p.Italic = WordCom.IsTrue(v);

                p.Color = WordCom.AsInt(font.Color);
                if (p.Check("Color", p.Color) && stopAtFirstUndefined) return p;

                p.Size = WordCom.AsDouble(font.Size);
                if (p.Check("Size", WordCom.IsUndefined(p.Size) ? WordCom.Undefined : 0) && stopAtFirstUndefined) return p;

                p.Name = WordCom.AsString(font.Name);
                if (p.Check("Name", p.Name.Length == 0 ? WordCom.Undefined : 0) && stopAtFirstUndefined) return p;

                v = WordCom.AsInt(font.Underline);
                if (p.Check("Underline", v) && stopAtFirstUndefined) return p;
                p.Underline = v == 0 || WordCom.IsUndefined(v) ? UnderlineKind.None : (v == 3 || v == 43 ? UnderlineKind.Double : UnderlineKind.Single);

                p.Highlight = WordCom.AsInt(range.HighlightColorIndex);
                if (p.Check("Highlight", p.Highlight) && stopAtFirstUndefined) return p;

                int sup = WordCom.AsInt(font.Superscript);
                if (p.Check("Superscript", sup) && stopAtFirstUndefined) return p;
                int sub = WordCom.AsInt(font.Subscript);
                if (p.Check("Subscript", sub) && stopAtFirstUndefined) return p;
                p.Position = WordCom.IsTrue(sup) ? VerticalPosition.Superscript : WordCom.IsTrue(sub) ? VerticalPosition.Subscript : VerticalPosition.Baseline;

                int strike = WordCom.AsInt(font.StrikeThrough);
                if (p.Check("Strike", strike) && stopAtFirstUndefined) return p;
                int dstrike = WordCom.AsInt(font.DoubleStrikeThrough);
                if (p.Check("DoubleStrike", dstrike) && stopAtFirstUndefined) return p;
                p.Strike = WordCom.IsTrue(strike) || WordCom.IsTrue(dstrike);

                v = WordCom.AsInt(font.AllCaps);
                if (p.Check("AllCaps", v) && stopAtFirstUndefined) return p;
                p.AllCaps = WordCom.IsTrue(v);

                v = WordCom.AsInt(font.Hidden);
                if (p.Check("Hidden", v) && stopAtFirstUndefined) return p;
                p.Hidden = WordCom.IsTrue(v);

                return p;
            }

            private bool Check(string property, int value)
            {
                if (!WordCom.IsUndefined(value)) return false;
                HasUndefined = true;
                _undefined.Add(property);
                return true;
            }

            public void FillUndefinedFrom(FormatProbe other)
            {
                if (_undefined.Contains("Bold")) Bold = other.Bold;
                if (_undefined.Contains("Italic")) Italic = other.Italic;
                if (_undefined.Contains("Color")) Color = other.Color;
                if (_undefined.Contains("Size")) Size = other.Size;
                if (_undefined.Contains("Name")) Name = other.Name;
                if (_undefined.Contains("Underline")) Underline = other.Underline;
                if (_undefined.Contains("Highlight")) Highlight = other.Highlight;
                if (_undefined.Contains("Superscript") || _undefined.Contains("Subscript")) Position = other.Position;
                if (_undefined.Contains("Strike") || _undefined.Contains("DoubleStrike")) Strike = other.Strike;
                if (_undefined.Contains("AllCaps")) AllCaps = other.AllCaps;
                // Texte partiellement masqué dans un bloc indivisible : on le garde visible.
                if (_undefined.Contains("Hidden")) Hidden = false;
                _undefined.Clear();
                HasUndefined = false;
            }
        }
    }
}
