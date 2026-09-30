using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableToExcel.Tests.Fakes
{
    // Imitation minimale du modèle objet de Word, suffisante pour exercer la couche de lecture
    // (liaison tardive « dynamic ») sans Word : positions de caractères, polices hétérogènes
    // (wdUndefined), marques de fin de cellule, codes de champ masqués, paragraphes, tableaux.

    public sealed class CharFormat
    {
        public string Name = "Calibri";
        public float Size = 11;
        public bool Bold, Italic, Strike, DoubleStrike, Superscript, Subscript, AllCaps, Hidden;
        public int Underline;
        public int Color = -16777216;
        public int? ResolvedThemeRgb;
        public int Highlight;
        public int Alignment;
        public string Style = "Normal";

        public CharFormat Clone() => (CharFormat)MemberwiseClone();
    }

    public sealed class FakeChar
    {
        public char C;
        public CharFormat F;
        /// <summary>Présent dans les positions mais absent de Range.Text (code de champ masqué).</summary>
        public bool Invisible;
        /// <summary>Marque de fin de cellule / de ligne : une position, rendue « \r\a ».</summary>
        public bool CellMark;
    }

    public sealed class FakeDocument
    {
        public readonly List<FakeChar> Chars = new List<FakeChar>();
        public readonly List<FakeTable> TableList = new List<FakeTable>();
        public readonly List<FakeField> FieldList = new List<FakeField>();
        public readonly List<FakeRevision> RevisionList = new List<FakeRevision>();
        public int RangeCalls;
        public int FontCalls;
        public bool Saved = true;
        public string Name = "Rapport.docx";
        public string Path = "";
        public string FullName = "Rapport.docx";

        public FakeRange Range(int start, int end)
        {
            RangeCalls++;
            return new FakeRange(this, start, end);
        }

        public FakeRange Range() => Range(0, Chars.Count);
        public FakeRange Content => new FakeRange(this, 0, Chars.Count);
        public FakeTables Tables => new FakeTables(TableList.Where(t => t.NestingLevel == 1).ToList());
        public FakeStyles Styles => new FakeStyles();
        public FakeApplication Application => new FakeApplication();

        internal string TextOf(int start, int end)
        {
            var sb = new StringBuilder();
            for (int i = start; i < end && i < Chars.Count; i++)
            {
                var ch = Chars[i];
                if (ch.Invisible) continue;
                if (ch.CellMark) sb.Append("\r\a");
                else sb.Append(ch.C);
            }
            return sb.ToString();
        }

        internal IEnumerable<FakeChar> Span(int start, int end)
        {
            if (end <= start)
            {
                int i = Math.Min(Math.Max(0, start), Chars.Count - 1);
                if (i >= 0) yield return Chars[i];
                yield break;
            }
            for (int i = start; i < end && i < Chars.Count; i++) yield return Chars[i];
        }

        /// <summary>Paragraphes : se terminent par \r ou par une marque de cellule.</summary>
        internal List<int[]> ParagraphBounds()
        {
            var list = new List<int[]>();
            int start = 0;
            for (int i = 0; i < Chars.Count; i++)
            {
                if (Chars[i].C == '\r' || Chars[i].CellMark)
                {
                    list.Add(new[] { start, i + 1 });
                    start = i + 1;
                }
            }
            if (start < Chars.Count) list.Add(new[] { start, Chars.Count });
            return list;
        }
    }

    public sealed class FakeApplication
    {
        public FakeCaptionLabels CaptionLabels => new FakeCaptionLabels();
    }

    public sealed class FakeCaptionLabels
    {
        public FakeNamed Item(int index) => new FakeNamed { Name = "Tableau" };
    }

    public sealed class FakeNamed
    {
        public string Name;
        public string NameLocal;
    }

    public sealed class FakeStyles
    {
        public FakeNamed Item(int index) => new FakeNamed { NameLocal = "Légende" };
    }

    public sealed class FakeRange
    {
        private readonly FakeDocument _doc;

        public FakeRange(FakeDocument doc, int start, int end)
        {
            _doc = doc;
            Start = start;
            End = end;
        }

        public int Start { get; set; }
        public int End { get; set; }
        public string Text => _doc.TextOf(Start, End);
        public string WordOpenXML { get; set; }
        public string XML { get; set; }

        public FakeFont Font
        {
            get
            {
                _doc.FontCalls++;
                return new FakeFont(_doc.Span(Start, End).Select(c => c.F).ToList());
            }
        }

        public int HighlightColorIndex => Uniform(f => f.Highlight);

        public FakeParagraphFormat ParagraphFormat => new FakeParagraphFormat { Alignment = Uniform(f => f.Alignment) };

        public int Orientation => 0;

        public FakeParagraphs Paragraphs
        {
            get
            {
                var result = new List<FakeParagraph>();
                var all = _doc.ParagraphBounds();
                for (int i = 0; i < all.Count; i++)
                {
                    var b = all[i];
                    bool intersects = End > Start ? b[0] < End && b[1] > Start : b[0] <= Start && Start < b[1];
                    if (intersects) result.Add(new FakeParagraph(_doc, i));
                }
                return new FakeParagraphs(result);
            }
        }

        public FakeTables Tables => new FakeTables(_doc.TableList.Where(t => t.NestingLevel == 1 && t.StartPos < Math.Max(End, Start + 1) && t.EndPos > Start).ToList());

        public FakeFields Fields => new FakeFields(_doc.FieldList.Where(f => f.Position >= Start && f.Position < End).ToList());

        /// <summary>Révisions qui touchent la plage (comme Word, chacune avec sa plage complète).</summary>
        public FakeRevisions Revisions => new FakeRevisions(_doc.RevisionList.Where(r => r.Start < End && r.End > Start).ToList());

        public FakeDocument Document => _doc;

        /// <summary>Cellules d'un tableau (Range.Cells) : toutes les cellules visibles, imbriquées comprises.</summary>
        public FakeCells Cells
        {
            get
            {
                var cells = _doc.TableList.SelectMany(t => t.CellList).Where(c => c.StartPos >= Start && c.EndPos <= End)
                    .OrderBy(c => c.StartPos).ToList();
                return new FakeCells(cells);
            }
        }

        private int Uniform(Func<CharFormat, int> selector)
        {
            var values = _doc.Span(Start, End).Select(c => selector(c.F)).Distinct().ToList();
            return values.Count == 1 ? values[0] : 9999999;
        }
    }

    public sealed class FakeParagraphFormat
    {
        public int Alignment { get; set; }
    }

    public sealed class FakeFont
    {
        private readonly List<CharFormat> _formats;

        public FakeFont(List<CharFormat> formats)
        {
            _formats = formats;
        }

        private object U<T>(Func<CharFormat, T> selector, object undefined)
        {
            var values = _formats.Select(selector).Distinct().ToList();
            return values.Count == 1 ? values[0] : undefined;
        }

        private int B(Func<CharFormat, bool> selector)
        {
            var values = _formats.Select(selector).Distinct().ToList();
            return values.Count == 1 ? (values[0] ? -1 : 0) : 9999999;
        }

        public string Name => (string)U(f => f.Name, "");
        public float Size => Convert.ToSingle(U(f => f.Size, 9999999f));
        public int Bold => B(f => f.Bold);
        public int Italic => B(f => f.Italic);
        public int StrikeThrough => B(f => f.Strike);
        public int DoubleStrikeThrough => B(f => f.DoubleStrike);
        public int Superscript => B(f => f.Superscript);
        public int Subscript => B(f => f.Subscript);
        public int AllCaps => B(f => f.AllCaps);
        public int Hidden => B(f => f.Hidden);
        public int Underline => (int)U(f => f.Underline, 9999999);
        public int Color => (int)U(f => f.Color, 9999999);

        public FakeColorFormat TextColor
        {
            get
            {
                var resolved = _formats.Select(f => f.ResolvedThemeRgb).Distinct().ToList();
                if (resolved.Count == 1 && resolved[0].HasValue) return new FakeColorFormat { RGB = resolved[0].Value };
                throw new InvalidOperationException("TextColor indisponible");
            }
        }
    }

    public sealed class FakeColorFormat
    {
        public int RGB { get; set; }
    }

    public sealed class FakeParagraph
    {
        private readonly FakeDocument _doc;
        private readonly int _index;

        public FakeParagraph(FakeDocument doc, int index)
        {
            _doc = doc;
            _index = index;
        }

        private int[] Bounds => _doc.ParagraphBounds()[_index];
        public FakeRange Range => new FakeRange(_doc, Bounds[0], Bounds[1]);
        public int Alignment => _doc.Chars[Bounds[0]].F.Alignment;
        public FakeNamed Style => new FakeNamed { NameLocal = _doc.Chars[Bounds[0]].F.Style };

        public FakeParagraph Previous() => _index > 0 ? new FakeParagraph(_doc, _index - 1) : null;

        public FakeParagraph Next() => _index + 1 < _doc.ParagraphBounds().Count ? new FakeParagraph(_doc, _index + 1) : null;
    }

    public sealed class FakeParagraphs : IEnumerable
    {
        private readonly List<FakeParagraph> _items;

        public FakeParagraphs(List<FakeParagraph> items)
        {
            _items = items;
        }

        public int Count => _items.Count;
        public FakeParagraph Item(int index) => _items[index - 1];
        public IEnumerator GetEnumerator() => _items.GetEnumerator();
    }

    public sealed class FakeField
    {
        public FakeDocument Doc;
        public int Position;
        public int Type = 12;
        public string CodeText;
        public int ResultStart, ResultEnd;
        public FakeNamedText Code => new FakeNamedText { Text = CodeText };
        public FakeRange Result => new FakeRange(Doc, ResultStart, ResultEnd);
    }

    /// <summary>Révision du suivi des modifications (wdRevisionInsert = 1, wdRevisionDelete = 2…).</summary>
    public sealed class FakeRevision
    {
        public FakeDocument Doc;
        public int Type;
        public int Start, End;
        public FakeRange Range => new FakeRange(Doc, Start, End);
    }

    public sealed class FakeRevisions : IEnumerable
    {
        private readonly List<FakeRevision> _items;

        public FakeRevisions(List<FakeRevision> items)
        {
            _items = items;
        }

        public int Count => _items.Count;
        public IEnumerator GetEnumerator() => _items.GetEnumerator();
    }

    public sealed class FakeNamedText
    {
        public string Text;
    }

    public sealed class FakeFields : IEnumerable
    {
        private readonly List<FakeField> _items;

        public FakeFields(List<FakeField> items)
        {
            _items = items;
        }

        public int Count => _items.Count;
        public IEnumerator GetEnumerator() => _items.GetEnumerator();
    }

    public sealed class FakeTable
    {
        internal FakeDocument Doc;
        public int StartPos;
        public int EndPos;
        public int NestingLevel { get; set; } = 1;
        public string Xml;
        public string Xml2003;
        public readonly List<FakeCell> CellList = new List<FakeCell>();
        public readonly List<FakeTable> Nested = new List<FakeTable>();

        public FakeRange Range
        {
            get
            {
                var r = new FakeRange(Doc, StartPos, EndPos) { WordOpenXML = Xml, XML = Xml2003 };
                return r;
            }
        }

        public FakeTables Tables => new FakeTables(Nested);
    }

    public sealed class FakeTables : IEnumerable
    {
        private readonly List<FakeTable> _items;

        public FakeTables(List<FakeTable> items)
        {
            _items = items;
        }

        public int Count => _items.Count;
        public FakeTable Item(int index) => _items[index - 1];
        public IEnumerator GetEnumerator() => _items.GetEnumerator();
    }

    public sealed class FakeCell
    {
        internal FakeDocument Doc;
        public int StartPos;
        public int EndPos;
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public float Width { get; set; } = 72;
        public int NestingLevel { get; set; } = 1;
        public int VerticalAlignment { get; set; }
        public FakeShading Shading { get; set; } = new FakeShading();
        public FakeBorders Borders { get; set; } = new FakeBorders();
        public FakeRange Range => new FakeRange(Doc, StartPos, EndPos);
    }

    public sealed class FakeShading
    {
        public int BackgroundPatternColor { get; set; } = -16777216;
        public int ForegroundPatternColor { get; set; } = -16777216;
        public int Texture { get; set; }
    }

    public sealed class FakeBorders
    {
        public readonly Dictionary<int, FakeBorder> Map = new Dictionary<int, FakeBorder>();
        public FakeBorder Item(int index) => Map.TryGetValue(index, out var b) ? b : new FakeBorder();
    }

    public sealed class FakeBorder
    {
        public int LineStyle { get; set; }
        public int LineWidth { get; set; } = 4;
        public int Color { get; set; } = -16777216;
    }

    public sealed class FakeCells : IEnumerable
    {
        private readonly List<FakeCell> _items;

        public FakeCells(List<FakeCell> items)
        {
            _items = items;
        }

        public int Count => _items.Count;
        public FakeCell Item(int index) => _items[index - 1];
        public IEnumerator GetEnumerator() => _items.GetEnumerator();
    }

    /// <summary>Construction pratique d'un document factice.</summary>
    public sealed class FakeDocumentBuilder
    {
        public readonly FakeDocument Doc = new FakeDocument();

        public CharFormat Normal = new CharFormat();

        public FakeDocumentBuilder Text(string text, CharFormat format = null)
        {
            foreach (char c in text) Doc.Chars.Add(new FakeChar { C = c, F = (format ?? Normal).Clone() });
            return this;
        }

        /// <summary>Position du premier caractère de <paramref name="text"/> dans le document.</summary>
        public int PositionOf(string text)
        {
            int index = new string(Doc.Chars.Select(c => c.C).ToArray()).IndexOf(text, StringComparison.Ordinal);
            if (index < 0) throw new ArgumentException("Texte absent du document : " + text);
            return index;
        }

        /// <summary>Marque <paramref name="text"/> comme révision (supprimé en suivi des modifications par défaut).</summary>
        public FakeDocumentBuilder Revision(string text, int type = 2)
        {
            int start = PositionOf(text);
            Doc.RevisionList.Add(new FakeRevision { Doc = Doc, Type = type, Start = start, End = start + text.Length });
            return this;
        }

        /// <summary>Marque <paramref name="text"/> comme résultat d'un renvoi vers une note (champ NOTEREF).</summary>
        public FakeDocumentBuilder NoteReference(string text)
        {
            int start = PositionOf(text);
            Doc.FieldList.Add(new FakeField { Doc = Doc, Position = start, Type = 72, CodeText = " NOTEREF _Ref1 \\h ", ResultStart = start, ResultEnd = start + text.Length });
            return this;
        }

        public FakeDocumentBuilder HiddenCode(string code)
        {
            foreach (char c in code) Doc.Chars.Add(new FakeChar { C = c, F = Normal.Clone(), Invisible = true });
            return this;
        }

        public FakeDocumentBuilder Paragraph(string text, CharFormat format = null, string seqCode = null)
        {
            if (seqCode != null) Doc.FieldList.Add(new FakeField { Position = Doc.Chars.Count, CodeText = seqCode });
            Text(text, format);
            return Text("\r", format);
        }

        /// <summary>
        /// Ajoute un tableau. <paramref name="rows"/> : pour chaque ligne, les cellules visibles
        /// (les continuations de fusion verticale sont absentes, comme dans Word), chaque cellule
        /// étant une liste de segments (texte, format).
        /// </summary>
        public FakeTable Table(IList<IList<FakeCellSpec>> rows, string xml = null)
        {
            var table = new FakeTable { Doc = Doc, StartPos = Doc.Chars.Count, Xml = xml };
            for (int r = 0; r < rows.Count; r++)
            {
                foreach (var spec in rows[r])
                {
                    var cell = new FakeCell { Doc = Doc, RowIndex = r + 1, ColumnIndex = spec.ColumnIndex, Width = spec.Width, StartPos = Doc.Chars.Count };
                    foreach (var run in spec.Runs) Text(run.Item1, run.Item2);
                    Doc.Chars.Add(new FakeChar { C = '\a', F = Normal.Clone(), CellMark = true });
                    cell.EndPos = Doc.Chars.Count;
                    if (spec.Configure != null) spec.Configure(cell);
                    table.CellList.Add(cell);
                }
                Doc.Chars.Add(new FakeChar { C = '\a', F = Normal.Clone(), CellMark = true }); // fin de ligne
            }
            table.EndPos = Doc.Chars.Count;
            Doc.TableList.Add(table);
            return table;
        }
    }

    public sealed class FakeCellSpec
    {
        public int ColumnIndex;
        public float Width = 72;
        public List<Tuple<string, CharFormat>> Runs = new List<Tuple<string, CharFormat>>();
        public Action<FakeCell> Configure;

        public static FakeCellSpec Of(int columnIndex, string text, CharFormat format = null, float width = 72)
        {
            var spec = new FakeCellSpec { ColumnIndex = columnIndex, Width = width };
            if (!string.IsNullOrEmpty(text)) spec.Runs.Add(Tuple.Create(text, format ?? new CharFormat()));
            return spec;
        }

        public FakeCellSpec And(string text, CharFormat format)
        {
            Runs.Add(Tuple.Create(text, format));
            return this;
        }
    }
}
