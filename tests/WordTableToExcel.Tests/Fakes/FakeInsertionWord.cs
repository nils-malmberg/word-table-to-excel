using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace WordTableToExcel.Tests.Fakes.Insertion
{
    /// <summary>
    /// Faux modèle objet de Word, réduit au texte d'un document : chaque paragraphe se termine par « \r » et chaque
    /// tableau occupe un caractère « \a ». Suffisant pour vérifier le choix du point d'insertion, la séparation des
    /// tableaux (Word fusionne deux tableaux accolés) et la création des légendes.
    /// </summary>
    public sealed class InsertionDocument
    {
        public const char TableChar = '\u0007';

        public readonly StringBuilder Text = new StringBuilder();
        public readonly List<string> InsertedXml = new List<string>();
        public readonly List<string> Highlighted = new List<string>();
        public int Captions;
        public int SelectionStart = -1;
        public bool FailXml;

        public InsertionDocument(string text)
        {
            Text.Append(text);
        }

        /// <summary>Texte lisible : « [T] » pour un tableau, « ¶ » pour une fin de paragraphe.</summary>
        public string Render()
        {
            return Text.ToString().Replace(TableChar.ToString(), "[T]").Replace("\r", "¶");
        }

        public void Insert(int position, string text)
        {
            if (position < 0 || position > Text.Length) throw new ArgumentOutOfRangeException("position");
            Text.Insert(position, text);
            // Word fusionne deux tableaux qui se touchent : cela ne doit jamais arriver.
            if (Text.ToString().Contains(new string(TableChar, 2))) throw new InvalidOperationException("Deux tableaux accolés : " + Render());
        }

        public bool InTable(int position)
        {
            return position >= 0 && position < Text.Length && Text[position] == TableChar;
        }

        public IEnumerable<int> TablePositions()
        {
            for (int i = 0; i < Text.Length; i++)
            {
                if (Text[i] == TableChar) yield return i;
            }
        }
    }

    public sealed class FakeInsertionApplication
    {
        public FakeInsertionApplication(InsertionDocument document, int selectionStart, int selectionEnd)
        {
            Document = document;
            Selection = new FakeSelection(new FakeRange(document, selectionStart, selectionEnd));
            CaptionLabels = new FakeCaptionLabels();
        }

        public InsertionDocument Document;
        public FakeSelection Selection { get; private set; }
        public FakeCaptionLabels CaptionLabels { get; private set; }
        public bool ScreenUpdating = true;
    }

    public sealed class FakeCaptionLabels
    {
        public object Item(object index)
        {
            return new { Name = "Tableau" };
        }
    }

    public sealed class FakeSelection
    {
        public FakeSelection(FakeRange range)
        {
            Range = range;
            Sections = new FakeSections();
        }

        public FakeRange Range { get; private set; }
        public FakeSections Sections { get; private set; }
    }

    public sealed class FakeSections
    {
        public object Item(int index)
        {
            return new { PageSetup = new FakePageSetup() };
        }
    }

    public sealed class FakePageSetup
    {
        public double PageWidth = 595.3;
        public double LeftMargin = 70.9;
        public double RightMargin = 70.9;
        public double Gutter = 0;
        public int GutterPos = 0;
        public object TextColumns = new { Count = 1, Width = 453.5 };
    }

    /// <summary>Document Word vu par <c>WordTableInserter</c> (champs uniquement).</summary>
    public sealed class FakeInsertionWordDocument
    {
        public readonly List<object> Fields = new List<object>();
    }

    public sealed class FakeRange
    {
        private readonly InsertionDocument _doc;

        public FakeRange(InsertionDocument doc, int start, int end)
        {
            _doc = doc;
            Start = start;
            End = end;
            Find = new FakeFind(this);
        }

        public int Start { get; set; }
        public int End { get; set; }
        public FakeFind Find { get; private set; }

        internal InsertionDocument Doc
        {
            get { return _doc; }
        }

        public string Text
        {
            get { return _doc.Text.ToString(Start, End - Start); }
        }

        public FakeRange Duplicate
        {
            get { return new FakeRange(_doc, Start, End); }
        }

        public void Collapse(int direction)
        {
            if (direction == 0) Start = End;
            else End = Start;
        }

        public void SetRange(int start, int end)
        {
            Start = start;
            End = end;
        }

        public void WholeStory()
        {
            Start = 0;
            End = _doc.Text.Length;
        }

        public object Information(int type)
        {
            if (type != 12) throw new NotSupportedException();
            return _doc.InTable(Start);
        }

        public FakeTables Tables
        {
            get
            {
                var list = new List<FakeTable>();
                foreach (var p in _doc.TablePositions())
                {
                    if ((p >= Start && p < Math.Max(End, Start + 1)) || (Start == End && p == Start)) list.Add(new FakeTable(_doc, p));
                }
                return new FakeTables(list);
            }
        }

        public FakeParagraphs Paragraphs
        {
            get
            {
                int p = Start;
                int start = p;
                while (start > 0 && _doc.Text[start - 1] != '\r' && _doc.Text[start - 1] != InsertionDocument.TableChar) start--;
                int end = p;
                while (end < _doc.Text.Length && _doc.Text[end] != '\r') end++;
                if (end < _doc.Text.Length) end++;
                return new FakeParagraphs(new FakeRange(_doc, start, end));
            }
        }

        public void InsertParagraphAfter()
        {
            _doc.Insert(End, "\r");
            End++;
        }

        public void InsertParagraphBefore()
        {
            _doc.Insert(Start, "\r");
            End++;
        }

        public void InsertBefore(string text)
        {
            _doc.Insert(Start, text);
            End += text.Length;
        }

        public void InsertAfter(string text)
        {
            _doc.Insert(End, text);
            End += text.Length;
        }

        public void InsertXML(string xml)
        {
            if (_doc.FailXml) throw new InvalidOperationException("InsertXML refusé");
            var parsed = XDocument.Parse(xml);
            if (!parsed.Descendants().Any(e => e.Name.LocalName == "tbl")) throw new InvalidOperationException("Pas de tableau");
            if (Start != End) throw new InvalidOperationException("La sélection serait remplacée");
            _doc.InsertedXml.Add(xml);
            _doc.Insert(Start, InsertionDocument.TableChar.ToString());
        }

        public void InsertFile(string FileName, bool ConfirmConversions, bool Link, bool Attachment)
        {
            if (!System.IO.File.Exists(FileName)) throw new InvalidOperationException("Fichier absent");
            _doc.InsertedXml.Add("docx:" + FileName);
            _doc.Insert(Start, InsertionDocument.TableChar.ToString());
        }

        public void InsertCaption(object Label, string Title, string TitleAutoText, int Position, int ExcludeLabel)
        {
            // Appelé sur la plage d'un tableau.
            _doc.Captions++;
            string text = "Tableau " + _doc.Captions + Title + "\r";
            if (Position == 0) _doc.Insert(Start, text);
            else _doc.Insert(End, text);
        }

        public int HighlightColorIndex
        {
            set { _doc.Highlighted.Add(Text); }
        }

        public void Select()
        {
            _doc.SelectionStart = Start;
        }
    }

    public sealed class FakeFind
    {
        private readonly FakeRange _range;

        public FakeFind(FakeRange range)
        {
            _range = range;
        }

        public void ClearFormatting()
        {
        }

        public bool Execute(string FindText, bool MatchCase, bool MatchWholeWord, bool MatchWildcards, bool Forward, int Wrap)
        {
            int index = _range.Doc.Text.ToString(_range.Start, _range.End - _range.Start).IndexOf(FindText, StringComparison.Ordinal);
            if (index < 0) return false;
            _range.SetRange(_range.Start + index, _range.Start + index + FindText.Length);
            return true;
        }
    }

    public sealed class FakeParagraphs
    {
        private readonly FakeRange _paragraph;

        public FakeParagraphs(FakeRange paragraph)
        {
            _paragraph = paragraph;
        }

        public object Item(int index)
        {
            return new { Range = _paragraph };
        }
    }

    public sealed class FakeTables : IEnumerable
    {
        private readonly List<FakeTable> _tables;

        public FakeTables(List<FakeTable> tables)
        {
            _tables = tables;
        }

        public int Count
        {
            get { return _tables.Count; }
        }

        public FakeTable Item(int index)
        {
            return _tables[index - 1];
        }

        public IEnumerator GetEnumerator()
        {
            return _tables.GetEnumerator();
        }
    }

    public sealed class FakeTable
    {
        private readonly InsertionDocument _doc;
        private readonly int _ordinal;

        public FakeTable(InsertionDocument doc, int position)
        {
            _doc = doc;
            // Comme dans Word, l'objet suit son tableau quand du texte est inséré avant lui.
            _ordinal = doc.TablePositions().TakeWhile(p => p < position).Count();
        }

        public FakeRange Range
        {
            get
            {
                int position = _doc.TablePositions().ElementAt(_ordinal);
                return new FakeRange(_doc, position, position + 1);
            }
        }
    }
}
