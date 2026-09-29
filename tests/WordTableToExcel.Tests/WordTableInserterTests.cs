using WordTableToExcel.Core.Model;
using WordTableToExcel.Tests.Fakes.Insertion;
using WordTableToExcel.Word;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class WordTableInserterTests
    {
        private static InsertionDocument Run(string text, int selectionStart, int selectionEnd, params CaptionRequest[] captions)
        {
            var doc = new InsertionDocument(text);
            var app = new FakeInsertionApplication(doc, selectionStart, selectionEnd);
            var inserter = new WordTableInserter(app, new FakeInsertionWordDocument(), 16);
            foreach (var caption in captions) inserter.Insert(XlsxWriterTests.SampleTable(), caption);
            inserter.UpdateCaptionNumbers();
            inserter.SelectEnd();
            return doc;
        }

        private static readonly CaptionRequest NoCaption = null;

        [Fact]
        public void EmptyParagraphReceivesTheTable()
        {
            var doc = Run("Intro\r\rSuite\r", 6, 6, NoCaption);
            Assert.Equal("Intro¶[T]¶Suite¶", doc.Render());
            Assert.Single(doc.InsertedXml);
        }

        [Fact]
        public void StartOfParagraphInsertsBefore()
        {
            Assert.Equal("Intro¶[T]Suite¶", Run("Intro\rSuite\r", 6, 6, NoCaption).Render());
        }

        [Fact]
        public void MiddleOfParagraphInsertsAfterItWithoutSplittingIt()
        {
            Assert.Equal("Bonjour tout le monde¶[T]¶", Run("Bonjour tout le monde\r", 7, 7, NoCaption).Render());
        }

        [Fact]
        public void SelectedTextIsNeverReplaced()
        {
            var doc = Run("Bonjour tout le monde\r", 8, 12, NoCaption);
            Assert.Equal("Bonjour tout le monde¶[T]¶", doc.Render());
        }

        [Fact]
        public void CursorInsideATableInsertsAfterItWithASeparator()
        {
            var doc = Run("A\r\u0007B\r", 2, 2, NoCaption);
            Assert.Equal("A¶[T]¶[T]B¶", doc.Render());
        }

        [Fact]
        public void EmptyParagraphRightAfterATableGetsASeparator()
        {
            Assert.Equal("[T]¶[T]¶", Run("\u0007\r", 1, 1, NoCaption).Render());
        }

        [Fact]
        public void SeveralTablesAreSeparatedAndCaptionsNumbered()
        {
            var above = new CaptionRequest { Title = "Ventes", Below = false };
            var toComplete = new CaptionRequest { Title = null, Below = false };
            var doc = Run("\r", 0, 0, above, toComplete, NoCaption);
            Assert.Equal("Tableau 1\u00A0: Ventes¶[T]¶Tableau 2\u00A0: [Titre du tableau]¶[T]¶[T]¶", doc.Render());
            Assert.Equal(new[] { "[Titre du tableau]" }, doc.Highlighted);
            Assert.Equal(doc.Text.Length - 1, doc.SelectionStart);
        }

        [Fact]
        public void CaptionsBelowTables()
        {
            var below = new CaptionRequest { Title = "Résultats", Below = true };
            var doc = Run("Texte\r", 5, 5, below, below);
            Assert.Equal("Texte¶[T]Tableau 1\u00A0: Résultats¶¶[T]Tableau 2\u00A0: Résultats¶¶", doc.Render());
        }

        [Fact]
        public void FallsBackToTemporaryDocxWhenXmlInsertionFails()
        {
            var doc = new InsertionDocument("\r") { FailXml = true };
            var app = new FakeInsertionApplication(doc, 0, 0);
            new WordTableInserter(app, new FakeInsertionWordDocument(), 16).Insert(XlsxWriterTests.SampleTable(), null);
            Assert.Equal("[T]¶", doc.Render());
            Assert.StartsWith("docx:", doc.InsertedXml[0]);
        }

        [Fact]
        public void AvailableWidthComesFromThePageSetup()
        {
            var app = new FakeInsertionApplication(new InsertionDocument("\r"), 0, 0);
            var inserter = new WordTableInserter(app, new FakeInsertionWordDocument(), 16);
            Assert.Equal(595.3 - 2 * 70.9, inserter.AvailableWidth(), 3);
            Assert.Equal("Tableau", inserter.CaptionLabel);
        }
    }
}
