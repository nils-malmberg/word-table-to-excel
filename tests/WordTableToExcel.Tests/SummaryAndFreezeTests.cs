using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Xlsx;
using Xunit;

namespace WordTableToExcel.Tests
{
    /// <summary>Feuille « Sommaire » en tête du classeur et lignes d'en-tête figées.</summary>
    public class SummaryAndFreezeTests
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static ZipArchive Write(XlsxExportOptions options, params TableModel[] tables)
        {
            var writer = new XlsxWorkbookWriter(options);
            foreach (var t in tables) writer.AddTable(t);
            var ms = new MemoryStream();
            writer.Save(ms);
            ms.Position = 0;
            return new ZipArchive(ms, ZipArchiveMode.Read);
        }

        private static XDocument Part(ZipArchive zip, string name)
        {
            var entry = zip.GetEntry(name);
            Assert.NotNull(entry);
            using (var s = entry.Open()) return XDocument.Load(s);
        }

        private static TableModel Table(int index, string sheetName, string caption, int startPage, int endPage)
        {
            var table = XlsxWriterTests.SampleTable(sheetName, caption);
            table.DocumentIndex = index;
            table.StartPage = startPage;
            table.EndPage = endPage;
            return table;
        }

        /// <summary>Texte de chaque cellule de la feuille (chaînes partagées résolues), par référence.</summary>
        private static System.Collections.Generic.Dictionary<string, string> Cells(ZipArchive zip, string sheet)
        {
            var strings = Part(zip, "xl/sharedStrings.xml").Root.Elements(S + "si")
                .Select(si => string.Concat(si.Descendants(S + "t").Select(t => t.Value))).ToList();
            return Part(zip, sheet).Descendants(S + "c").Where(c => c.Element(S + "v") != null).ToDictionary(
                c => (string)c.Attribute("r"),
                c => (string)c.Attribute("t") == "s" ? strings[int.Parse(c.Element(S + "v").Value)] : c.Element(S + "v").Value);
        }

        [Fact]
        public void Summary_IsTheFirstSheet_WithCaptionsPagesAndLinks()
        {
            var options = new XlsxExportOptions { IncludeSummary = true, Title = "Rapport annuel" };
            using (var zip = Write(options, Table(1, "Tableau 1 - Ventes", "Tableau 1 : Ventes par région", 3, 4), Table(4, "Tableau_4", null, 5, 5)))
            {
                var names = Part(zip, "xl/workbook.xml").Descendants(S + "sheet").Select(e => (string)e.Attribute("name")).ToList();
                Assert.Equal(new[] { "Sommaire", "Tableau 1 - Ventes", "Tableau_4" }, names);
                Assert.Equal(3, Part(zip, "[Content_Types].xml").Root.Elements().Count(e => ((string)e.Attribute("PartName") ?? "").StartsWith("/xl/worksheets/")));

                var cells = Cells(zip, "xl/worksheets/sheet1.xml");
                Assert.Equal("Tableaux exportés de « Rapport annuel »", cells["A1"]);
                Assert.Equal("N°", cells["A3"]);
                Assert.Equal("1", cells["A4"]);
                Assert.Equal("Tableau 1 : Ventes par région", cells["B4"]);
                Assert.Equal("3-4", cells["C4"]);
                Assert.Equal("Tableau 1 - Ventes", cells["D4"]);
                Assert.Equal("4", cells["A5"]);
                Assert.Equal("(sans légende)", cells["B5"]);
                Assert.Equal("5", cells["C5"]);

                var summary = Part(zip, "xl/worksheets/sheet1.xml");
                var links = summary.Descendants(S + "hyperlink").ToList();
                Assert.Equal(new[] { "D4", "D5" }, links.Select(l => (string)l.Attribute("ref")));
                Assert.Equal("'Tableau 1 - Ventes'!A1", (string)links[0].Attribute("location"));
                Assert.Equal("'Tableau_4'!A1", (string)links[1].Attribute("location"));

                // Le sommaire est l'onglet actif (un seul onglet sélectionné), ligne d'en-tête figée.
                Assert.Equal("1", (string)summary.Descendants(S + "sheetView").Single().Attribute("tabSelected"));
                Assert.Equal("3", (string)summary.Descendants(S + "pane").Single().Attribute("ySplit"));
                Assert.Null(Part(zip, "xl/worksheets/sheet2.xml").Descendants(S + "sheetView").Single().Attribute("tabSelected"));
            }
        }

        [Fact]
        public void Summary_NameNeverCollides_AndIsRecognizedAtImport()
        {
            using (var zip = Write(new XlsxExportOptions { IncludeSummary = true }, Table(1, "Sommaire", "Sommaire", 1, 1), Table(2, "Tableau_2", null, 2, 2)))
            {
                var names = Part(zip, "xl/workbook.xml").Descendants(S + "sheet").Select(e => (string)e.Attribute("name")).ToList();
                Assert.Equal("Sommaire (2)", names[0]);
            }
            Assert.True(XlsxWorkbookWriter.IsSummarySheetName("Sommaire"));
            Assert.True(XlsxWorkbookWriter.IsSummarySheetName("sommaire (3)"));
            Assert.False(XlsxWorkbookWriter.IsSummarySheetName("Sommaire des ventes"));
            Assert.False(XlsxWorkbookWriter.IsSummarySheetName("Tableau_1"));
        }

        [Fact]
        public void WithoutSummary_FirstTableSheetIsActive()
        {
            using (var zip = Write(new XlsxExportOptions(), Table(1, "Tableau_1", null, 1, 1), Table(2, "Tableau_2", null, 2, 2)))
            {
                Assert.Equal(2, Part(zip, "xl/workbook.xml").Descendants(S + "sheet").Count());
                Assert.Equal("1", (string)Part(zip, "xl/worksheets/sheet1.xml").Descendants(S + "sheetView").Single().Attribute("tabSelected"));
            }
        }

        [Fact]
        public void PageText()
        {
            Assert.Equal("", XlsxWorkbookWriter.PageText(0, 0));
            Assert.Equal("7", XlsxWorkbookWriter.PageText(7, 7));
            Assert.Equal("7", XlsxWorkbookWriter.PageText(7, 0));
            Assert.Equal("7-9", XlsxWorkbookWriter.PageText(7, 9));
        }

        [Fact]
        public void RepeatedHeaderRows_AreFrozen_BelowTheCaption()
        {
            var withCaption = Table(1, "Tableau_1", "Tableau 1 : Ventes", 1, 1);
            withCaption.HeaderRowCount = 1;
            var noCaption = Table(2, "Tableau_2", null, 1, 1);
            noCaption.HeaderRowCount = 2;
            var allHeader = Table(3, "Tableau_3", null, 1, 1);
            allHeader.HeaderRowCount = allHeader.RowCount; // tout le tableau répété : rien à figer
            var none = Table(4, "Tableau_4", null, 1, 1);

            using (var zip = Write(new XlsxExportOptions { IncludeCaptionRow = true }, withCaption, noCaption, allHeader, none))
            {
                var pane = Part(zip, "xl/worksheets/sheet1.xml").Descendants(S + "pane").Single();
                Assert.Equal("3", (string)pane.Attribute("ySplit")); // légende (A1), ligne vide, 1 ligne d'en-tête
                Assert.Equal("A4", (string)pane.Attribute("topLeftCell"));
                Assert.Equal("frozen", (string)pane.Attribute("state"));
                Assert.Equal("2", (string)Part(zip, "xl/worksheets/sheet2.xml").Descendants(S + "pane").Single().Attribute("ySplit"));
                Assert.Empty(Part(zip, "xl/worksheets/sheet3.xml").Descendants(S + "pane"));
                Assert.Empty(Part(zip, "xl/worksheets/sheet4.xml").Descendants(S + "pane"));
            }
        }

        [Fact]
        public void Parser_ReadsRepeatedHeaderRows()
        {
            const string xml = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:tbl>"
                + @"<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>En-tête 1</w:t></w:r></w:p></w:tc></w:tr>"
                + @"<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>En-tête 2</w:t></w:r></w:p></w:tc></w:tr>"
                + @"<w:tr><w:tc><w:p><w:r><w:t>Donnée</w:t></w:r></w:p></w:tc></w:tr>"
                + @"<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Pas en tête</w:t></w:r></w:p></w:tc></w:tr>"
                + @"</w:tbl></w:body></w:document>";
            Assert.Equal(2, WordXmlTableParser.Parse(xml).HeaderRowCount);
        }
    }
}
