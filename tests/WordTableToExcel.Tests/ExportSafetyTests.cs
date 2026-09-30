using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Text;
using WordTableToExcel.Tests.Fakes;
using WordTableToExcel.Word;
using Xunit;
using static WordTableToExcel.Tests.TestUtil;

namespace WordTableToExcel.Tests
{
    /// <summary>Informations absentes du classeur toujours signalées ; XML de Word chargé sans les parties inutiles.</summary>
    public class ExportSafetyTests
    {
        private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        [Fact]
        public void Clean_ReportsTruncationAtExcelLimit()
        {
            bool truncated;
            var runs = CellTextSanitizer.Clean(new[] { new TextRun(new string('a', 30000), null, null), new TextRun(new string('b', 5000), new RunFormat { Bold = true }, null) }, out truncated);
            Assert.True(truncated);
            Assert.Equal(CellTextSanitizer.ExcelMaxCellLength, runs.Sum(r => r.Text.Length));

            CellTextSanitizer.Clean(new[] { new TextRun(new string('a', CellTextSanitizer.ExcelMaxCellLength), null, null) }, out truncated);
            Assert.False(truncated);
        }

        [Fact]
        public void CountObjects_CountsEachVisibleObjectOnce()
        {
            var tbl = XElement.Parse(@"<w:tbl xmlns:w=""" + W + @""" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:v=""urn:schemas-microsoft-com:vml"">
  <w:tr><w:tc><w:p>
    <w:r><w:drawing/></w:r>
    <mc:AlternateContent><mc:Choice Requires=""wps""><w:r><w:drawing/></w:r></mc:Choice><mc:Fallback><w:r><w:pict/></w:r></mc:Fallback></mc:AlternateContent>
    <w:r><w:drawing><w:txbxContent><w:p><w:r><w:drawing/></w:r></w:p></w:txbxContent></w:drawing></w:r>
    <w:r><w:object><v:shape/></w:object></w:r>
    <w:del w:id=""1"" w:author=""A""><w:r><w:drawing/></w:r></w:del>
  </w:p></w:tc></w:tr>
  <w:tr><w:trPr><w:del w:id=""2"" w:author=""A""/></w:trPr><w:tc><w:p><w:r><w:pict/></w:r></w:p></w:tc></w:tr>
</w:tbl>");
            // Image, zone de texte (avec son image intérieure), forme (et sa représentation de secours), objet OLE ;
            // pas l'image supprimée ni celle de la ligne supprimée.
            Assert.Equal(4, WordXmlTableParser.CountObjects(tbl));
        }

        /// <summary>Parties que Word joint au XML de chaque tableau et que l'analyse ne lit pas.</summary>
        private const string UselessParts =
            @"<pkg:part pkg:name=""/word/media/image1.png"" pkg:contentType=""image/png"" pkg:compression=""store""><pkg:binaryData>iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk</pkg:binaryData></pkg:part>"
            + @"<pkg:part pkg:name=""/word/fontTable.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml""><pkg:xmlData><w:fonts xmlns:w=""" + W + @"""><w:font w:name=""Calibri""/></w:fonts></pkg:xmlData></pkg:part>"
            + @"<pkg:part pkg:name=""/word/settings.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml""><pkg:xmlData><w:settings xmlns:w=""" + W + @"""><w:zoom w:percent=""100""/></w:settings></pkg:xmlData></pkg:part>";

        private static string WithUselessParts(string flat)
        {
            int start = flat.IndexOf("<pkg:part", System.StringComparison.Ordinal);
            int end = flat.LastIndexOf("</pkg:package>", System.StringComparison.Ordinal);
            // Au début et à la fin du paquet : parties sautées avant et après les parties utiles.
            return flat.Substring(0, start) + UselessParts + flat.Substring(start, end - start) + UselessParts + flat.Substring(end);
        }

        private static string Signature(TableLayout layout)
        {
            var sb = new StringBuilder();
            sb.Append(layout.RowCount).Append('x').Append(layout.ColumnCount).Append('|').Append(layout.XmlText).Append('\n');
            foreach (var c in layout.Cells)
            {
                sb.Append(c).Append(' ').Append(c.Fill).Append(' ').Append(c.Borders == null ? "" : c.Borders.Top + "/" + c.Borders.Bottom);
                if (c.Content != null)
                {
                    foreach (var run in c.Content.Runs)
                    {
                        sb.Append(" «").Append(run.Text).Append("» ").Append(run.Format.FontName).Append(run.Format.Bold).Append(run.Format.Color).Append(run.Format.Size);
                    }
                }
                sb.Append('\n');
            }
            return sb.ToString();
        }

        [Fact]
        public void Load_SkipsUnusedParts_WithTheSameResult()
        {
            string flat = Fixture("medium_shading_merged.flat.xml");
            string heavy = WithUselessParts(flat);

            bool skipped;
            var light = WordXmlTableParser.Load(heavy, true, out skipped);
            Assert.True(skipped);
            Assert.DoesNotContain(light.Descendants(), e => e.Name.LocalName == "binaryData" || e.Name.LocalName == "fonts" || e.Name.LocalName == "settings");
            Assert.Equal(2, light.Root.Elements().Count()); // document et styles

            var expected = WordXmlTableParser.Parse(WordXmlTableParser.Load(flat, false, out skipped));
            Assert.True(expected.HasContent);
            Assert.Equal(Signature(expected), Signature(WordXmlTableParser.Parse(heavy)));
        }

        [Fact]
        public void Load_UnusualPackage_IsReadEntirely()
        {
            // Partie principale au nom et au type inattendus : écartée d'abord, puis relecture complète.
            string flat = Fixture("medium_shading_merged.flat.xml")
                .Replace(@"pkg:name=""/word/document.xml""", @"pkg:name=""/word/autre.xml""")
                .Replace("wordprocessingml.document.main+xml", "wordprocessingml.autre+xml");

            bool skipped;
            var light = WordXmlTableParser.Load(flat, true, out skipped);
            Assert.True(skipped);
            Assert.DoesNotContain(light.Descendants(), e => e.Name.LocalName == "tbl");

            var layout = WordXmlTableParser.Parse(flat);
            Assert.NotNull(layout);
            Assert.Equal(5, layout.RowCount);
        }

        [Fact]
        public void Load_PlainDocument_IsReadAsIs()
        {
            const string xml = @"<w:document xmlns:w=""" + W + @"""><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>";
            bool skipped;
            var document = WordXmlTableParser.Load(xml, true, out skipped);
            Assert.False(skipped);
            Assert.Equal("document", document.Root.Name.LocalName);
            Assert.Equal(1, WordXmlTableParser.Parse(xml).RowCount);
        }

        private static readonly CharFormat Plain = new CharFormat { Name = "Arial", Size = 10 };

        [Fact]
        public void Reader_ReportsObjectsNotExported()
        {
            const string xml = @"<w:document xmlns:w=""" + W + @"""><w:body><w:tbl>"
                + @"<w:tblGrid><w:gridCol w:w=""2160""/></w:tblGrid><w:tr>"
                + @"<w:tc><w:tcPr><w:tcW w:w=""2160"" w:type=""dxa""/></w:tcPr><w:p><w:r><w:t>Logo</w:t></w:r><w:r><w:drawing/></w:r></w:p></w:tc>"
                + @"</w:tr></w:tbl></w:body></w:document>";
            var b = new FakeDocumentBuilder();
            var table = b.Table(new List<IList<FakeCellSpec>> { new List<FakeCellSpec> { FakeCellSpec.Of(1, "Logo", Plain, 108) } }, xml);

            var model = new WordTableReader(b.Doc, null) { ReadContentFromXml = true }.Read(table, 1, null);

            Assert.Equal("Logo", At(model, 0, 0).PlainText);
            Assert.Contains(model.Omissions, o => o.StartsWith("1 image ou objet"));
        }

        [Fact]
        public void Reader_ReportsTruncatedCells()
        {
            var b = new FakeDocumentBuilder();
            var rows = new List<IList<FakeCellSpec>>
            {
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Court", Plain, 108), FakeCellSpec.Of(2, new string('x', 40000), Plain, 108) }
            };
            var table = b.Table(rows);

            var reader = new WordTableReader(b.Doc, null);
            var model = reader.Read(table, 1, null);

            Assert.True(reader.LastReadCellByCell);
            Assert.Equal(CellTextSanitizer.ExcelMaxCellLength, At(model, 0, 1).PlainText.Length);
            Assert.Single(model.Omissions);
            Assert.Contains("L1C2", model.Omissions[0]);
            Assert.Contains("coupé", model.Omissions[0]);
        }
    }
}
