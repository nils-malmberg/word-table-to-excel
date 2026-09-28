using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Xlsx;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class XlsxWriterTests
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>Tableau d'exemple : fusions, texte enrichi, fonds, bordures, alignements, légende.</summary>
        public static TableModel SampleTable(string sheetName = "Tableau 1 - Ventes", string caption = "Tableau 1 : Ventes par région")
        {
            var table = new TableModel
            {
                DocumentIndex = 1,
                Caption = caption,
                CaptionPosition = CaptionPosition.Above,
                SheetName = sheetName,
                RowCount = 3,
                ColumnCount = 3,
                ColumnWidthsPt = new[] { 100.0, 80, 80 },
                RowHeightsPt = new double[3],
                RowHeightExact = new bool[3]
            };
            var thin = new BorderLine(BorderStyle.Thin, Rgb.FromHex("7BA0CD"));
            var header = new RunFormat { FontName = "Arial", Size = 10, Bold = true, Color = Rgb.White };

            var title = new CellModel { Row = 0, Column = 0, ColumnSpan = 2, Fill = Rgb.FromHex("4F81BD"), HorizontalAlignment = HorizontalAlignment.Center };
            title.Runs.Add(new TextRun("Région & ventes <2023>", header, null));
            title.Borders = new CellBorders { Top = thin, Bottom = thin, Left = thin, Right = thin };
            table.Cells.Add(title);

            var evolution = new CellModel { Row = 0, Column = 2, Fill = Rgb.FromHex("4F81BD") };
            evolution.Runs.Add(new TextRun("Évolution", header, null));
            table.Cells.Add(evolution);

            var nord = new CellModel { Row = 1, Column = 0, RowSpan = 2, VerticalAlignment = VerticalAlignment.Center };
            nord.Runs.Add(new TextRun("Nord ", new RunFormat { FontName = "Arial", Size = 10 }, null));
            nord.Runs.Add(new TextRun("(dont Lille)", new RunFormat { FontName = "Arial", Size = 8, Italic = true, Color = Rgb.FromHex("FF0000") }, null));
            nord.Runs.Add(new TextRun("\nm²", new RunFormat { FontName = "Arial", Size = 10, Position = VerticalPosition.Superscript, Underline = UnderlineKind.Double, Strike = true }, null));
            table.Cells.Add(nord);

            var value = new CellModel { Row = 1, Column = 1, HorizontalAlignment = HorizontalAlignment.Right, Fill = Rgb.FromHex("FFFF00") };
            value.Runs.Add(new TextRun("1 234,50", new RunFormat { FontName = "Arial", Size = 10 }, null));
            value.Borders = new CellBorders { Bottom = new BorderLine(BorderStyle.Thick, Rgb.FromHex("FF0000")) };
            table.Cells.Add(value);

            var percent = new CellModel { Row = 1, Column = 2, TextRotation = 90 };
            percent.Runs.Add(new TextRun("12,5 %", new RunFormat { FontName = "Arial", Size = 10 }, null));
            table.Cells.Add(percent);

            table.Cells.Add(new CellModel { Row = 2, Column = 1 });
            var last = new CellModel { Row = 2, Column = 2 };
            last.Runs.Add(new TextRun("007", new RunFormat(), null));
            table.Cells.Add(last);
            return table;
        }

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

        [Fact]
        public void PackageContainsAllParts()
        {
            using (var zip = Write(new XlsxExportOptions(), SampleTable(), SampleTable("Tableau_2", null)))
            {
                foreach (var name in new[] { "[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml", "xl/workbook.xml",
                    "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml" })
                {
                    Assert.NotNull(Part(zip, name));
                }
                var sheets = Part(zip, "xl/workbook.xml").Descendants(S + "sheet").Select(e => (string)e.Attribute("name")).ToList();
                Assert.Equal(new[] { "Tableau 1 - Ventes", "Tableau_2" }, sheets);
            }
        }

        [Fact]
        public void CaptionRowMergesAndBorders()
        {
            using (var zip = Write(new XlsxExportOptions { IncludeCaptionRow = true }, SampleTable()))
            {
                var sheet = Part(zip, "xl/worksheets/sheet1.xml");
                var merges = sheet.Descendants(S + "mergeCell").Select(e => (string)e.Attribute("ref")).ToList();
                Assert.Equal(new[] { "A3:B3", "A4:A5" }, merges);

                var strings = Part(zip, "xl/sharedStrings.xml").Descendants(S + "si").ToList();
                var a1 = sheet.Descendants(S + "c").Single(c => (string)c.Attribute("r") == "A1");
                Assert.Equal("s", (string)a1.Attribute("t"));
                Assert.Equal("Tableau 1 : Ventes par région", strings[int.Parse(a1.Element(S + "v").Value)].Value);

                // Les cases couvertes par une fusion existent (pour porter fond et bordures).
                Assert.NotNull(sheet.Descendants(S + "c").SingleOrDefault(c => (string)c.Attribute("r") == "B3"));
                Assert.NotNull(sheet.Descendants(S + "c").SingleOrDefault(c => (string)c.Attribute("r") == "A5"));

                var cols = sheet.Descendants(S + "col").ToList();
                Assert.Equal(3, cols.Count);
            }
        }

        [Fact]
        public void WithoutCaptionRow_TableStartsAtA1()
        {
            using (var zip = Write(new XlsxExportOptions { IncludeCaptionRow = false }, SampleTable()))
            {
                var sheet = Part(zip, "xl/worksheets/sheet1.xml");
                Assert.Contains("A1:B1", sheet.Descendants(S + "mergeCell").Select(e => (string)e.Attribute("ref")));
            }
        }

        [Fact]
        public void RichTextRunsKeepFormatting()
        {
            using (var zip = Write(new XlsxExportOptions(), SampleTable()))
            {
                var rich = Part(zip, "xl/sharedStrings.xml").Descendants(S + "si").Single(si => si.Elements(S + "r").Any());
                var runs = rich.Elements(S + "r").ToList();
                Assert.Equal(3, runs.Count);
                Assert.Equal("Nord ", runs[0].Element(S + "t").Value);
                var italic = runs[1].Element(S + "rPr");
                Assert.NotNull(italic.Element(S + "i"));
                Assert.Equal("8", (string)italic.Element(S + "sz").Attribute("val"));
                Assert.Equal("FFFF0000", (string)italic.Element(S + "color").Attribute("rgb"));
                Assert.Equal("Arial", (string)italic.Element(S + "rFont").Attribute("val"));
                var sup = runs[2].Element(S + "rPr");
                Assert.Equal("superscript", (string)sup.Element(S + "vertAlign").Attribute("val"));
                Assert.Equal("double", (string)sup.Element(S + "u").Attribute("val"));
                Assert.NotNull(sup.Element(S + "strike"));
                Assert.Equal("\nm²", runs[2].Element(S + "t").Value);
            }
        }

        [Fact]
        public void StylesContainFillsFontsBordersAlignment()
        {
            using (var zip = Write(new XlsxExportOptions(), SampleTable()))
            {
                var styles = Part(zip, "xl/styles.xml");
                var fills = styles.Descendants(S + "fgColor").Select(e => (string)e.Attribute("rgb")).ToList();
                Assert.Contains("FF4F81BD", fills);
                Assert.Contains("FFFFFF00", fills);
                Assert.Contains(styles.Descendants(S + "font"), f => f.Element(S + "b") != null && (string)f.Element(S + "name").Attribute("val") == "Arial");
                Assert.Contains(styles.Descendants(S + "bottom"), b => (string)b.Attribute("style") == "thick");
                Assert.Contains(styles.Descendants(S + "alignment"), a => (string)a.Attribute("textRotation") == "90");
                Assert.Contains(styles.Descendants(S + "alignment"), a => (string)a.Attribute("horizontal") == "center" && (string)a.Attribute("wrapText") == "1");

                // Les compteurs déclarés correspondent au contenu.
                foreach (var name in new[] { "fonts", "fills", "borders", "cellXfs" })
                {
                    var e = styles.Root.Element(S + name);
                    Assert.Equal(e.Elements().Count(), (int)e.Attribute("count"));
                }
            }
        }

        [Fact]
        public void NumbersAreConvertedOnlyWhenRequested()
        {
            var options = new XlsxExportOptions { ConvertNumbers = true, NumberCulture = new CultureInfo("fr-FR") };
            using (var zip = Write(options, SampleTable()))
            {
                var sheet = Part(zip, "xl/worksheets/sheet1.xml");
                var b4 = sheet.Descendants(S + "c").Single(c => (string)c.Attribute("r") == "B4");
                Assert.Null(b4.Attribute("t"));
                Assert.Equal("1234.5", b4.Element(S + "v").Value);
                var c4 = sheet.Descendants(S + "c").Single(c => (string)c.Attribute("r") == "C4");
                Assert.Equal("0.125", c4.Element(S + "v").Value);
                var c5 = sheet.Descendants(S + "c").Single(c => (string)c.Attribute("r") == "C5");
                Assert.Equal("s", (string)c5.Attribute("t")); // « 007 » reste du texte

                var formats = Part(zip, "xl/styles.xml").Descendants(S + "numFmt").Select(e => (string)e.Attribute("formatCode")).ToList();
                Assert.Contains("#,##0.00", formats);
                Assert.Contains("0.0%", formats);
            }

            using (var zip = Write(new XlsxExportOptions(), SampleTable()))
            {
                var b4 = Part(zip, "xl/worksheets/sheet1.xml").Descendants(S + "c").Single(c => (string)c.Attribute("r") == "B4");
                Assert.Equal("s", (string)b4.Attribute("t"));
            }
        }

        [Fact]
        public void RowHeights()
        {
            var table = SampleTable();
            table.RowHeightsPt[2] = 40;
            table.RowHeightExact[2] = true;
            var longText = new CellModel { Row = 2, Column = 0, ColumnSpan = 1 };
            using (var zip = Write(new XlsxExportOptions { IncludeCaptionRow = false }, table))
            {
                var rows = Part(zip, "xl/worksheets/sheet1.xml").Descendants(S + "row").ToList();
                var row3 = rows.Single(r => (string)r.Attribute("r") == "3");
                Assert.Equal("40", (string)row3.Attribute("ht"));
                Assert.Equal("1", (string)row3.Attribute("customHeight"));
            }
        }

        [Fact]
        public void EstimatedHeightGrowsWithText()
        {
            var table = SampleTable();
            var shortCell = new CellModel { Row = 0, Column = 0 };
            shortCell.Runs.Add(new TextRun("abc", new RunFormat { Size = 10 }, null));
            var longCell = new CellModel { Row = 0, Column = 0 };
            longCell.Runs.Add(new TextRun(new string('x', 300), new RunFormat { Size = 10 }, null));
            Assert.True(XlsxWorkbookWriter.EstimateHeight(table, longCell) > 3 * XlsxWorkbookWriter.EstimateHeight(table, shortCell));
        }

        [Fact]
        public void DuplicateSheetNamesAreRejected()
        {
            var writer = new XlsxWorkbookWriter(new XlsxExportOptions());
            writer.AddTable(SampleTable("A"));
            Assert.Throws<ArgumentException>(() => writer.AddTable(SampleTable("a")));
            Assert.Throws<InvalidOperationException>(() => new XlsxWorkbookWriter(null).Save(new MemoryStream()));
        }

        [Fact]
        public void InvalidXmlCharactersNeverReachTheFile()
        {
            var table = SampleTable();
            table.Cells[0].Runs[0].Text = "a\u0001b\uFFFEc\uD800";
            using (var zip = Write(new XlsxExportOptions(), table))
            {
                Assert.Contains(Part(zip, "xl/sharedStrings.xml").Descendants(S + "t"), t => t.Value == "abc");
            }
        }

        [Fact]
        public void ColumnNames()
        {
            Assert.Equal("A", ExcelUnits.ColumnName(0));
            Assert.Equal("Z", ExcelUnits.ColumnName(25));
            Assert.Equal("AA", ExcelUnits.ColumnName(26));
            Assert.Equal("XFD", ExcelUnits.ColumnName(16383));
            Assert.Equal("C7", ExcelUnits.CellReference(6, 2));
        }

        /// <summary>
        /// Écrit un classeur d'exemple pour une vérification externe (openpyxl, LibreOffice, Excel) si
        /// la variable d'environnement WTTE_SAMPLE_XLSX indique un chemin.
        /// </summary>
        [Fact]
        public void WriteSampleWorkbookForExternalValidation()
        {
            string path = Environment.GetEnvironmentVariable("WTTE_SAMPLE_XLSX");
            if (string.IsNullOrEmpty(path)) return;
            var writer = new XlsxWorkbookWriter(new XlsxExportOptions { ConvertNumbers = true, NumberCulture = new CultureInfo("fr-FR"), Title = "Exemple" });
            writer.AddTable(SampleTable());
            writer.AddTable(SampleTable("Tableau_2", null));
            writer.Save(path);
        }
    }

    public class ZipWriterTests
    {
        [Fact]
        public void RoundTrip()
        {
            var ms = new MemoryStream();
            var big = Encoding.UTF8.GetBytes(string.Concat(Enumerable.Repeat("<row>données répétées</row>", 2000)));
            var random = new byte[3000];
            new Random(42).NextBytes(random);
            using (var zip = new ZipWriter(ms))
            {
                zip.AddEntry("a/text.xml", "<é>ü</é>");
                zip.AddEntry("b/big.xml", big);
                zip.AddEntry("c/random.bin", random); // incompressible : stocké tel quel
                zip.AddEntry("d/empty.txt", new byte[0]);
            }

            ms.Position = 0;
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                Assert.Equal(4, archive.Entries.Count);
                Assert.Equal("<é>ü</é>", new StreamReader(archive.GetEntry("a/text.xml").Open()).ReadToEnd());
                var copy = new MemoryStream();
                archive.GetEntry("b/big.xml").Open().CopyTo(copy);
                Assert.Equal(big, copy.ToArray());
                Assert.True(archive.GetEntry("b/big.xml").CompressedLength < big.Length / 10);
                copy = new MemoryStream();
                archive.GetEntry("c/random.bin").Open().CopyTo(copy);
                Assert.Equal(random, copy.ToArray());
                Assert.Equal(0, archive.GetEntry("d/empty.txt").Length);
            }
        }
    }
}
