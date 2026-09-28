using System.Collections.Generic;
using System.Linq;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Export;
using WordTableToExcel.Tests.Fakes;
using WordTableToExcel.Word;
using Xunit;
using static WordTableToExcel.Tests.TestUtil;

namespace WordTableToExcel.Tests
{
    public class WordRunReaderTests
    {
        private static CharFormat F(bool bold = false, bool italic = false, int color = -16777216, float size = 11, string name = "Calibri")
        {
            return new CharFormat { Bold = bold, Italic = italic, Color = color, Size = size, Name = name };
        }

        [Fact]
        public void UniformText_IsReadWithFewCalls()
        {
            var b = new FakeDocumentBuilder();
            b.Text("Bonjour tout le monde", F(bold: true, size: 12, name: "Arial"));
            var reader = new WordRunReader(b.Doc, null);

            var runs = reader.Read(0, b.Doc.Chars.Count);

            var run = Assert.Single(runs);
            Assert.Equal("Bonjour tout le monde", run.Text);
            Assert.True(run.Format.Bold);
            Assert.Equal("Arial", run.Format.FontName);
            Assert.Equal(12, run.Format.Size);
            Assert.Null(run.Format.Color);
            Assert.True(b.Doc.RangeCalls <= 2);
        }

        [Fact]
        public void MixedFormatting_IsSplitIntoRuns()
        {
            var b = new FakeDocumentBuilder();
            b.Text("Chiffre ", F());
            b.Text("d'affaires", F(bold: true, italic: true, color: 0x0000FF)); // BGR : rouge
            b.Text(" 2023", F());
            var reader = new WordRunReader(b.Doc, null);

            var runs = reader.Read(0, b.Doc.Chars.Count);

            Assert.Equal(new[] { "Chiffre ", "d'affaires", " 2023" }, runs.Select(r => r.Text));
            Assert.True(runs[1].Format.Bold && runs[1].Format.Italic);
            Assert.Equal(new Rgb(255, 0, 0), runs[1].Format.Color);
            Assert.False(runs[2].Format.Bold);
        }

        [Fact]
        public void LongText_CostIsLogarithmic()
        {
            var b = new FakeDocumentBuilder();
            b.Text(new string('a', 3000), F());
            b.Text("important", F(bold: true));
            b.Text(new string('b', 3000), F());
            var reader = new WordRunReader(b.Doc, null);

            var runs = reader.Read(0, b.Doc.Chars.Count);

            Assert.Equal(3, runs.Count);
            Assert.Equal("important", runs[1].Text);
            Assert.True(b.Doc.RangeCalls < 80, "appels : " + b.Doc.RangeCalls);
        }

        [Fact]
        public void HiddenTextIsSkipped_AllCapsIsApplied()
        {
            var b = new FakeDocumentBuilder();
            b.Text("visible ", F());
            b.Text("secret", new CharFormat { Hidden = true });
            b.Text("majuscules", new CharFormat { AllCaps = true });
            var reader = new WordRunReader(b.Doc, null);

            var text = string.Concat(reader.Read(0, b.Doc.Chars.Count).Select(r => r.Text));

            Assert.Equal("visible MAJUSCULES", text);
        }

        [Fact]
        public void HiddenFieldCode_FallsBackToParagraphs()
        {
            var b = new FakeDocumentBuilder();
            b.Text("Voir page ", F(bold: true));
            b.HiddenCode(" PAGEREF _Toc1 \\h ");
            b.Text("12", F(bold: true));
            b.Text("\r", F());
            b.Text("Suite", F(italic: true));
            var reader = new WordRunReader(b.Doc, null);

            var runs = reader.Read(0, b.Doc.Chars.Count);

            Assert.Equal("Voir page 12\rSuite", string.Concat(runs.Select(r => r.Text)));
            Assert.True(runs[0].Format.Bold);
            Assert.True(runs.Last().Format.Italic);
        }

        [Fact]
        public void ThemeColor_IsResolvedThroughTextColor()
        {
            var b = new FakeDocumentBuilder();
            int accent1 = unchecked((int)0xD400FFFF);
            b.Text("thème", new CharFormat { Color = accent1, ResolvedThemeRgb = 0xBD814F });
            var reader = new WordRunReader(b.Doc, null);

            Assert.Equal(Hex("4F81BD"), reader.Read(0, b.Doc.Chars.Count).Single().Format.Color);
        }

        [Fact]
        public void ThemeColor_FallsBackToThemeLookup()
        {
            var b = new FakeDocumentBuilder();
            b.Text("thème", new CharFormat { Color = unchecked((int)0xD400FFFF) });
            var reader = new WordRunReader(b.Doc, index => index == 5 ? Hex("4F81BD") : (Rgb?)null);

            Assert.Equal(Hex("4F81BD"), reader.Read(0, b.Doc.Chars.Count).Single().Format.Color);
        }

        [Fact]
        public void SurrogatePairsAreNotSplit_HighlightIsRead()
        {
            var b = new FakeDocumentBuilder();
            b.Text("a", F());
            b.Text("😀", F(bold: true));
            b.Text("b", new CharFormat { Highlight = 7 });
            var reader = new WordRunReader(b.Doc, null);

            var runs = reader.Read(0, b.Doc.Chars.Count);

            Assert.Equal(new[] { "a", "😀", "b" }, runs.Select(r => r.Text));
            Assert.Equal(Hex("FFFF00"), runs[2].Highlight);
        }

        [Fact]
        public void UnderlineStrikeAndPosition()
        {
            var b = new FakeDocumentBuilder();
            b.Text("x", new CharFormat { Underline = 3 });
            b.Text("y", new CharFormat { Underline = 1, Strike = true });
            b.Text("2", new CharFormat { Superscript = true });
            b.Text("3", new CharFormat { Subscript = true, DoubleStrike = true });
            var runs = new WordRunReader(b.Doc, null).Read(0, b.Doc.Chars.Count);

            Assert.Equal(UnderlineKind.Double, runs[0].Format.Underline);
            Assert.Equal(UnderlineKind.Single, runs[1].Format.Underline);
            Assert.True(runs[1].Format.Strike);
            Assert.Equal(VerticalPosition.Superscript, runs[2].Format.Position);
            Assert.Equal(VerticalPosition.Subscript, runs[3].Format.Position);
            Assert.True(runs[3].Format.Strike);
        }
    }

    public class WordTableReaderTests
    {
        private static readonly CharFormat Plain = new CharFormat { Name = "Arial", Size = 10 };

        private static FakeCellSpec Cell(int col, string text, float width = 108)
        {
            return FakeCellSpec.Of(col, text, Plain, width);
        }

        /// <summary>Même tableau que la ressource medium_shading_merged : 5 lignes, fusions B1:C1 et A2:A3.</summary>
        private static FakeTable BuildSampleTable(FakeDocumentBuilder b, string xml)
        {
            b.Paragraph("Tableau 1 : Ventes par région", Plain, " SEQ Tableau \\* ARABIC ");
            var rows = new List<IList<FakeCellSpec>>
            {
                new List<FakeCellSpec> { Cell(1, "Région"), Cell(2, "Ventes", 216), Cell(3, "Évolution") },
                new List<FakeCellSpec> { Cell(1, "Nord"), Cell(2, "1 200"), Cell(3, "1 350"), Cell(4, "12,5 %") },
                new List<FakeCellSpec> { Cell(2, "800"), Cell(3, "760"), Cell(4, "-5 %") },
                new List<FakeCellSpec> { Cell(1, "Sud"), Cell(2, "2 000"), Cell(3, "2 100"), Cell(4, "5 %") },
                new List<FakeCellSpec> { Cell(1, "Total"), Cell(2, "4 000"), Cell(3, "4 210"), Cell(4, "5,25 %") }
            };
            rows[1][1].And(" (+", Plain).And("hausse", new CharFormat { Name = "Arial", Size = 10, Bold = true, Highlight = 4 }).And(")", Plain);
            rows[1][0].Configure = c => c.VerticalAlignment = 1;
            var table = b.Table(rows, xml);
            b.Paragraph("Source : service commercial.", Plain);
            return table;
        }

        [Fact]
        public void ReadsStructureFromXml_AndContentFromCom()
        {
            var b = new FakeDocumentBuilder();
            var table = BuildSampleTable(b, Fixture("medium_shading_merged.flat.xml"));
            var model = new WordTableReader(b.Doc, null).Read(table, 1, null);

            Assert.Equal(5, model.RowCount);
            Assert.Equal(4, model.ColumnCount);
            Assert.Equal(18, model.Cells.Count);
            Assert.Empty(model.Warnings);

            var ventes = At(model, 0, 1);
            Assert.Equal("Ventes", ventes.PlainText);
            Assert.Equal(2, ventes.ColumnSpan);
            Assert.Equal(Hex("4F81BD"), ventes.Fill);

            var nord = At(model, 1, 0);
            Assert.Equal("Nord", nord.PlainText);
            Assert.Equal(2, nord.RowSpan);

            // La ligne qui suit la fusion verticale est correctement appariée.
            Assert.Equal("800", At(model, 2, 1).PlainText);
            Assert.Equal("-5 %", At(model, 2, 3).PlainText);
            Assert.Equal(Hex("FFFF00"), At(model, 3, 3).Fill);

            var rich = At(model, 1, 1);
            Assert.Equal("1 200 (+hausse)", rich.PlainText);
            Assert.True(rich.Runs.Single(r => r.Text == "hausse").Format.Bold);
            Assert.Equal("Arial", rich.Runs[0].Format.FontName);
        }

        [Fact]
        public void FallbackWithoutXml_UsesWidthsAndComFormatting()
        {
            var b = new FakeDocumentBuilder();
            var table = BuildSampleTable(b, null);
            table.CellList[0].Shading.BackgroundPatternColor = 0xBD814F;             // BGR de 4F81BD
            table.CellList[0].Borders.Map[-3] = new FakeBorder { LineStyle = 7, LineWidth = 6, Color = 0x0000FF };

            var model = new WordTableReader(b.Doc, null).Read(table, 1, null);

            Assert.Equal(5, model.RowCount);
            Assert.Equal(4, model.ColumnCount);
            Assert.Equal(2, At(model, 0, 1).ColumnSpan);
            Assert.Equal(2, At(model, 1, 0).RowSpan);
            Assert.Equal("800", At(model, 2, 1).PlainText);
            Assert.Equal(Hex("4F81BD"), At(model, 0, 0).Fill);
            Assert.Equal(new BorderLine(BorderStyle.Double, Hex("FF0000")), At(model, 0, 0).Borders.Bottom);
            Assert.Equal(VerticalAlignment.Center, At(model, 1, 0).VerticalAlignment);
            // Surlignage partiel sans fond de cellule : devient le fond (Excel ne surligne pas une partie de cellule).
            Assert.Equal(Hex("00FF00"), At(model, 1, 1).Fill);
        }

        [Fact]
        public void InconsistentXml_FallsBackToWidths()
        {
            var b = new FakeDocumentBuilder();
            const string oneCell = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>";
            var table = BuildSampleTable(b, oneCell);

            var model = new WordTableReader(b.Doc, null).Read(table, 1, null);

            Assert.Equal(5, model.RowCount);
            Assert.Contains(model.Warnings, w => w.Contains("largeurs"));
        }

        [Fact]
        public void ProgressIsReported()
        {
            var b = new FakeDocumentBuilder();
            var table = BuildSampleTable(b, Fixture("medium_shading_merged.flat.xml"));
            var calls = new List<int>();
            new WordTableReader(b.Doc, null).Read(table, 1, (done, total) => calls.Add(done));
            Assert.Equal(18, calls.Last());
        }

        [Fact]
        public void DocumentGuardRestoresSavedFlag()
        {
            var b = new FakeDocumentBuilder();
            using (new WordDocumentGuard(b.Doc))
            {
                b.Doc.Saved = false; // une lecture aurait marqué le document comme modifié
            }
            Assert.True(b.Doc.Saved);

            b.Doc.Saved = false;
            using (new WordDocumentGuard(b.Doc))
            {
            }
            Assert.False(b.Doc.Saved); // un document déjà modifié reste modifié
        }
    }

    public class WordCaptionScannerTests
    {
        private static readonly CharFormat Plain = new CharFormat();

        private static FakeTable SmallTable(FakeDocumentBuilder b, string text)
        {
            return b.Table(new List<IList<FakeCellSpec>> { new List<FakeCellSpec> { FakeCellSpec.Of(1, text) } });
        }

        [Fact]
        public void FindsCaptionsAboveAndBelow_AndIgnoresFigures()
        {
            var b = new FakeDocumentBuilder();
            b.Paragraph("Introduction", Plain);
            b.Paragraph("Tableau 1 : Ventes", Plain, " SEQ Tableau \\* ARABIC ");
            var t1 = SmallTable(b, "a");
            b.Paragraph("Source : INSEE", Plain);
            b.Paragraph("", Plain);
            b.Paragraph("Tableau 2 \u2013 Coûts", Plain);
            b.Paragraph("", Plain);
            var t2 = SmallTable(b, "b");
            b.Paragraph("Figure 1 : Carte", Plain, " SEQ Figure \\* ARABIC ");
            var t3 = SmallTable(b, "c");
            b.Paragraph("Texte courant.", Plain);

            var scanner = new WordCaptionScanner(b.Doc, new CaptionMatcher(), null);
            var c1 = scanner.Scan(t1, 1);
            var c2 = scanner.Scan(t2, 2);
            var c3 = scanner.Scan(t3, 3);

            Assert.Equal("Tableau 1 : Ventes", c1.Above.Text);
            Assert.True(c1.Above.FromSequenceField);
            Assert.Null(c1.Below);
            Assert.Equal("Tableau 2 \u2013 Coûts", c2.Above.Text);
            Assert.False(c2.Above.FromSequenceField);
            Assert.Null(c2.Below);          // « Figure 1 » est une légende de figure
            Assert.Null(c3.Above);
            Assert.Null(c3.Below);

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(new[] { c1, c2, c3 }, out convention);
            Assert.Equal(CaptionPosition.Above, convention);
            Assert.NotNull(result[0].Caption);
            Assert.NotNull(result[1].Caption);
            Assert.Null(result[2].Caption);
        }

        [Fact]
        public void CaptionsBelowTables()
        {
            var b = new FakeDocumentBuilder();
            var t1 = SmallTable(b, "a");
            b.Paragraph("Table 1: Results", Plain, " SEQ Table \\* ARABIC ");
            b.Paragraph("Some text.", Plain);
            var t2 = SmallTable(b, "b");
            b.Paragraph("Table 2: Costs", Plain, " SEQ Table \\* ARABIC ");

            var scanner = new WordCaptionScanner(b.Doc, new CaptionMatcher(), null);
            var contexts = new[] { scanner.Scan(t1, 1), scanner.Scan(t2, 2) };
            CaptionPosition convention;
            var result = CaptionAssigner.Assign(contexts, out convention);

            Assert.Equal(CaptionPosition.Below, convention);
            Assert.Equal("Table 1: Results", result[0].Caption.Text);
            Assert.Equal("Table 2: Costs", result[1].Caption.Text);
        }

        [Fact]
        public void AdjacentTablesDoNotStealCaptionsThroughEachOther()
        {
            var b = new FakeDocumentBuilder();
            var t1 = SmallTable(b, "a");
            b.Paragraph("", Plain);
            var t2 = SmallTable(b, "b");

            var scanner = new WordCaptionScanner(b.Doc, new CaptionMatcher(), null);
            Assert.Null(scanner.Scan(t1, 1).Below);
            Assert.Null(scanner.Scan(t2, 2).Above);
        }

        [Fact]
        public void CaptionStyleAllowsUnnumberedCaption()
        {
            var b = new FakeDocumentBuilder();
            b.Paragraph("Tableaux comparatifs", new CharFormat { Style = "Légende" });
            var t1 = SmallTable(b, "a");

            var scanner = new WordCaptionScanner(b.Doc, new CaptionMatcher(), null);
            Assert.Equal("Tableaux comparatifs", scanner.Scan(t1, 1).Above.Text);
        }
    }

    public class ExportPlanTests
    {
        private static List<TableEntry> Entries()
        {
            return new List<TableEntry>
            {
                new TableEntry { Index = 1, Caption = "Tableau 1 : Ventes", CaptionPosition = CaptionPosition.Above },
                new TableEntry { Index = 2 },
                new TableEntry { Index = 3, Caption = "Tableau 1 : Ventes", CaptionPosition = CaptionPosition.Below },
                new TableEntry { Index = 4 }
            };
        }

        [Fact]
        public void OptionA_OnlyCaptionedTables()
        {
            var plan = ExportPlan.Build(Entries(), false);
            Assert.Equal(new[] { "Tableau 1 - Ventes", "Tableau 1 - Ventes (2)" }, plan.Select(p => p.SheetName));
            Assert.Equal(new[] { 1, 3 }, plan.Select(p => p.Table.Index));
        }

        [Fact]
        public void OptionB_AllTables_DefaultNamesUseDocumentPosition()
        {
            var plan = ExportPlan.Build(Entries(), true);
            Assert.Equal(new[] { "Tableau 1 - Ventes", "Tableau_2", "Tableau 1 - Ventes (2)", "Tableau_4" }, plan.Select(p => p.SheetName));
        }
    }

    public class WordColorTests
    {
        [Fact]
        public void Decode()
        {
            Assert.Null(WordColor.Decode(WordColor.Automatic, null));
            Assert.Null(WordColor.Decode(9999999, null));
            Assert.Equal(new Rgb(255, 0, 0), WordColor.Decode(0x0000FF, null));
            Assert.Equal(new Rgb(0x12, 0x34, 0x56), WordColor.Decode(0x563412, null));
            Assert.Null(WordColor.Decode(unchecked((int)0xD400FFFF), null));
            Assert.Equal(Hex("4F81BD"), WordColor.Decode(unchecked((int)0xD400FFFF), i => i == 5 ? Hex("4F81BD") : (Rgb?)null));
            Assert.Equal(new Rgb(0, 0, 0), WordColor.Decode(unchecked((int)0xDD00FFFF), i => i == 1 ? new Rgb(0, 0, 0) : (Rgb?)null)); // Texte 1
        }

        [Fact]
        public void Highlights()
        {
            Assert.Null(WordColor.FromHighlightIndex(0));
            Assert.Equal(Hex("FFFF00"), WordColor.FromHighlightIndex(7));
            Assert.Equal(Hex("00FF00"), WordColor.FromHighlightIndex(4));
            Assert.Null(WordColor.FromHighlightIndex(9999999));
        }

        [Fact]
        public void AlignmentMapping()
        {
            Assert.Equal(HorizontalAlignment.Left, WordTableReader.MapAlignment(0));
            Assert.Equal(HorizontalAlignment.Center, WordTableReader.MapAlignment(1));
            Assert.Equal(HorizontalAlignment.Right, WordTableReader.MapAlignment(2));
            Assert.Equal(HorizontalAlignment.Justify, WordTableReader.MapAlignment(3));
            Assert.Equal(HorizontalAlignment.General, WordTableReader.MapAlignment(9999999));
        }
    }
}

namespace WordTableToExcel.Tests
{
    using System;
    using System.Globalization;
    using WordTableToExcel.Core.Xlsx;

    public class EndToEndTests
    {
        /// <summary>
        /// Document factice → lecture Word → classeur Excel. Si WTTE_E2E_XLSX est défini, le classeur
        /// est aussi écrit à cet emplacement pour un contrôle visuel (LibreOffice, Excel).
        /// </summary>
        [Fact]
        public void DocumentToWorkbook()
        {
            var b = new FakeDocumentBuilder();
            var plain = new CharFormat { Name = "Arial", Size = 10 };
            b.Paragraph("Tableau 1 : Ventes par région", plain, " SEQ Tableau \\* ARABIC ");
            var rows = new List<IList<FakeCellSpec>>
            {
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Région", plain, 108), FakeCellSpec.Of(2, "Ventes", plain, 216), FakeCellSpec.Of(3, "Évolution", plain, 108) },
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Nord", plain, 108), FakeCellSpec.Of(2, "1 200", plain, 108), FakeCellSpec.Of(3, "1 350", plain, 108), FakeCellSpec.Of(4, "12,5 %", plain, 108) },
                new List<FakeCellSpec> { FakeCellSpec.Of(2, "800", plain, 108), FakeCellSpec.Of(3, "760", plain, 108), FakeCellSpec.Of(4, "-5 %", plain, 108) },
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Sud", plain, 108), FakeCellSpec.Of(2, "2 000", plain, 108), FakeCellSpec.Of(3, "2 100", plain, 108), FakeCellSpec.Of(4, "5 %", plain, 108) },
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Total", new CharFormat { Name = "Arial", Size = 10, Bold = true }, 108), FakeCellSpec.Of(2, "4 000", plain, 108), FakeCellSpec.Of(3, "4 210", plain, 108), FakeCellSpec.Of(4, "5,25 %", plain, 108) }
            };
            foreach (var cell in rows[0]) cell.Runs[0] = Tuple.Create(cell.Runs[0].Item1, new CharFormat { Name = "Arial", Size = 10, Bold = true, Color = 0xFFFFFF });
            rows[3][0].And(" (", plain).And("dont Marseille", new CharFormat { Name = "Arial", Size = 8, Italic = true, Color = 0x0000C0 }).And(")", plain);
            var table = b.Table(rows, TestUtil.Fixture("medium_shading_merged.flat.xml"));
            b.Paragraph("Source : service commercial.", plain);
            var t2 = b.Table(new List<IList<FakeCellSpec>>
            {
                new List<FakeCellSpec> { FakeCellSpec.Of(1, "Sans légende", plain, 150), FakeCellSpec.Of(2, "x", plain, 60) }
            });

            var scanner = new WordCaptionScanner(b.Doc, new CaptionMatcher(), null);
            CaptionPosition convention;
            var captions = CaptionAssigner.Assign(new[] { scanner.Scan(table, 1), scanner.Scan(t2, 2) }, out convention);
            var entries = new List<TableEntry>
            {
                new TableEntry { Index = 1, Caption = captions[0].Caption?.Text, CaptionPosition = captions[0].Position },
                new TableEntry { Index = 2, Caption = captions[1].Caption?.Text, CaptionPosition = captions[1].Position }
            };

            var writer = new XlsxWorkbookWriter(new XlsxExportOptions { ConvertNumbers = true, NumberCulture = new CultureInfo("fr-FR") });
            var reader = new WordTableReader(b.Doc, null);
            foreach (var sheet in ExportPlan.Build(entries, true))
            {
                var model = reader.Read(b.Doc.Tables.Item(sheet.Table.Index), sheet.Table.Index, null);
                model.Caption = sheet.Table.Caption;
                model.SheetName = sheet.SheetName;
                writer.AddTable(model);
            }

            Assert.Equal(2, writer.SheetCount);
            Assert.Equal("Tableau 1 : Ventes par région", entries[0].Caption);
            Assert.Null(entries[1].Caption);

            var ms = new System.IO.MemoryStream();
            writer.Save(ms);
            Assert.True(ms.Length > 1000);

            string path = Environment.GetEnvironmentVariable("WTTE_E2E_XLSX");
            if (!string.IsNullOrEmpty(path)) System.IO.File.WriteAllBytes(path, ms.ToArray());
        }
    }
}
