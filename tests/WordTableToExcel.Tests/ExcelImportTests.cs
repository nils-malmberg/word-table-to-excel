using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Xlsx;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class ExcelImportTests
    {
        private static readonly CultureInfo FrFr = CultureInfo.GetCultureInfo("fr-FR");
        private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");

        private static string FixturePath(string name)
        {
            return Path.Combine(AppContext.BaseDirectory, "fixtures", name);
        }

        private static SheetImport Import(XlsxWorkbook workbook, string sheet, CultureInfo culture = null, ExcelImportOptions options = null, CellRange? range = null)
        {
            culture = culture ?? FrFr;
            var format = new ExcelFormatSettings(culture, culture, workbook.Date1904, workbook.Styles.Colors);
            var converter = new SheetConverter(workbook, format, options ?? new ExcelImportOptions(), null);
            var info = workbook.Sheets.Single(s => s.Name == sheet);
            return converter.Convert(workbook.ReadSheet(info), range);
        }

        private static string Text(TableModel table, int row, int column)
        {
            return TestUtil.At(table, row, column).PlainText;
        }

        private static string Nb(string text)
        {
            // Séparateurs du format français : espace fine insécable (ICU) ou insécable (Windows) pour les milliers.
            return text.Replace("\u202F", " ").Replace("\u00A0", " ");
        }

        // ------------------------------------------------------------------ lecture

        [Fact]
        public void ReadsWorkbookStructure()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            Assert.Equal(new[] { "Ventes 2024", "Planning", "Achats", "Paramètres", "Vide", "Très large" }, wb.Sheets.Select(s => s.Name).ToArray());
            Assert.Equal(XlsxSheetState.Hidden, wb.Sheets[3].State);
            Assert.All(wb.Sheets, s => Assert.Equal(XlsxSheetKind.Worksheet, s.Kind));
            Assert.False(wb.Date1904);

            var ventes = wb.ReadSheet(wb.Sheets[0]);
            var b4 = ventes.Cell(3, 1);
            Assert.Equal(XlsxCellType.Number, b4.Type);
            Assert.Equal(125430.5, b4.Number);
            var f4 = ventes.Cell(3, 5);
            Assert.True(f4.HasFormula);
            Assert.Equal(528281.5, f4.Number, 6);
            Assert.True(ventes.IsRowHidden(8));
            Assert.True(ventes.IsColumnHidden(7));
            Assert.Single(ventes.ConditionalRules);

            var planning = wb.ReadSheet(wb.Sheets[1]);
            Assert.Contains(new CellRange(1, 0, 3, 0), planning.MergedRanges);
            Assert.Contains(new CellRange(6, 1, 6, 7), planning.MergedRanges);
            Assert.Equal(XlsxCellType.Error, planning.Cell(4, 7).Type);
            Assert.True(planning.Cell(6, 1).Text.IsRich);
        }

        [Fact]
        public void RejectsNonXlsxFilesWithClearMessages()
        {
            var ex = Assert.Throws<ExcelImportException>(() => XlsxWorkbook.Load(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0 }));
            Assert.Contains("97-2003", ex.Message);
            Assert.True(ex.ConvertibleByExcel);

            ex = Assert.Throws<ExcelImportException>(() => XlsxWorkbook.Load(Encoding.UTF8.GetBytes("Nom;Valeur\r\nA;1")));
            Assert.Contains("pas un classeur", ex.Message);

            ex = Assert.Throws<ExcelImportException>(() => XlsxWorkbook.Load(new byte[0]));
            Assert.Contains("vide", ex.Message);

            // Archive ZIP qui n'est pas un classeur (document Word).
            var docx = WordTableXmlWriter.BuildDocx(XlsxWriterTests.SampleTable());
            ex = Assert.Throws<ExcelImportException>(() => XlsxWorkbook.Load(docx));
            Assert.Contains("document Word", ex.Message);

            // Archive tronquée.
            var xlsx = File.ReadAllBytes(FixturePath("import_complexe.xlsx"));
            ex = Assert.Throws<ExcelImportException>(() => XlsxWorkbook.Load(xlsx.Take(xlsx.Length / 2).ToArray()));
        }

        [Fact]
        public void ZipReaderDetectsCorruption()
        {
            byte[] data;
            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipWriter(ms))
                {
                    zip.AddEntry("a.txt", "Bonjour les tableaux ! " + new string('x', 5000));
                    zip.Finish();
                }
                data = ms.ToArray();
            }
            var reader = new ZipReader(data);
            Assert.True(reader.Contains("A.TXT"));
            Assert.StartsWith("Bonjour", Encoding.UTF8.GetString(reader.Read("a.txt")));
            Assert.Null(reader.Read("absent.xml"));

            // Octet altéré dans les données compressées : le contrôle CRC (ou la décompression) échoue.
            data[40] ^= 0x55;
            Assert.ThrowsAny<Exception>(() => new ZipReader(data).Read("a.txt"));
        }

        [Theory]
        [InlineData("A1", 0, 0, 0, 0)]
        [InlineData("$B$2:$D$10", 1, 1, 9, 3)]
        [InlineData("Feuil1!C3:A1", 0, 0, 2, 2)]
        [InlineData("'Mes données'!$A$1:$F$20", 0, 0, 19, 5)]
        [InlineData("XFD1048576", 1048575, 16383, 1048575, 16383)]
        public void ParsesRanges(string text, int r1, int c1, int r2, int c2)
        {
            CellRange range;
            Assert.True(CellRange.TryParse(text, out range));
            Assert.Equal(new CellRange(r1, c1, r2, c2), range);
        }

        [Fact]
        public void ParsesWholeRowsAndColumns()
        {
            CellRange range;
            Assert.True(CellRange.TryParse("Feuil1!$1:$2", out range));
            Assert.Equal(0, range.FirstRow);
            Assert.Equal(1, range.LastRow);
            Assert.Equal(CellReference.MaxColumns - 1, range.LastColumn);
            Assert.True(CellRange.TryParse("B:D", out range));
            Assert.Equal(1, range.FirstColumn);
            Assert.Equal(3, range.LastColumn);
            Assert.Equal(CellReference.MaxRows - 1, range.LastRow);
            Assert.False(CellRange.TryParse("A0", out range));
            Assert.False(CellRange.TryParse("XFE1", out range));
            Assert.False(CellRange.TryParse("A1:B2:C3", out range));
            Assert.False(CellRange.TryParse("", out range));
        }

        [Fact]
        public void DecodesExcelEscapes()
        {
            Assert.Equal("a\r\nb", XlsxPackage.DecodeEscapes("a_x000D__x000A_b"));
            Assert.Equal("_x0041_", XlsxPackage.DecodeEscapes("_x005F_x0041_"));
            Assert.Equal("_xZZZZ_", XlsxPackage.DecodeEscapes("_xZZZZ_"));
        }

        // ------------------------------------------------------------------ conversion (fichier de test)

        [Fact]
        public void ConvertsSalesSheetWithCaptionHiddenRowsAndFormats()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            var result = Import(wb, "Ventes 2024");
            Assert.True(result.Succeeded, result.Error);
            Assert.Equal("Tableau 1 : Ventes trimestrielles par région", result.Caption);
            Assert.Equal("Ventes trimestrielles par région", SheetConverter.CaptionTitle(result.Caption));
            var t = result.Table;
            Assert.Equal(9, t.RowCount);      // lignes 3 à 12, sans la ligne 9 masquée
            Assert.Equal(7, t.ColumnCount);   // A à G, colonne H masquée

            Assert.Equal("Région", Text(t, 0, 0));
            var header = TestUtil.At(t, 0, 1);
            Assert.Equal(TestUtil.Hex("4472C4"), header.Fill);
            Assert.True(header.PrimaryFormat.Bold);
            Assert.Equal(Rgb.White, header.PrimaryFormat.Color);
            Assert.Equal(HorizontalAlignment.Center, header.HorizontalAlignment);
            Assert.Equal(30, t.RowHeightsPt[0]);

            Assert.Equal("125 430,50 €", Nb(Text(t, 1, 1)));
            Assert.Equal(HorizontalAlignment.Right, TestUtil.At(t, 1, 1).HorizontalAlignment);
            Assert.Equal("528 281,50 €", Nb(Text(t, 1, 5)));
            Assert.Equal("8,3%", Text(t, 1, 6));

            // Évolution négative : rouge du format de nombre et fond de la mise en forme conditionnelle.
            var negative = TestUtil.At(t, 2, 6);
            Assert.Equal("-1,1%", negative.PlainText);
            Assert.Equal(new Rgb(255, 0, 0), negative.PrimaryFormat.Color);
            Assert.Equal(TestUtil.Hex("FFC7CE"), negative.Fill);

            // La ligne masquée n'apparaît pas : la ligne « Total » suit « Hauts-de-France ».
            Assert.Equal("Hauts-de-France", Text(t, 5, 0));
            Assert.Equal("Total", Text(t, 6, 0));
            var total = TestUtil.At(t, 6, 1);
            Assert.Equal(BorderStyle.Double, total.Borders.Top.Style);
            Assert.Equal(BorderStyle.Medium, total.Borders.Bottom.Style);
            Assert.Equal(TestUtil.Hex("1F3864"), total.Borders.Bottom.Color);
            // Bordure partagée : le bas de la ligne précédente reçoit la même double bordure.
            Assert.Equal(BorderStyle.Double, TestUtil.At(t, 5, 1).Borders.Bottom.Style);

            // La note de bas de tableau déborde dans Excel : cellule fusionnée avec les cellules vides voisines.
            var note = TestUtil.At(t, 8, 0);
            Assert.StartsWith("Source : service commercial", note.PlainText);
            Assert.True(note.ColumnSpan > 1);
            Assert.True(note.PrimaryFormat.Italic);

            Assert.Contains(result.Warnings, w => w.Contains("ligne masquée"));
        }

        [Fact]
        public void ConvertsPlanningSheetValuesExactlyAsDisplayed()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            var result = Import(wb, "Planning");
            Assert.True(result.Succeeded, result.Error);
            var t = result.Table;
            Assert.Null(result.Caption);

            var phase = TestUtil.At(t, 1, 0);
            Assert.Equal("Phase 1", phase.PlainText);
            Assert.Equal(3, phase.RowSpan);
            Assert.Equal(VerticalAlignment.Center, phase.VerticalAlignment);

            Assert.Equal("lundi 8 janvier 2024", Text(t, 1, 2));
            Assert.Equal("19/01/2024 17:30", Text(t, 1, 3));
            Assert.Equal("36:00", Text(t, 1, 4));
            Assert.Equal("VRAI", Text(t, 1, 5));
            Assert.Equal(HorizontalAlignment.Center, TestUtil.At(t, 1, 5).HorizontalAlignment);
            Assert.Equal("100%", Text(t, 1, 6));
            Assert.Equal("Ateliers avec les utilisateurs\net les métiers (4 sessions)", Text(t, 2, 7));
            Assert.Equal("#DIV/0!", Text(t, 4, 7));
            Assert.Equal(90, TestUtil.At(t, 0, 5).TextRotation);
            Assert.Equal(1, TestUtil.At(t, 1, 1).IndentLevel);

            var note = TestUtil.At(t, 6, 1);
            Assert.Equal(7, note.ColumnSpan);
            Assert.Equal("Attention : livrable en retard de deux semaines", note.PlainText);
            var attention = note.Runs[0];
            Assert.True(attention.Format.Bold);
            Assert.Equal(TestUtil.Hex("C00000"), attention.Format.Color);
            Assert.Contains(note.Runs, r => r.Text == "deux semaines" && r.Format.Italic && r.Format.Underline == UnderlineKind.Single);

            Assert.Equal("(1 234,50)", Nb(Text(t, 7, 1)));
            Assert.Equal(new Rgb(255, 0, 0), TestUtil.At(t, 7, 1).PrimaryFormat.Color);
            Assert.Equal("1,23E-04", Text(t, 7, 2));
            Assert.Equal("3 14/99", Text(t, 7, 3));
            Assert.Equal("1,23457E+12", Text(t, 7, 4));
            Assert.Equal("FAUX", Text(t, 7, 5));
            Assert.Equal("0612345678", Text(t, 7, 6));
            Assert.Equal("0,3", Text(t, 7, 7));
        }

        [Fact]
        public void EmptyHiddenAndTooWideSheets()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            Assert.Equal("La feuille est vide.", Import(wb, "Vide").Error);
            var wide = Import(wb, "Très large");
            Assert.False(wide.Succeeded);
            Assert.Contains("63", wide.Error);
            // Une plage plus étroite indiquée par l'utilisateur est acceptée.
            Assert.True(Import(wb, "Très large", range: new CellRange(0, 0, 0, 9)).Succeeded);
            // Le texte de la feuille masquée déborde sur la colonne B : elle fait partie de la plage.
            var hidden = Import(wb, "Paramètres");
            Assert.True(hidden.Succeeded);
            Assert.Equal("Feuille masquée", Text(hidden.Table, 0, 0));
        }

        [Fact]
        public void FormulasWithoutStoredValuesAreReported()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe_sans_calcul.xlsx"));
            var result = Import(wb, "Ventes 2024");
            Assert.True(result.Succeeded);
            Assert.Contains(result.Warnings, w => w.Contains("pas de résultat enregistré"));
            Assert.Equal(string.Empty, Text(result.Table, 1, 5));
        }

        [Fact]
        public void ExcelTableStyleIsReproduced()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe_sans_calcul.xlsx"));
            var t = Import(wb, "Achats").Table;
            var accent = wb.Styles.Colors.Accent(1);
            var header = TestUtil.At(t, 0, 0);
            Assert.Equal(accent, header.Fill);
            Assert.True(header.PrimaryFormat.Bold);
            Assert.Equal(Rgb.White, header.PrimaryFormat.Color);
            Assert.Equal(1, t.HeaderRowCount);
            // Bandes : première ligne de données colorée, deuxième non.
            Assert.NotNull(TestUtil.At(t, 1, 0).Fill);
            Assert.Null(TestUtil.At(t, 2, 0).Fill);
            Assert.Equal("189,99 €", Nb(Text(t, 4, 2)));
        }

        [Fact]
        public void FitsWideTablesWithoutBreakingNumbers()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            var options = new ExcelImportOptions { AvailableWidthPt = 480 };
            var result = Import(wb, "Ventes 2024", options: options);
            var t = result.Table;
            Assert.True(t.ColumnWidthsPt.Sum() <= 480.5);
            Assert.Contains(result.Warnings, w => w.Contains("réduite"));
            Assert.DoesNotContain(result.Warnings, w => w.Contains("reste plus large"));
            // Les colonnes de montants restent assez larges pour leur valeur la plus longue ;
            // c'est la colonne de texte (« Région ») qui cède, ses mots passant à la ligne.
            // Largeurs Excel : 13 caractères = 91 pixels = 68,25 pt ; 15 caractères = 105 pixels = 78,75 pt.
            for (int c = 1; c <= 4; c++) Assert.True(t.ColumnWidthsPt[c] >= 68.25 - 0.01, "Colonne " + c + " réduite : " + t.ColumnWidthsPt[c]);
            Assert.True(t.ColumnWidthsPt[5] >= 78.75 - 0.01, "Colonne Total réduite : " + t.ColumnWidthsPt[5]);
            Assert.True(t.ColumnWidthsPt[0] < 24 * 7 * 0.75);
        }

        // ------------------------------------------------------------------ cas construits

        private const string TwoStyles = "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"#,##0.00\\ &quot;€&quot;\"/></numFmts>"
            + "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"14\"/><color rgb=\"FF1F4E79\"/><name val=\"Arial\"/></font></fonts>"
            + "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
            + "<fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.79998168889431442\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
            + "<borders count=\"2\"><border/><border><left style=\"thin\"><color indexed=\"64\"/></left><right style=\"thin\"><color indexed=\"64\"/></right>"
            + "<top style=\"thin\"><color indexed=\"64\"/></top><bottom style=\"thin\"><color indexed=\"64\"/></bottom></border></borders>"
            + "<cellXfs count=\"5\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
            + "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"1\" applyNumberFormat=\"1\"/>"
            + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\"><alignment horizontal=\"centerContinuous\"/></xf>"
            + "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
            + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"centerContinuous\"/></xf></cellXfs>"
            + "<dxfs count=\"1\"><dxf><font><b/><color rgb=\"FF9C0006\"/></font><fill><patternFill><bgColor rgb=\"FFFFC7CE\"/></patternFill></fill></dxf></dxfs>";

        [Fact]
        public void CenterAcrossSelectionBecomesAMergedCell()
        {
            var b = new XlsxBuilder { Styles = TwoStyles };
            b.SharedStrings.Add("Titre centré");
            b.SharedStrings.Add("x");
            b.Sheet("F", "<sheetData><row r=\"1\"><c r=\"A1\" s=\"2\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"4\"/><c r=\"C1\" s=\"4\"/></row>"
                       + "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>1</v></c><c r=\"B2\" s=\"1\"><v>2</v></c><c r=\"C2\" s=\"1\"><v>3</v></c></row></sheetData>");
            var wb = XlsxWorkbook.Load(b.Build());
            var t = Import(wb, "F").Table;
            var title = TestUtil.At(t, 0, 0);
            Assert.Equal(3, title.ColumnSpan);
            Assert.Equal(HorizontalAlignment.Center, title.HorizontalAlignment);
            Assert.Equal("Arial", title.PrimaryFormat.FontName);
            Assert.Equal(14, title.PrimaryFormat.Size);
            Assert.Equal(TestUtil.Hex("1F4E79"), title.PrimaryFormat.Color);
            // Couleur de thème avec nuance (Accent 1 éclairci à 80 %).
            Assert.Equal(ExcelColors.ApplyTint(wb.Styles.Colors.Accent(1), 0.79998168889431442), title.Fill);
            Assert.Equal("1,00 €", Nb(Text(t, 1, 0)));
        }

        [Fact]
        public void ConditionalFormattingRules()
        {
            var b = new XlsxBuilder { Styles = TwoStyles };
            var rows = new StringBuilder("<sheetData>");
            for (int r = 1; r <= 6; r++) rows.Append("<row r=\"" + r + "\"><c r=\"A" + r + "\"><v>" + (r * 10) + "</v></c><c r=\"B" + r + "\"><v>" + r + "</v></c></row>");
            rows.Append("</sheetData>");
            rows.Append("<conditionalFormatting sqref=\"A1:A6\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"2\" operator=\"greaterThan\"><formula>$B$1*40</formula></cfRule></conditionalFormatting>");
            rows.Append("<conditionalFormatting sqref=\"B1:B6\"><cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>");
            rows.Append("<conditionalFormatting sqref=\"A1:B6\"><cfRule type=\"dataBar\" priority=\"3\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar></cfRule></conditionalFormatting>");
            b.Sheet("F", rows.ToString());
            var wb = XlsxWorkbook.Load(b.Build());
            var result = Import(wb, "F", EnUs);
            var t = result.Table;
            // A : valeurs > 40 (B1 × 40) en rouge gras sur fond rose.
            Assert.Null(TestUtil.At(t, 3, 0).Fill);
            Assert.Equal(TestUtil.Hex("FFC7CE"), TestUtil.At(t, 4, 0).Fill);
            Assert.True(TestUtil.At(t, 4, 0).PrimaryFormat.Bold);
            Assert.Equal(TestUtil.Hex("9C0006"), TestUtil.At(t, 4, 0).PrimaryFormat.Color);
            // B : lignes paires.
            Assert.Null(TestUtil.At(t, 0, 1).Fill);
            Assert.Equal(TestUtil.Hex("FFC7CE"), TestUtil.At(t, 1, 1).Fill);
            // Barres de données : signalées comme non reproduites.
            Assert.Contains(result.Warnings, w => w.Contains("mise en forme conditionnelle"));
        }

        [Fact]
        public void ColorScaleAndTop10()
        {
            var b = new XlsxBuilder();
            var rows = new StringBuilder("<sheetData>");
            for (int r = 1; r <= 5; r++) rows.Append("<row r=\"" + r + "\"><c r=\"A" + r + "\"><v>" + (r - 1) * 25 + "</v></c></row>");
            rows.Append("</sheetData><conditionalFormatting sqref=\"A1:A5\"><cfRule type=\"colorScale\" priority=\"1\"><colorScale>"
                      + "<cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFFFFFF\"/><color rgb=\"FF000000\"/></colorScale></cfRule></conditionalFormatting>");
            b.Sheet("F", rows.ToString());
            var t = Import(XlsxWorkbook.Load(b.Build()), "F", EnUs).Table;
            Assert.Equal(Rgb.White, TestUtil.At(t, 0, 0).Fill);
            Assert.Equal(new Rgb(128, 128, 128), TestUtil.At(t, 2, 0).Fill);
            Assert.Equal(Rgb.Black, TestUtil.At(t, 4, 0).Fill);
        }

        [Fact]
        public void PrintAreaHiddenRowsAndDefaultWidths()
        {
            var b = new XlsxBuilder();
            b.SharedStrings.Add("a");
            b.DefinedNames = "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Ma feuille'!$B$2:$C$4</definedName>"
                           + "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">'Ma feuille'!$2:$2</definedName>";
            var data = new StringBuilder("<cols><col min=\"2\" max=\"2\" width=\"20.7109375\" customWidth=\"1\"/></cols><sheetData>");
            for (int r = 1; r <= 5; r++)
            {
                data.Append("<row r=\"" + r + "\"" + (r == 3 ? " hidden=\"1\"" : string.Empty) + (r == 4 ? " ht=\"30\" customHeight=\"1\"" : string.Empty) + ">");
                for (int c = 0; c < 4; c++) data.Append("<c r=\"" + (char)('A' + c) + r + "\"><v>" + (r * 10 + c) + "</v></c>");
                data.Append("</row>");
            }
            data.Append("</sheetData>");
            b.Sheet("Ma feuille", data.ToString());
            var wb = XlsxWorkbook.Load(b.Build());
            Assert.Equal(new CellRange(1, 1, 3, 2), wb.Sheets[0].PrintArea);
            var result = Import(wb, "Ma feuille", EnUs);
            var t = result.Table;
            Assert.Equal(new CellRange(1, 1, 3, 2), result.Range);
            Assert.Equal(2, t.RowCount); // ligne 3 masquée
            Assert.Equal("21", Text(t, 0, 0));
            Assert.Equal("42", Text(t, 1, 1));
            Assert.Equal(1, t.HeaderRowCount);
            Assert.Equal(30, t.RowHeightsPt[1]);
            // Largeur enregistrée 20,7109375 (20 caractères affichés, Calibri 11) = 145 pixels = 108,75 pt ; largeur standard = 64 pixels = 48 pt.
            Assert.Equal(108.75, t.ColumnWidthsPt[0], 2);
            Assert.Equal(48, t.ColumnWidthsPt[1], 1);
        }

        [Fact]
        public void InlineStringsBooleansErrorsAndDates()
        {
            var b = new XlsxBuilder { Styles = TwoStyles };
            b.Sheet("F", "<sheetData><row r=\"1\">"
                       + "<c r=\"A1\" t=\"inlineStr\"><is><r><rPr><b/><sz val=\"11\"/><rFont val=\"Calibri\"/></rPr><t>Gras</t></r><r><t xml:space=\"preserve\"> normal</t></r></is></c>"
                       + "<c r=\"B1\" t=\"b\"><v>1</v></c><c r=\"C1\" t=\"e\"><v>#N/A</v></c><c r=\"D1\" s=\"3\"><v>45366</v></c>"
                       + "<c r=\"E1\" t=\"d\" s=\"3\"><v>2024-03-15T00:00:00</v></c><c r=\"F1\" t=\"str\"><f>A1</f><v>calculé</v></c>"
                       + "<c r=\"G1\"><v>1E-3</v></c></row></sheetData>");
            var wb = XlsxWorkbook.Load(b.Build());
            var t = Import(wb, "F", FrFr).Table;
            var a1 = TestUtil.At(t, 0, 0);
            Assert.Equal("Gras normal", a1.PlainText);
            Assert.True(a1.Runs[0].Format.Bold);
            Assert.False(a1.Runs[1].Format.Bold);
            Assert.Equal("VRAI", Text(t, 0, 1));
            Assert.Equal("#N/A", Text(t, 0, 2));
            Assert.Equal("15/03/2024", Text(t, 0, 3));
            Assert.Equal("15/03/2024", Text(t, 0, 4));
            Assert.Equal("calculé", Text(t, 0, 5));
            Assert.Equal("0,001", Text(t, 0, 6));
            Assert.Equal("TRUE", Text(Import(wb, "F", EnUs).Table, 0, 1));
        }

        [Fact]
        public void MergedCellBordersComeFromTheEdgeCells()
        {
            var b = new XlsxBuilder { Styles = TwoStyles };
            b.SharedStrings.Add("Fusion");
            b.Sheet("F", "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"1\"/></row><row r=\"2\"><c r=\"A2\" s=\"1\"/><c r=\"B2\" s=\"1\"/></row>"
                       + "<row r=\"3\"><c r=\"A3\"><v>1</v></c></row></sheetData><mergeCells count=\"1\"><mergeCell ref=\"A1:B2\"/></mergeCells>");
            var t = Import(XlsxWorkbook.Load(b.Build()), "F").Table;
            var merged = TestUtil.At(t, 0, 0);
            Assert.Equal(2, merged.RowSpan);
            Assert.Equal(2, merged.ColumnSpan);
            Assert.Equal(BorderStyle.Thin, merged.Borders.Right.Style);
            Assert.Equal(BorderStyle.Thin, merged.Borders.Bottom.Style);
            // La cellule en dessous reçoit la bordure partagée.
            Assert.Equal(BorderStyle.Thin, TestUtil.At(t, 2, 0).Borders.Top.Style);
        }

        // ------------------------------------------------------------------ aller-retour Word → Excel → Word

        [Fact]
        public void ExportedWorkbookImportsBackIdentically()
        {
            var source = XlsxWriterTests.SampleTable();
            var writer = new XlsxWorkbookWriter(new XlsxExportOptions { IncludeCaptionRow = true, NumberCulture = FrFr });
            writer.AddTable(source);
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                writer.Save(ms);
                bytes = ms.ToArray();
            }

            var wb = XlsxWorkbook.Load(bytes);
            var result = Import(wb, source.SheetName);
            Assert.True(result.Succeeded, result.Error);
            Assert.Equal(source.Caption, result.Caption);
            var t = result.Table;
            Assert.Equal(source.RowCount, t.RowCount);
            Assert.Equal(source.ColumnCount, t.ColumnCount);
            foreach (var original in source.Cells)
            {
                var imported = TestUtil.At(t, original.Row, original.Column);
                Assert.Equal(original.Row, imported.Row);
                Assert.Equal(original.Column, imported.Column);
                Assert.Equal(original.RowSpan, imported.RowSpan);
                Assert.Equal(original.ColumnSpan, imported.ColumnSpan);
                Assert.Equal(original.PlainText, imported.PlainText);
                Assert.Equal(original.Fill, imported.Fill);
                Assert.Equal(original.TextRotation, imported.TextRotation);
            }
            var nord = TestUtil.At(t, 1, 0);
            Assert.Equal(3, nord.Runs.Count);
            Assert.True(nord.Runs[1].Format.Italic);
            Assert.Equal(TestUtil.Hex("FF0000"), nord.Runs[1].Format.Color);
            Assert.Equal(VerticalPosition.Superscript, nord.Runs[2].Format.Position);
            Assert.Equal(BorderStyle.Thick, TestUtil.At(t, 1, 1).Borders.Bottom.Style);
            Assert.Equal("Arial", TestUtil.At(t, 0, 0).PrimaryFormat.FontName);
        }

        // ------------------------------------------------------------------ XML Word généré

        [Fact]
        public void GeneratedWordXmlRoundTripsThroughTheWordParser()
        {
            var source = XlsxWriterTests.SampleTable();
            source.HeaderRowCount = 1;
            string xml = WordTableXmlWriter.BuildDocumentXml(source);
            var layout = Core.Layout.WordXmlTableParser.Parse(xml);
            Assert.Equal(source.RowCount, layout.RowCount);
            Assert.Equal(source.ColumnCount, layout.ColumnCount);
            foreach (var original in source.Cells)
            {
                var cell = TestUtil.At(layout, original.Row, original.Column);
                Assert.Equal(original.Row, cell.Row);
                Assert.Equal(original.Column, cell.Column);
                Assert.Equal(original.RowSpan, cell.RowSpan);
                Assert.Equal(original.ColumnSpan, cell.ColumnSpan);
                Assert.Equal(original.Fill, cell.Fill);
                Assert.Equal(original.TextRotation, cell.TextRotation);
                Assert.Equal(original.VerticalAlignment, cell.VerticalAlignment);
            }
            Assert.Equal(BorderStyle.Thick, TestUtil.At(layout, 1, 1).Borders.Bottom.Style);
            Assert.Equal(TestUtil.Hex("FF0000"), TestUtil.At(layout, 1, 1).Borders.Bottom.Color);
            Assert.Equal(100, layout.ColumnWidthsPt[0], 1);

            var doc = System.Xml.Linq.XDocument.Parse(xml);
            System.Xml.Linq.XNamespace w = WordTableXmlWriter.WordNamespace;
            Assert.Single(doc.Descendants(w + "tblHeader"));
            Assert.Equal("fixed", doc.Descendants(w + "tblLayout").Single().Attribute(w + "type").Value);
            // Saut de ligne du texte « \nm² » et exposant.
            Assert.NotEmpty(doc.Descendants(w + "br"));
            Assert.Contains(doc.Descendants(w + "vertAlign"), e => e.Attribute(w + "val").Value == "superscript");
            // Caractères spéciaux échappés, pas de perte de texte.
            Assert.Contains(doc.Descendants(w + "t"), e => e.Value == "Région & ventes <2023>");
        }

        [Fact]
        public void FlatOpcAndDocxPackagesAreWellFormed()
        {
            var table = XlsxWriterTests.SampleTable();
            var flat = System.Xml.Linq.XDocument.Parse(WordTableXmlWriter.BuildFlatOpc(table));
            System.Xml.Linq.XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            var parts = flat.Root.Elements(pkg + "part").Select(p => p.Attribute(pkg + "name").Value).ToList();
            Assert.Contains("/word/document.xml", parts);
            Assert.Contains("/_rels/.rels", parts);

            var zip = new ZipReader(WordTableXmlWriter.BuildDocx(table));
            Assert.True(zip.Contains("word/document.xml"));
            Assert.True(zip.Contains("[Content_Types].xml"));
            var layout = Core.Layout.WordXmlTableParser.Parse(Encoding.UTF8.GetString(zip.Read("word/document.xml")));
            Assert.Equal(table.RowCount, layout.RowCount);
        }

        [Fact]
        public void EndToEndImportProducesValidWordTables()
        {
            var wb = XlsxWorkbook.Load(FixturePath("import_complexe.xlsx"));
            foreach (var info in wb.Sheets)
            {
                var result = Import(wb, info.Name, FrFr, new ExcelImportOptions { AvailableWidthPt = 453 });
                if (!result.Succeeded) continue;
                string xml = WordTableXmlWriter.BuildFlatOpc(result.Table);
                System.Xml.Linq.XDocument.Parse(xml);
                var layout = Core.Layout.WordXmlTableParser.Parse(WordTableXmlWriter.BuildDocumentXml(result.Table));
                Assert.Equal(result.Table.RowCount, layout.RowCount);
                Assert.Equal(result.Table.ColumnCount, layout.ColumnCount);
                Assert.Equal(result.Table.Cells.Count, layout.Cells.Count);

                // Aller-retour Excel → Word → lecture rapide du XML (export de l'application) : mêmes valeurs,
                // même mise en forme des caractères.
                layout = Core.Layout.WordXmlTableParser.Parse(xml);
                Assert.True(layout.HasContent);
                foreach (var source in result.Table.Cells)
                {
                    var read = layout.Cells.Single(c => c.Row == source.Row && c.Column == source.Column);
                    var runs = Core.Text.CellTextSanitizer.Clean(read.Content.Runs);
                    Assert.Equal(source.PlainText.Replace(' ', ' ').Replace(' ', ' '), string.Concat(runs.Select(r => r.Text)).Replace(' ', ' ').Replace(' ', ' '));
                    var expected = Core.Text.CellTextSanitizer.Clean(source.Runs);
                    for (int i = 0; i < Math.Min(expected.Count, runs.Count); i++)
                    {
                        Assert.Equal(expected[i].Format.Bold, runs[i].Format.Bold);
                        Assert.Equal(expected[i].Format.Italic, runs[i].Format.Italic);
                        Assert.Equal(expected[i].Format.Color ?? Rgb.Black, runs[i].Format.Color ?? Rgb.Black);
                        Assert.Equal(expected[i].Format.FontName, runs[i].Format.FontName);
                        Assert.Equal(expected[i].Format.Size, runs[i].Format.Size);
                    }
                }
            }
        }
    }
}
