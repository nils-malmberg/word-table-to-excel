using System.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;
using Xunit;
using static WordTableToExcel.Tests.TestUtil;

namespace WordTableToExcel.Tests
{
    public class WordXmlTableParserTests
    {
        private static readonly Rgb Accent1 = Hex("4F81BD");
        private static readonly Rgb Band = Hex("D3DFEE");
        private static readonly Rgb StyleBorder = Hex("7BA0CD");

        [Fact]
        public void Word2010FlatOpc_StructureAndMerges()
        {
            var layout = WordXmlTableParser.Parse(Fixture("medium_shading_merged.flat.xml"));

            Assert.NotNull(layout);
            Assert.True(layout.HasCellFormatting);
            Assert.Equal(5, layout.RowCount);
            Assert.Equal(4, layout.ColumnCount);
            Assert.Equal(18, layout.Cells.Count);

            var ventes = At(layout, 0, 1);
            Assert.Equal(2, ventes.ColumnSpan);
            Assert.Same(ventes, At(layout, 0, 2));

            var nord = At(layout, 1, 0);
            Assert.Equal(2, nord.RowSpan);
            Assert.Same(nord, At(layout, 2, 0));

            // Ligne 3 de Word : la continuation de fusion verticale est ignorée dans l'ordre des cellules visibles.
            var row2 = layout.CellsStartingOnSourceRow(2);
            Assert.Equal(3, row2.Count);
            Assert.Equal(1, row2[0].Column);
            Assert.Equal(0, row2[0].SourceOrdinal);
            Assert.Equal(1, row2[0].SourceIndex);
            Assert.Equal(4, layout.SourceCellCounts[2]);

            Assert.All(layout.ColumnWidthsPt, w => Assert.True(w > 50));
        }

        [Fact]
        public void Word2010FlatOpc_TableStyleShading()
        {
            var layout = WordXmlTableParser.Parse(Fixture("medium_shading_merged.flat.xml"));

            // Ligne d'en-tête du style.
            Assert.Equal(Accent1, At(layout, 0, 0).Fill);
            Assert.Equal(Accent1, At(layout, 0, 1).Fill);
            Assert.Equal(Accent1, At(layout, 0, 3).Fill);
            // Lignes à bandes : 1re ligne de données colorée, 2e non.
            Assert.Equal(Band, At(layout, 1, 1).Fill);
            Assert.Equal(Band, At(layout, 1, 0).Fill);
            Assert.Null(At(layout, 2, 1).Fill);
            Assert.Equal(Band, At(layout, 3, 1).Fill);
            // Trame directe prioritaire sur le style.
            Assert.Equal(Hex("FFFF00"), At(layout, 3, 3).Fill);
            // Dernière ligne (bande paire, sans trame).
            Assert.Null(At(layout, 4, 1).Fill);
        }

        [Fact]
        public void Word2010FlatOpc_Borders()
        {
            var layout = WordXmlTableParser.Parse(Fixture("medium_shading_merged.flat.xml"));

            var direct = At(layout, 1, 1).Borders.Bottom;
            Assert.Equal(BorderStyle.Thick, direct.Style);
            Assert.Equal(Hex("FF0000"), direct.Color);

            Assert.Equal(new BorderLine(BorderStyle.Thin, StyleBorder), At(layout, 0, 0).Borders.Top);
            Assert.False(At(layout, 0, 1).Borders.Left.IsVisible);             // insideV « nil » de l'en-tête
            Assert.Equal(BorderStyle.Double, At(layout, 4, 1).Borders.Top.Style); // haut de la dernière ligne
            Assert.Equal(new BorderLine(BorderStyle.Thin, StyleBorder), At(layout, 2, 1).Borders.Top); // insideH du tableau
            Assert.False(At(layout, 2, 1).Borders.Left.IsVisible);             // pas de bordure verticale intérieure
            Assert.True(At(layout, 2, 3).Borders.Right.IsVisible);             // bord droit du tableau
        }

        [Fact]
        public void Word2003Xml_LowercaseVmergeTextFlowAndExactHeight()
        {
            const string xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:wordDocument xmlns:w=""http://schemas.microsoft.com/office/word/2003/wordml"" xmlns:wx=""http://schemas.microsoft.com/office/word/2003/auxHint"">
 <w:styles>
  <w:style w:type=""table"" w:default=""on"" w:styleId=""TableNormal""><w:name w:val=""Normal Table""/><w:tblPr><w:tblInd w:w=""0"" w:type=""dxa""/></w:tblPr></w:style>
  <w:style w:type=""table"" w:styleId=""TableGrid""><w:name w:val=""Table Grid""/><w:basedOn w:val=""TableNormal""/>
   <w:tblPr><w:tblBorders>
    <w:top w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
    <w:left w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
    <w:bottom w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
    <w:right w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
    <w:insideH w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
    <w:insideV w:val=""single"" w:sz=""4"" wx:bdrwidth=""10"" w:space=""0"" w:color=""auto""/>
   </w:tblBorders></w:tblPr></w:style>
 </w:styles>
 <w:body><wx:sect>
  <w:tbl>
   <w:tblPr><w:tblStyle w:val=""TableGrid""/><w:tblW w:w=""0"" w:type=""auto""/><w:tblLook w:val=""01E0""/></w:tblPr>
   <w:tblGrid><w:gridCol w:w=""2880""/><w:gridCol w:w=""2880""/><w:gridCol w:w=""2880""/></w:tblGrid>
   <w:tr><w:trPr><w:trHeight w:val=""567"" w:h-rule=""exact""/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w=""2880"" w:type=""dxa""/><w:vmerge w:val=""restart""/><w:shd w:val=""clear"" w:color=""auto"" w:fill=""C0C0C0""/><w:vAlign w:val=""center""/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w=""5760"" w:type=""dxa""/><w:gridSpan w:val=""2""/><w:textFlow w:val=""bt-lr""/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
   </w:tr>
   <w:tr>
    <w:tc><w:tcPr><w:tcW w:w=""2880"" w:type=""dxa""/><w:vmerge/></w:tcPr><w:p/></w:tc>
    <w:tc><w:tcPr><w:tcW w:w=""2880"" w:type=""dxa""/><w:shd w:val=""pct50"" w:color=""000000"" w:fill=""FFFFFF""/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w=""2880"" w:type=""dxa""/><w:tcBorders><w:bottom w:val=""nil""/></w:tcBorders></w:tcPr><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
   </w:tr>
  </w:tbl>
 </wx:sect></w:body>
</w:wordDocument>";

            var layout = WordXmlTableParser.Parse(xml);

            Assert.Equal(2, layout.RowCount);
            Assert.Equal(3, layout.ColumnCount);
            Assert.Equal(4, layout.Cells.Count);

            var a = At(layout, 0, 0);
            Assert.Equal(2, a.RowSpan);
            Assert.Equal(Hex("C0C0C0"), a.Fill);
            Assert.Equal(VerticalAlignment.Center, a.VerticalAlignment);

            var b = At(layout, 0, 1);
            Assert.Equal(2, b.ColumnSpan);
            Assert.Equal(90, b.TextRotation);

            Assert.Equal(28.35, layout.RowHeightsPt[0], 2);
            Assert.True(layout.RowHeightExact[0]);
            Assert.Equal(0, layout.RowHeightsPt[1]);

            Assert.Equal(new BorderLine(BorderStyle.Thin, null), a.Borders.Top);   // style « Table Grid »
            Assert.True(At(layout, 1, 1).Borders.Left.IsVisible);                  // insideV
            Assert.False(At(layout, 1, 2).Borders.Bottom.IsVisible);               // nil direct
            Assert.Equal(new Rgb(128, 128, 128), At(layout, 1, 1).Fill);           // motif 50 %
        }

        [Fact]
        public void LibreOfficeWord2003Export_IsParsedRobustly()
        {
            var layout = WordXmlTableParser.Parse(Fixture("libreoffice_word2003.xml"));

            Assert.NotNull(layout);
            Assert.Equal(5, layout.RowCount);
            Assert.Equal(4, layout.ColumnCount);
            Assert.Equal(2, At(layout, 0, 1).ColumnSpan);
            Assert.Equal(Accent1, At(layout, 0, 0).Fill);   // trame « solid »
            Assert.Equal(Band, At(layout, 1, 1).Fill);
            Assert.Equal(BorderStyle.Thick, At(layout, 1, 1).Borders.Bottom.Style);
            Assert.All(layout.ColumnWidthsPt, w => Assert.InRange(w, 100, 120)); // largeurs décimales
        }

        [Fact]
        public void RowsInsideContentControls_AreFound()
        {
            const string xml = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=""1000""/><w:gridCol w:w=""1000""/></w:tblGrid>
 <w:tr><w:tc><w:p/></w:tc><w:sdt><w:sdtPr/><w:sdtContent><w:tc><w:p/></w:tc></w:sdtContent></w:sdt></w:tr>
 <w:sdt><w:sdtPr/><w:sdtContent>
  <w:tr><w:trPr><w:gridBefore w:val=""1""/></w:trPr><w:tc><w:tcPr><w:tcBorders><w:top w:val=""dashed"" w:sz=""12"" w:color=""00FF00""/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
 </w:sdtContent></w:sdt>
</w:tbl></w:body></w:document>";

            var layout = WordXmlTableParser.Parse(xml);

            Assert.Equal(2, layout.RowCount);
            Assert.Equal(2, layout.ColumnCount);
            Assert.Equal(3, layout.Cells.Count);
            var shifted = layout.Cells.Single(c => c.Row == 1);
            Assert.Equal(1, shifted.Column);
            Assert.Equal(new BorderLine(BorderStyle.MediumDashed, Hex("00FF00")), shifted.Borders.Top);
            // Sans style de tableau ni bordure : aucune bordure.
            Assert.False(At(layout, 0, 0).Borders.Top.IsVisible);
        }

        [Fact]
        public void NestedTable_OnlyOuterStructureIsUsed()
        {
            const string xml = @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body>
<w:tbl><w:tblGrid><w:gridCol w:w=""2000""/></w:tblGrid>
 <w:tr><w:tc><w:tbl><w:tblGrid><w:gridCol w:w=""500""/><w:gridCol w:w=""500""/></w:tblGrid><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr>
</w:tbl></w:body></w:document>";

            var layout = WordXmlTableParser.Parse(xml);

            Assert.Equal(1, layout.RowCount);
            Assert.Equal(1, layout.ColumnCount);
            Assert.Single(layout.Cells);
        }

        [Fact]
        public void NoTable_ReturnsNull()
        {
            Assert.Null(WordXmlTableParser.Parse(@"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body><w:p/></w:body></w:document>"));
            Assert.Null(WordXmlTableParser.Parse(""));
        }

        [Theory]
        [InlineData("single", 4, BorderStyle.Thin)]
        [InlineData("single", 12, BorderStyle.Medium)]
        [InlineData("single", 24, BorderStyle.Thick)]
        [InlineData("double", 4, BorderStyle.Double)]
        [InlineData("thinthicksmallgap", 12, BorderStyle.Double)]
        [InlineData("dotted", 4, BorderStyle.Dotted)]
        [InlineData("dashed", 4, BorderStyle.Dashed)]
        [InlineData("dotdash", 16, BorderStyle.MediumDashDot)]
        [InlineData("dotdotdash", 4, BorderStyle.DashDotDot)]
        [InlineData("wave", 6, BorderStyle.Thin)]
        public void BorderStyleMapping(string val, int eighths, BorderStyle expected)
        {
            Assert.Equal(expected, OoxmlFormat.MapBorderStyle(val, eighths));
        }
    }
}
