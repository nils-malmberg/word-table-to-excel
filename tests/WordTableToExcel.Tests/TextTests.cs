using System.Globalization;
using System.Linq;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Text;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class CellTextSanitizerTests
    {
        private static TextRun R(string text, bool bold = false)
        {
            return new TextRun(text, new RunFormat { Bold = bold }, null);
        }

        [Fact]
        public void ParagraphMarksAndLineBreaks()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("a\rb\u000Bc\r\a") });
            Assert.Equal("a\nb\nc", runs.Single().Text);
        }

        [Fact]
        public void FieldCodesAreRemoved_ResultsKept()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("x\u0013 SEQ Tableau \\* ARABIC \u00141\u0015y"), R("\u0013 PAGE \u0015z") });
            Assert.Equal("x1yz", string.Concat(runs.Select(r => r.Text)));
        }

        [Fact]
        public void FieldCodeSpanningRuns()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("a\u0013 HYPER"), R("LINK \"x\" \u0014lien", true), R("\u0015b") });
            Assert.Equal("a", runs[0].Text);
            Assert.Equal("lien", runs[1].Text);
            Assert.True(runs[1].Format.Bold);
            Assert.Equal("b", runs[2].Text);
        }

        [Fact]
        public void AdjacentRunsWithSameFormatAreMerged()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("ab"), R("cd"), R("ef", true), R("") });
            Assert.Equal(2, runs.Count);
            Assert.Equal("abcd", runs[0].Text);
        }

        [Fact]
        public void TrailingEmptyParagraphsAreRemoved()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("abc\r"), R("\r", true) });
            Assert.Single(runs);
            Assert.Equal("abc", runs[0].Text);
        }

        [Fact]
        public void InvalidCharactersAreRemoved()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("a\u0001b\u0008c\uD800d\u001Fe\u001Ef\tg") });
            Assert.Equal("abcde-f g", runs[0].Text);
        }

        [Fact]
        public void SurrogatePairsAreKept()
        {
            var runs = CellTextSanitizer.Clean(new[] { R("ok 😀") });
            Assert.Equal("ok 😀", runs[0].Text);
        }

        [Fact]
        public void TruncatedToExcelLimit()
        {
            var runs = CellTextSanitizer.Clean(new[] { R(new string('x', 40000)), R("y", true) });
            Assert.Single(runs);
            Assert.Equal(CellTextSanitizer.ExcelMaxCellLength, runs[0].Text.Length);
        }
    }

    public class NumberParserTests
    {
        private static readonly CultureInfo Fr = new CultureInfo("fr-FR");
        private static readonly CultureInfo En = new CultureInfo("en-US");

        [Theory]
        [InlineData("1 234,50", 1234.5, "#,##0.00")]
        [InlineData("1\u00A0234,50", 1234.5, "#,##0.00")]
        [InlineData("1\u202F234\u202F567", 1234567, "#,##0")]
        [InlineData("12,5 %", 0.125, "0.0%")]
        [InlineData("-5 %", -0.05, "0%")]
        [InlineData("\u22123,2", -3.2, "0.0")]
        [InlineData("2024", 2024, null)]
        [InlineData("0,5", 0.5, "0.0")]
        [InlineData("3.14", 3.14, "0.00")]
        [InlineData("45 €", 45, "0\" €\"")]
        [InlineData("1,234.56", 1234.56, "#,##0.00")]
        [InlineData("1.234,56", 1234.56, "#,##0.00")]
        [InlineData("+7", 7, null)]
        [InlineData("0", 0, null)]
        public void French(string text, double expected, string format)
        {
            double value;
            string numberFormat;
            Assert.True(NumberParser.TryParse(text, Fr, out value, out numberFormat), text);
            Assert.Equal(expected, value, 10);
            Assert.Equal(format, numberFormat);
        }

        [Theory]
        [InlineData("007")]
        [InlineData("01 23 45 67 89")]
        [InlineData("1.234")]
        [InlineData("12:30")]
        [InlineData("1 2")]
        [InlineData("abc")]
        [InlineData("")]
        [InlineData("1.2.3")]
        [InlineData("12,5,3")]
        [InlineData("2023-01-05")]
        [InlineData("1234567890123456")]
        [InlineData("N° 5")]
        public void FrenchRejected(string text)
        {
            double value;
            string numberFormat;
            Assert.False(NumberParser.TryParse(text, Fr, out value, out numberFormat), text);
        }

        [Theory]
        [InlineData("1,234", 1234, "#,##0")]
        [InlineData("1.234", 1.234, "0.000")]
        [InlineData("$1,200.00", 1200, "\"$\"#,##0.00")]
        [InlineData("-$5", -5, "\"$\"0")]
        public void English(string text, double expected, string format)
        {
            double value;
            string numberFormat;
            Assert.True(NumberParser.TryParse(text, En, out value, out numberFormat), text);
            Assert.Equal(expected, value, 10);
            Assert.Equal(format, numberFormat);
        }
    }
}
