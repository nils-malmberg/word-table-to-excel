using System.Globalization;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class ExcelNumberFormatTests
    {
        private static readonly CultureInfo EnUs = CultureInfo.GetCultureInfo("en-US");
        private static readonly CultureInfo FrFr = CultureInfo.GetCultureInfo("fr-FR");

        private static string F(double value, string code, CultureInfo culture = null)
        {
            var settings = new ExcelFormatSettings(culture ?? EnUs, culture ?? EnUs, false, null);
            return settings.Get(code).FormatNumber(value, settings).Text;
        }

        private static string T(string text, string code)
        {
            var settings = new ExcelFormatSettings(EnUs, EnUs, false, null);
            return settings.Get(code).FormatText(text, settings).Text;
        }

        [Theory]
        [InlineData(0, "0")]
        [InlineData(1, "1")]
        [InlineData(-1, "-1")]
        [InlineData(0.5, "0.5")]
        [InlineData(1234.5678, "1234.5678")]
        [InlineData(0.30000000000000004, "0.3")]
        [InlineData(1.0 / 3, "0.333333333")]
        [InlineData(2.0 / 3, "0.666666667")]
        [InlineData(12345678901, "12345678901")]
        [InlineData(123456789012, "1.23457E+11")]
        [InlineData(100000000000, "1E+11")]
        [InlineData(1234567.891234, "1234567.891")]
        [InlineData(0.0001, "0.0001")]
        [InlineData(0.000012345, "0.000012345")]
        [InlineData(0.0000123456, "1.23456E-05")]
        [InlineData(1e-10, "1E-10")]
        [InlineData(-0.0000123456, "-1.23456E-05")]
        public void GeneralFormatFollowsExcelElevenCharacterRule(double value, string expected)
        {
            Assert.Equal(expected, F(value, "General"));
        }

        [Theory]
        [InlineData(1.005, "0.00", "1.01")]              // arrondi sur la valeur décimale (15 chiffres), comme Excel
        [InlineData(2.675, "0.00", "2.68")]
        [InlineData(99.995, "0.00", "100.00")]
        [InlineData(0.5, "0", "1")]
        [InlineData(-0.5, "0", "-1")]
        [InlineData(1234567.891, "#,##0.00", "1,234,567.89")]
        [InlineData(1234567, "#,###", "1,234,567")]
        [InlineData(0, "#,###", "")]
        [InlineData(5, "00000", "00005")]
        [InlineData(123456789, "000-00-0000", "123-45-6789")]
        [InlineData(5551234567, "(###) ###-####", "(555) 123-4567")]
        [InlineData(1.5, "0.0#", "1.5")]
        [InlineData(1.25, "0.0#", "1.25")]
        [InlineData(1, "#.##", "1.")]
        [InlineData(0.5, "#.##", ".5")]
        [InlineData(1.5, "???.???", "1.5")]
        [InlineData(1234567, "#,##0,", "1,235")]
        [InlineData(1234567, "0.0,,\"M\"", "1.2M")]
        [InlineData(0.256, "0%", "26%")]
        [InlineData(0.25678, "0.00%", "25.68%")]
        [InlineData(1.005, "0%", "101%")]
        [InlineData(-0.001, "0.00", "-0.00")]            // Excel conserve le signe d'une valeur négative arrondie à zéro
        public void FixedFormats(double value, string code, string expected)
        {
            Assert.Equal(expected, F(value, code));
        }

        [Theory]
        [InlineData(12345, "0.00E+00", "1.23E+04")]
        [InlineData(0.00012345, "0.00E+00", "1.23E-04")]
        [InlineData(0, "0.00E+00", "0.00E+00")]
        [InlineData(9.999, "0.00E+00", "1.00E+01")]
        [InlineData(-12345, "0.00E-00", "-1.23E04")]
        [InlineData(12345, "##0.0E+0", "12.3E+3")]
        [InlineData(0.00012345, "##0.0E+0", "123.5E-6")]
        [InlineData(12345, "00.00E+00", "01.23E+04")]
        public void ScientificFormats(double value, string code, string expected)
        {
            Assert.Equal(expected, F(value, code));
        }

        [Theory]
        [InlineData(0.5, "# ?/?", "1/2")]
        [InlineData(1.5, "# ?/?", "1 1/2")]
        [InlineData(2, "# ?/?", "2")]
        [InlineData(0, "# ?/?", "0")]
        [InlineData(0.333, "# ?/?", "1/3")]
        [InlineData(3.14159265, "# ??/??", "3 14/99")]
        [InlineData(3.14159265, "# ???/???", "3 16/113")]
        [InlineData(2.5, "?/?", "5/2")]
        [InlineData(0.5, "# ?/8", "4/8")]
        [InlineData(1.3, "# ??/100", "1 30/100")]
        [InlineData(-1.5, "# ?/?", "-1 1/2")]
        [InlineData(0.99, "# ?/8", "1")]
        public void FractionFormats(double value, string code, string expected)
        {
            Assert.Equal(expected, F(value, code));
        }

        [Fact]
        public void SectionsAndColors()
        {
            var settings = new ExcelFormatSettings(EnUs, EnUs, false, null);
            var format = settings.Get("[Blue]0.00;[Red]-0.00;\"zero\";\"text: \"@");
            var positive = format.FormatNumber(3, settings);
            var negative = format.FormatNumber(-3, settings);
            Assert.Equal("3.00", positive.Text);
            Assert.Equal(new Rgb(0, 0, 255), positive.Color);
            Assert.Equal("-3.00", negative.Text);
            Assert.Equal(new Rgb(255, 0, 0), negative.Color);
            Assert.Equal("zero", format.FormatNumber(0, settings).Text);
            Assert.Equal("text: abc", format.FormatText("abc", settings).Text);

            Assert.Equal("(1,234.57)", F(-1234.567, "#,##0.00;(#,##0.00)"));
            Assert.Equal("", F(5, ";;;"));
        }

        [Fact]
        public void AccountingFormatsDropAlignmentPadding()
        {
            const string accounting = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
            Assert.Equal("1,234.57", F(1234.567, accounting));
            Assert.Equal("(1,234.57)", F(-1234.567, accounting));
            Assert.Equal("-", F(0, accounting));
            Assert.Equal("$1,234.57", F(1234.567, "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)"));
            Assert.Equal("-$5.00", F(-5, "\"$\"#,##0.00"));
        }

        [Fact]
        public void ConditionalSections()
        {
            Assert.Equal("1,235K", F(1234567, "[>=1000000]#,##0,\"K\";0"));
            Assert.Equal("12", F(12, "[>=1000000]#,##0,\"K\";0"));
            Assert.Equal("-5", F(-5, "[<=100]0;[>100]0.0"));
            Assert.Equal("neg 5", F(-5, "[<0]\"neg \"0;\"pos \"0"));
            Assert.Equal("pos 5", F(5, "[<0]\"neg \"0;\"pos \"0"));
        }

        [Fact]
        public void TextValuesUseTheTextSectionOnly()
        {
            Assert.Equal("abc", T("abc", "0.00"));
            Assert.Equal("abc", T("abc", "@"));
            Assert.Equal("Ref. abc", T("abc", "\"Ref. \"@"));
            Assert.Equal("[text]", T("abc", "0;-0;0;\"[text]\""));
            Assert.Equal("  abc", T("  abc", "General"));
        }

        [Fact]
        public void NumberInTextFormatIsShownAsGeneral()
        {
            Assert.Equal("1234.5", F(1234.5, "@"));
        }

        [Theory]
        [InlineData(45366.75, "yyyy-mm-dd", "2024-03-15")]
        [InlineData(45366.75, "dd/mm/yyyy hh:mm", "15/03/2024 18:00")]
        [InlineData(45366.75, "d-mmm-yy", "15-Mar-24")]
        [InlineData(45366.75, "mmmm d, yyyy", "March 15, 2024")]
        [InlineData(45366.75, "dddd", "Friday")]
        [InlineData(45366.75, "ddd", "Fri")]
        [InlineData(45366.75, "mmmmm", "M")]
        [InlineData(45366.52, "h:mm AM/PM", "12:28 PM")]
        [InlineData(45366.25, "h:mm a/p", "6:00 a")]
        [InlineData(45366.52, "hh \"h\" mm", "12 h 28")]
        [InlineData(1.52, "[h]:mm", "36:28")]
        [InlineData(0.52, "[mm]:ss", "748:48")]
        [InlineData(0.000123456, "hh:mm:ss", "00:00:11")]   // arrondi à la seconde
        [InlineData(0.000123456, "hh:mm:ss.000", "00:00:10.667")]
        [InlineData(0.999999, "yyyy-mm-dd", "1900-01-00")]  // pas d'arrondi au jour suivant sans heure affichée
        [InlineData(0.999999, "hh:mm:ss", "00:00:00")]
        [InlineData(60, "yyyy-mm-dd", "1900-02-29")]        // 29 février 1900 fictif d'Excel
        [InlineData(61, "yyyy-mm-dd", "1900-03-01")]
        [InlineData(1, "dddd", "Sunday")]
        [InlineData(45366.5, "[$-40C]dddd d mmmm yyyy", "vendredi 15 mars 2024")]
        public void DateAndTimeFormats(double value, string code, string expected)
        {
            Assert.Equal(expected, F(value, code));
        }

        [Fact]
        public void NegativeDateIsUnrepresentable()
        {
            var settings = new ExcelFormatSettings(EnUs, EnUs, false, null);
            var result = settings.Get("dd/mm/yyyy").FormatNumber(-5, settings);
            Assert.True(result.Unrepresentable);
            Assert.Equal("-5", result.Text);
        }

        [Fact]
        public void Date1904System()
        {
            var settings = new ExcelFormatSettings(EnUs, EnUs, true, null);
            Assert.Equal("1904-01-01", settings.Get("yyyy-mm-dd").FormatNumber(0, settings).Text);
            Assert.Equal("2024-03-15", settings.Get("yyyy-mm-dd").FormatNumber(45366 - 1462, settings).Text);
        }

        [Fact]
        public void FrenchSeparatorsAndNames()
        {
            var nfi = FrFr.NumberFormat;
            Assert.Equal("1" + nfi.NumberGroupSeparator + "234,57", F(1234.5678, "#,##0.00", FrFr));
            Assert.Equal("1234,5678", F(1234.5678, "General", FrFr));
            Assert.Equal("1,23457E+11", F(123456789012, "General", FrFr));
            Assert.Equal("vendredi 15 mars 2024", F(45366.75, "dddd d mmmm yyyy", FrFr));
            Assert.Equal("1" + nfi.NumberGroupSeparator + "234,50 €", F(1234.5, "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-", FrFr));
            Assert.Equal("-3,00 €", F(-3, "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-", FrFr));
            Assert.Equal("1" + nfi.NumberGroupSeparator + "234,50 €", F(1234.5, "#,##0.00 [$€-40C]", FrFr));
        }

        [Fact]
        public void BuiltInFormatsFollowTheRegionalSettings()
        {
            Assert.Equal("15/03/2024", F(45366, ExcelNumberFormat.BuiltInCode(14, FrFr), FrFr));
            Assert.Equal("3/15/2024", F(45366, ExcelNumberFormat.BuiltInCode(14, EnUs), EnUs));
            Assert.Equal("10.00%", F(0.1, ExcelNumberFormat.BuiltInCode(10, EnUs)));
            Assert.Equal("[h]:mm:ss", ExcelNumberFormat.BuiltInCode(46, EnUs));
            Assert.Equal("General", ExcelNumberFormat.BuiltInCode(999, EnUs));
        }

        [Fact]
        public void LocalizedBooleansAndErrors()
        {
            Assert.Equal("VRAI", ExcelLocaleTexts.Boolean(true, FrFr));
            Assert.Equal("FALSE", ExcelLocaleTexts.Boolean(false, EnUs));
            Assert.Equal("#VALEUR!", ExcelLocaleTexts.Error("#VALUE!", FrFr));
            Assert.Equal("#DIV/0!", ExcelLocaleTexts.Error("#DIV/0!", EnUs));
            Assert.Equal("#SPILL!", ExcelLocaleTexts.Error("#SPILL!", FrFr));
        }

        [Fact]
        public void DecimalArithmeticIsExact()
        {
            Assert.Equal("0.3", ExcelDecimal.FromDouble(0.1 + 0.2).ToString());
            Assert.Equal("1.01", ExcelDecimal.FromDouble(1.005).RoundDecimals(2).ToString());
            Assert.Equal("1000", ExcelDecimal.FromDouble(999.96).RoundDecimals(1).ToString());
            Assert.Equal("0", ExcelDecimal.FromDouble(0.004).RoundDecimals(2).ToString());
            Assert.Equal("123456789012346", ExcelDecimal.FromDouble(123456789012345.6).ToString());
        }
    }
}
