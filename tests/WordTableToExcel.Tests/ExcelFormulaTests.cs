using System.Collections.Generic;
using WordTableToExcel.Core.ExcelImport;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class ExcelFormulaTests
    {
        private sealed class Grid : ExcelFormula.IContext
        {
            public readonly Dictionary<string, object> Values = new Dictionary<string, object>();

            public object Value(int row, int column)
            {
                object v;
                return Values.TryGetValue(CellReference.Format(row, column), out v) ? v : null;
            }

            public string SheetName
            {
                get { return "F"; }
            }
        }

        private static object Eval(string formula, Grid grid, string at = "A1", string anchor = "A1")
        {
            int r, c, ar, ac;
            CellReference.TryParse(at, out r, out c);
            CellReference.TryParse(anchor, out ar, out ac);
            return ExcelFormula.Parse(formula).Evaluate(grid, r, c, ar, ac);
        }

        [Fact]
        public void ArithmeticComparisonAndPrecedence()
        {
            var g = new Grid();
            Assert.Equal(7.0, Eval("1+2*3", g));
            Assert.Equal(4.0, Eval("-2^2", g));
            Assert.Equal(0.5, Eval("50%", g));
            Assert.Equal(true, Eval("2>=2", g));
            Assert.Equal(true, Eval("\"abc\"=\"ABC\"", g));
            Assert.Equal("a1", Eval("\"a\"&1", g));
            Assert.IsType<ExcelErrorValue>(Eval("1/0", g));
        }

        [Fact]
        public void RelativeAndAbsoluteReferences()
        {
            var g = new Grid();
            g.Values["C2"] = 150.0;
            g.Values["C3"] = 50.0;
            g.Values["B1"] = 100.0;
            // Règle ancrée en A2 : « $C2>$B$1 » évaluée en A3 regarde C3.
            Assert.Equal(true, Eval("$C2>$B$1", g, "A2", "A2"));
            Assert.Equal(false, Eval("$C2>$B$1", g, "A3", "A2"));
            Assert.Equal(200.0, Eval("SUM(C2:C3)", g));
            Assert.Equal(100.0, Eval("AVERAGE(C2:C3)", g));
            Assert.Equal(2.0, Eval("COUNT(C1:C5)", g));
        }

        [Fact]
        public void FunctionsUsedInConditionalFormatting()
        {
            var g = new Grid();
            g.Values["A1"] = "Terminé";
            Assert.Equal(true, Eval("MOD(ROW(),2)=0", g, "B4"));
            Assert.Equal(false, Eval("MOD(ROW(),2)=0", g, "B5"));
            Assert.Equal(true, Eval("AND(COLUMN()>1,ISBLANK(B1))", g, "C1"));
            Assert.Equal(true, Eval("ISNUMBER(SEARCH(\"term\",A1))", g));
            Assert.Equal(true, Eval("IF(LEN(A1)>3,TRUE,FALSE)", g));
            Assert.Equal(true, Eval("_xlfn.ISODD(3)", g));
            Assert.Equal(1.0, Eval("COUNTIF(A1:A3,\"terminé\")", g));
        }

        [Fact]
        public void UnsupportedSyntaxIsReported()
        {
            Assert.Throws<UnsupportedFormulaException>(() => ExcelFormula.Parse("Feuil2!A1>0"));
            Assert.Throws<UnsupportedFormulaException>(() => ExcelFormula.Parse("MonNom>0"));
            Assert.Throws<UnsupportedFormulaException>(() => ExcelFormula.Parse("(1+2"));
            var g = new Grid();
            Assert.Throws<UnsupportedFormulaException>(() => Eval("VLOOKUP(1,A1:B2,2,FALSE)", g));
        }
    }
}
