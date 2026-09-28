using System;
using System.IO;
using System.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Tests
{
    internal static class TestUtil
    {
        public static string Fixture(string name)
        {
            return File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "fixtures", name));
        }

        public static LayoutCell At(TableLayout layout, int row, int column)
        {
            return layout.Cells.Single(c => c.Row <= row && row <= c.LastRow && c.Column <= column && column <= c.LastColumn);
        }

        public static CellModel At(TableModel table, int row, int column)
        {
            return table.Cells.Single(c => c.Row <= row && row < c.Row + c.RowSpan && c.Column <= column && column < c.Column + c.ColumnSpan);
        }

        public static Rgb Hex(string hex)
        {
            return Rgb.FromHex(hex).Value;
        }
    }
}
