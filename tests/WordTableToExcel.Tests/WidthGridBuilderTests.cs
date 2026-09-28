using System.Collections.Generic;
using System.Linq;
using WordTableToExcel.Core.Layout;
using Xunit;
using static WordTableToExcel.Tests.TestUtil;

namespace WordTableToExcel.Tests
{
    public class WidthGridBuilderTests
    {
        private static ComCellInfo C(int row, int col, double width)
        {
            return new ComCellInfo { RowIndex = row, ColumnIndex = col, WidthPt = width, Tag = row + ":" + col };
        }

        [Fact]
        public void RegularGrid()
        {
            var layout = WidthGridBuilder.Build(new List<ComCellInfo> { C(1, 1, 100), C(1, 2, 100), C(2, 1, 100), C(2, 2, 100) });
            Assert.Equal(2, layout.RowCount);
            Assert.Equal(2, layout.ColumnCount);
            Assert.All(layout.Cells, c => Assert.Equal(1, c.ColumnSpan));
            Assert.Equal(new[] { 100.0, 100.0 }, layout.ColumnWidthsPt);
        }

        [Fact]
        public void HorizontalMergeFromWidths()
        {
            var layout = WidthGridBuilder.Build(new List<ComCellInfo> { C(1, 1, 200), C(2, 1, 100), C(2, 2, 100.4) });
            Assert.Equal(2, layout.ColumnCount);
            Assert.Equal(2, At(layout, 0, 0).ColumnSpan);
        }

        [Fact]
        public void VerticalMergeFromMissingCell()
        {
            // Word ne liste pas la cellule (2,1), continuation de la fusion verticale de (1,1).
            var layout = WidthGridBuilder.Build(new List<ComCellInfo> { C(1, 1, 80), C(1, 2, 120), C(2, 2, 120) });
            Assert.Equal(2, layout.RowCount);
            var merged = At(layout, 0, 0);
            Assert.Equal(2, merged.RowSpan);
            Assert.Same(merged, At(layout, 1, 0));
            var cell = At(layout, 1, 1);
            Assert.Equal("2:2", cell.SourceTag);
            Assert.Equal(new[] { 80.0, 120.0 }, layout.ColumnWidthsPt);
        }

        [Fact]
        public void Empty()
        {
            var layout = WidthGridBuilder.Build(new List<ComCellInfo>());
            Assert.Empty(layout.Cells);
        }
    }
}
