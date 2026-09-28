using System;
using System.Collections.Generic;
using System.Linq;

namespace WordTableToExcel.Core.Layout
{
    /// <summary>Cellule telle que la voit le modèle objet de Word (Cell.RowIndex, Cell.ColumnIndex, Cell.Width).</summary>
    public sealed class ComCellInfo
    {
        public int RowIndex;     // base 1
        public int ColumnIndex;  // base 1
        public double WidthPt;
        public object Tag;
    }

    /// <summary>
    /// Mode de repli quand le XML du tableau n'est pas disponible (Word 2000/2002, document protégé…) :
    /// reconstruit la grille à partir de la largeur des cellules. Les fusions horizontales sont
    /// déduites des bords de colonnes communs à toutes les lignes ; les fusions verticales, des
    /// cellules absentes (Word masque les continuations de fusion verticale dans la collection Cells).
    /// </summary>
    public static class WidthGridBuilder
    {
        private const double Tolerance = 1.5; // points

        private sealed class Placed
        {
            public LayoutCell Cell;
            public double Left;
            public double Right;
        }

        public static TableLayout Build(IList<ComCellInfo> cells)
        {
            var layout = new TableLayout { HasCellFormatting = false };
            if (cells == null || cells.Count == 0) return layout;

            var byRow = cells.GroupBy(c => c.RowIndex).OrderBy(g => g.Key).ToList();
            int rowCount = byRow.Last().Key;
            layout.RowCount = rowCount;
            layout.RowHeightsPt = new double[rowCount];
            layout.RowHeightExact = new bool[rowCount];
            layout.SourceCellCounts = new int[rowCount];

            var placed = new List<Placed>();
            Dictionary<int, Placed> previous = new Dictionary<int, Placed>();

            for (int r = 1; r <= rowCount; r++)
            {
                var rowCells = cells.Where(c => c.RowIndex == r).OrderBy(c => c.ColumnIndex).ToDictionary(c => c.ColumnIndex);
                var current = new Dictionary<int, Placed>();
                int maxIndex = Math.Max(rowCells.Count == 0 ? 0 : rowCells.Keys.Max(), previous.Count == 0 ? 0 : previous.Keys.Max());
                double x = 0;
                int ordinal = 0;

                for (int k = 1; k <= maxIndex; k++)
                {
                    ComCellInfo info;
                    Placed above;
                    if (rowCells.TryGetValue(k, out info))
                    {
                        double width = info.WidthPt > 0 && info.WidthPt < 5000 ? info.WidthPt : 50;
                        var p = new Placed
                        {
                            Left = x,
                            Right = x + width,
                            Cell = new LayoutCell
                            {
                                Row = r - 1,
                                SourceRow = r - 1,
                                SourceIndex = k - 1,
                                SourceOrdinal = ordinal++,
                                SourceTag = info.Tag
                            }
                        };
                        placed.Add(p);
                        current[k] = p;
                        x = p.Right;
                    }
                    else if (previous.TryGetValue(k, out above))
                    {
                        // Continuation d'une fusion verticale.
                        above.Cell.RowSpan = (r - 1) - above.Cell.Row + 1;
                        current[k] = above;
                        x = above.Right;
                    }
                }
                layout.SourceCellCounts[r - 1] = maxIndex;
                previous = current;
            }

            // Bords de colonnes communs.
            var edges = new List<double>();
            foreach (var p in placed)
            {
                edges.Add(p.Left);
                edges.Add(p.Right);
            }
            edges.Sort();
            var boundaries = new List<double>();
            foreach (var e in edges)
            {
                if (boundaries.Count == 0 || e - boundaries[boundaries.Count - 1] > Tolerance) boundaries.Add(e);
            }

            layout.ColumnCount = Math.Max(1, boundaries.Count - 1);
            layout.ColumnWidthsPt = new double[layout.ColumnCount];
            for (int c = 0; c < layout.ColumnCount && c + 1 < boundaries.Count; c++)
            {
                layout.ColumnWidthsPt[c] = boundaries[c + 1] - boundaries[c];
            }

            foreach (var p in placed)
            {
                int start = Nearest(boundaries, p.Left);
                int end = Nearest(boundaries, p.Right);
                if (end <= start) end = start + 1;
                p.Cell.Column = Math.Min(start, layout.ColumnCount - 1);
                p.Cell.ColumnSpan = Math.Max(1, Math.Min(end, layout.ColumnCount) - p.Cell.Column);
                layout.Cells.Add(p.Cell);
            }
            return layout;
        }

        private static int Nearest(List<double> boundaries, double value)
        {
            int best = 0;
            double bestDistance = double.MaxValue;
            for (int i = 0; i < boundaries.Count; i++)
            {
                double d = Math.Abs(boundaries[i] - value);
                if (d < bestDistance)
                {
                    bestDistance = d;
                    best = i;
                }
            }
            return best;
        }
    }
}
