using System;
using System.Globalization;
using System.Text;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Plage rectangulaire de cellules (lignes et colonnes en base 0, bornes incluses).</summary>
    public struct CellRange : IEquatable<CellRange>
    {
        public CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            FirstRow = Math.Min(firstRow, lastRow);
            LastRow = Math.Max(firstRow, lastRow);
            FirstColumn = Math.Min(firstColumn, lastColumn);
            LastColumn = Math.Max(firstColumn, lastColumn);
        }

        public readonly int FirstRow;
        public readonly int FirstColumn;
        public readonly int LastRow;
        public readonly int LastColumn;

        public int RowCount
        {
            get { return LastRow - FirstRow + 1; }
        }

        public int ColumnCount
        {
            get { return LastColumn - FirstColumn + 1; }
        }

        public bool Contains(int row, int column)
        {
            return row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
        }

        public bool Intersects(CellRange other)
        {
            return other.FirstRow <= LastRow && other.LastRow >= FirstRow && other.FirstColumn <= LastColumn && other.LastColumn >= FirstColumn;
        }

        /// <summary>« A1:C10 » (ou « B2 » pour une seule cellule).</summary>
        public override string ToString()
        {
            string first = CellReference.Format(FirstRow, FirstColumn);
            return FirstRow == LastRow && FirstColumn == LastColumn ? first : first + ":" + CellReference.Format(LastRow, LastColumn);
        }

        public bool Equals(CellRange other)
        {
            return FirstRow == other.FirstRow && FirstColumn == other.FirstColumn && LastRow == other.LastRow && LastColumn == other.LastColumn;
        }

        public override bool Equals(object obj)
        {
            return obj is CellRange && Equals((CellRange)obj);
        }

        public override int GetHashCode()
        {
            return ((FirstRow * 397 ^ FirstColumn) * 397 ^ LastRow) * 397 ^ LastColumn;
        }

        /// <summary>Intersection des deux plages (null si elles sont disjointes).</summary>
        public CellRange? Intersect(CellRange other)
        {
            if (!Intersects(other)) return null;
            return new CellRange(Math.Max(FirstRow, other.FirstRow), Math.Max(FirstColumn, other.FirstColumn),
                                 Math.Min(LastRow, other.LastRow), Math.Min(LastColumn, other.LastColumn));
        }

        /// <summary>
        /// Lit « A1:C10 », « $A$1:$C$10 », « B2 », « Feuil1!A1:C10 » (le nom de feuille est ignoré),
        /// ainsi que les lignes entières (« 1:3 ») et les colonnes entières (« A:F »).
        /// </summary>
        public static bool TryParse(string text, out CellRange range)
        {
            range = default(CellRange);
            if (string.IsNullOrEmpty(text)) return false;
            string s = text.Trim();
            int bang = s.LastIndexOf('!');
            if (bang >= 0) s = s.Substring(bang + 1);
            s = s.Replace("$", string.Empty).Trim();
            string[] parts = s.Split(':');
            if (parts.Length < 1 || parts.Length > 2) return false;
            int r1, c1, r2, c2;
            if (parts.Length == 2)
            {
                int a, b;
                if (TryParseRowNumber(parts[0], out a) && TryParseRowNumber(parts[1], out b))
                {
                    range = new CellRange(a, 0, b, CellReference.MaxColumns - 1);
                    return true;
                }
                if (TryParseColumnName(parts[0], out a) && TryParseColumnName(parts[1], out b))
                {
                    range = new CellRange(0, a, CellReference.MaxRows - 1, b);
                    return true;
                }
            }
            if (!CellReference.TryParse(parts[0], out r1, out c1)) return false;
            if (parts.Length == 2)
            {
                if (!CellReference.TryParse(parts[1], out r2, out c2)) return false;
            }
            else
            {
                r2 = r1;
                c2 = c1;
            }
            range = new CellRange(r1, c1, r2, c2);
            return true;
        }

        private static bool TryParseRowNumber(string s, out int row)
        {
            row = -1;
            int r;
            if (!int.TryParse(s.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out r) || r < 1 || r > CellReference.MaxRows) return false;
            row = r - 1;
            return true;
        }

        private static bool TryParseColumnName(string s, out int column)
        {
            column = -1;
            s = s.Trim();
            if (s.Length == 0 || s.Length > 3) return false;
            int col = 0;
            foreach (char ch in s)
            {
                char c = char.ToUpperInvariant(ch);
                if (c < 'A' || c > 'Z') return false;
                col = col * 26 + (c - 'A' + 1);
            }
            if (col > CellReference.MaxColumns) return false;
            column = col - 1;
            return true;
        }
    }

    /// <summary>Références de cellule Excel (« AB12 ») ↔ indices de ligne et de colonne en base 0.</summary>
    public static class CellReference
    {
        public const int MaxRows = 1048576;
        public const int MaxColumns = 16384;

        public static bool TryParse(string reference, out int row, out int column)
        {
            row = column = -1;
            if (string.IsNullOrEmpty(reference)) return false;
            int i = 0;
            int col = 0;
            while (i < reference.Length && char.IsLetter(reference[i]))
            {
                char c = char.ToUpperInvariant(reference[i]);
                if (c < 'A' || c > 'Z') return false;
                col = col * 26 + (c - 'A' + 1);
                if (col > MaxColumns) return false;
                i++;
            }
            if (i == 0 || i == reference.Length) return false;
            int r;
            if (!int.TryParse(reference.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out r)) return false;
            if (r < 1 || r > MaxRows) return false;
            row = r - 1;
            column = col - 1;
            return true;
        }

        public static string ColumnName(int column)
        {
            var sb = new StringBuilder();
            int n = column + 1;
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        public static string Format(int row, int column)
        {
            return ColumnName(column) + (row + 1).ToString(CultureInfo.InvariantCulture);
        }
    }
}
