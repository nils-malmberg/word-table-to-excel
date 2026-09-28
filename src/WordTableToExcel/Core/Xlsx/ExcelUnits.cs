using System;
using System.Text;

namespace WordTableToExcel.Core.Xlsx
{
    /// <summary>Conversions d'unités et de références Excel.</summary>
    public static class ExcelUnits
    {
        /// <summary>Hauteur de ligne maximale autorisée par Excel (points).</summary>
        public const double MaxRowHeight = 409.0;
        /// <summary>Hauteur de ligne par défaut (Calibri 11).</summary>
        public const double DefaultRowHeight = 15.0;
        /// <summary>Largeur maximale d'une colonne Excel (caractères).</summary>
        public const double MaxColumnWidth = 255.0;

        /// <summary>Largeur en points → largeur de colonne Excel (nombre de caractères « 0 » en Calibri 11).</summary>
        public static double PointsToColumnWidth(double points)
        {
            if (points <= 0) return 0;
            double pixels = points * 96.0 / 72.0;
            double width = Math.Truncate((pixels - 5) / 7.0 * 100 + 0.5) / 100;
            if (width < 1) width = 1;
            return Math.Min(width, MaxColumnWidth);
        }

        /// <summary>Largeur de colonne Excel → points (inverse approximatif).</summary>
        public static double ColumnWidthToPoints(double width)
        {
            return (width * 7.0 + 5) * 72.0 / 96.0;
        }

        /// <summary>Index de colonne base 0 → lettres (0 → A, 26 → AA).</summary>
        public static string ColumnName(int index)
        {
            var sb = new StringBuilder();
            int n = index + 1;
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        /// <summary>Référence de cellule (ligne et colonne base 0) → « B3 ».</summary>
        public static string CellReference(int row, int column)
        {
            return ColumnName(column) + (row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
