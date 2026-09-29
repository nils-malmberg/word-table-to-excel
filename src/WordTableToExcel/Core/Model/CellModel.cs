using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableToExcel.Core.Model
{
    public enum HorizontalAlignment
    {
        General,
        Left,
        Center,
        Right,
        Justify,
        Distributed
    }

    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    }

    public enum BorderStyle
    {
        None,
        Hair,
        Thin,
        Medium,
        Thick,
        Double,
        Dotted,
        Dashed,
        MediumDashed,
        DashDot,
        MediumDashDot,
        DashDotDot,
        MediumDashDotDot
    }

    /// <summary>Trait de bordure (style + couleur ; couleur null = automatique).</summary>
    public sealed class BorderLine : IEquatable<BorderLine>
    {
        public static readonly BorderLine None = new BorderLine(BorderStyle.None, null);

        public BorderLine(BorderStyle style, Rgb? color)
        {
            Style = style;
            Color = style == BorderStyle.None ? null : color;
        }

        public readonly BorderStyle Style;
        public readonly Rgb? Color;

        public bool IsVisible
        {
            get { return Style != BorderStyle.None; }
        }

        public bool Equals(BorderLine other)
        {
            return !ReferenceEquals(other, null) && Style == other.Style && Nullable.Equals(Color, other.Color);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as BorderLine);
        }

        public override int GetHashCode()
        {
            return ((int)Style * 397) ^ Color.GetHashCode();
        }

        public override string ToString()
        {
            return Style + (Color.HasValue ? " " + Color.Value : string.Empty);
        }
    }

    public sealed class CellBorders
    {
        public BorderLine Left = BorderLine.None;
        public BorderLine Right = BorderLine.None;
        public BorderLine Top = BorderLine.None;
        public BorderLine Bottom = BorderLine.None;

        public CellBorders Clone()
        {
            return (CellBorders)MemberwiseClone();
        }
    }

    /// <summary>
    /// Cellule d'un tableau, positionnée dans une grille régulière (lignes × colonnes).
    /// Une cellule fusionnée occupe <see cref="RowSpan"/> × <see cref="ColumnSpan"/> positions.
    /// </summary>
    public sealed class CellModel
    {
        public int Row;
        public int Column;
        public int RowSpan = 1;
        public int ColumnSpan = 1;

        public readonly List<TextRun> Runs = new List<TextRun>();

        /// <summary>Couleur de fond ; null = aucune.</summary>
        public Rgb? Fill;
        public HorizontalAlignment HorizontalAlignment = HorizontalAlignment.General;
        public VerticalAlignment VerticalAlignment = VerticalAlignment.Top;
        /// <summary>Rotation au sens Excel : 0, 90 (bas → haut) ou 180 (haut → bas).</summary>
        public int TextRotation;
        /// <summary>Retrait du texte en niveaux Excel (1 niveau ≈ largeur de 3 caractères) ; 0 = aucun.</summary>
        public int IndentLevel;
        public CellBorders Borders = new CellBorders();

        public bool IsMerged
        {
            get { return RowSpan > 1 || ColumnSpan > 1; }
        }

        public string PlainText
        {
            get
            {
                if (Runs.Count == 1) return Runs[0].Text;
                var sb = new StringBuilder();
                foreach (var run in Runs) sb.Append(run.Text);
                return sb.ToString();
            }
        }

        /// <summary>Mise en forme « principale » de la cellule (celle du premier segment non vide).</summary>
        public RunFormat PrimaryFormat
        {
            get
            {
                foreach (var run in Runs)
                {
                    if (run.Text.Trim().Length > 0) return run.Format;
                }
                return Runs.Count > 0 ? Runs[0].Format : null;
            }
        }

        public override string ToString()
        {
            return string.Format("[{0},{1} {2}x{3}] {4}", Row, Column, RowSpan, ColumnSpan, PlainText);
        }
    }
}
