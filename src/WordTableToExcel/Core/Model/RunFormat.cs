using System;
using System.Globalization;
using System.Text;

namespace WordTableToExcel.Core.Model
{
    public enum UnderlineKind
    {
        None,
        Single,
        Double
    }

    public enum VerticalPosition
    {
        Baseline,
        Superscript,
        Subscript
    }

    /// <summary>Mise en forme des caractères d'un segment de texte (police, taille, styles, couleur).</summary>
    public sealed class RunFormat : IEquatable<RunFormat>
    {
        public string FontName;
        public double? Size;
        public bool Bold;
        public bool Italic;
        public bool Strike;
        public UnderlineKind Underline;
        public VerticalPosition Position;
        /// <summary>Couleur du texte ; null = automatique.</summary>
        public Rgb? Color;

        public RunFormat Clone()
        {
            return (RunFormat)MemberwiseClone();
        }

        public bool Equals(RunFormat other)
        {
            if (ReferenceEquals(other, null)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(FontName, other.FontName, StringComparison.Ordinal)
                && Nullable.Equals(Size, other.Size)
                && Bold == other.Bold
                && Italic == other.Italic
                && Strike == other.Strike
                && Underline == other.Underline
                && Position == other.Position
                && Nullable.Equals(Color, other.Color);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as RunFormat);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int h = FontName == null ? 0 : StringComparer.Ordinal.GetHashCode(FontName);
                h = h * 31 + Size.GetHashCode();
                h = h * 31 + (Bold ? 1 : 0);
                h = h * 31 + (Italic ? 1 : 0);
                h = h * 31 + (Strike ? 1 : 0);
                h = h * 31 + (int)Underline;
                h = h * 31 + (int)Position;
                h = h * 31 + Color.GetHashCode();
                return h;
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(FontName ?? "(défaut)");
            if (Size.HasValue) sb.Append(' ').Append(Size.Value.ToString(CultureInfo.InvariantCulture)).Append("pt");
            if (Bold) sb.Append(" gras");
            if (Italic) sb.Append(" italique");
            if (Strike) sb.Append(" barré");
            if (Underline != UnderlineKind.None) sb.Append(" souligné");
            if (Position != VerticalPosition.Baseline) sb.Append(' ').Append(Position);
            if (Color.HasValue) sb.Append(' ').Append(Color.Value);
            return sb.ToString();
        }
    }

    /// <summary>Segment de texte homogène d'une cellule.</summary>
    public sealed class TextRun
    {
        public TextRun(string text, RunFormat format, Rgb? highlight)
        {
            Text = text ?? string.Empty;
            Format = format ?? new RunFormat();
            Highlight = highlight;
        }

        public string Text;
        public RunFormat Format;
        /// <summary>Surlignage (ou trame de caractères) ; Excel ne sait pas l'appliquer à une partie de cellule.</summary>
        public Rgb? Highlight;

        public bool HasSameFormatting(TextRun other)
        {
            return other != null && Format.Equals(other.Format) && Nullable.Equals(Highlight, other.Highlight);
        }

        public override string ToString()
        {
            return "\"" + Text + "\" [" + Format + "]";
        }
    }
}
