using System;
using System.Globalization;

namespace WordTableToExcel.Core.Model
{
    /// <summary>Couleur RVB opaque (sans canal alpha).</summary>
    public struct Rgb : IEquatable<Rgb>
    {
        public readonly byte R;
        public readonly byte G;
        public readonly byte B;

        public Rgb(int r, int g, int b)
        {
            R = Clamp(r);
            G = Clamp(g);
            B = Clamp(b);
        }

        public static readonly Rgb Black = new Rgb(0, 0, 0);
        public static readonly Rgb White = new Rgb(255, 255, 255);

        /// <summary>Lit une couleur hexadécimale « RRGGBB » (format OOXML). Renvoie null pour « auto » ou une valeur invalide.</summary>
        public static Rgb? FromHex(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return null;
            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) hex = hex.Substring(1);
            if (hex.Length == 8) hex = hex.Substring(2); // AARRGGBB
            if (hex.Length != 6) return null;
            int value;
            if (!int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value)) return null;
            return new Rgb((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
        }

        /// <summary>Convertit une valeur COM Office « BGR » (0x00BBGGRR).</summary>
        public static Rgb FromBgr(int bgr)
        {
            return new Rgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
        }

        /// <summary>Mélange deux couleurs : <paramref name="ratio"/> = part de <paramref name="other"/> (0..1).</summary>
        public Rgb Blend(Rgb other, double ratio)
        {
            if (ratio < 0) ratio = 0;
            if (ratio > 1) ratio = 1;
            return new Rgb(
                (int)Math.Round(R + (other.R - R) * ratio),
                (int)Math.Round(G + (other.G - G) * ratio),
                (int)Math.Round(B + (other.B - B) * ratio));
        }

        /// <summary>Forme « RRGGBB ».</summary>
        public string ToHex()
        {
            return R.ToString("X2", CultureInfo.InvariantCulture)
                 + G.ToString("X2", CultureInfo.InvariantCulture)
                 + B.ToString("X2", CultureInfo.InvariantCulture);
        }

        /// <summary>Forme « FFRRGGBB » attendue par SpreadsheetML.</summary>
        public string ToArgbHex()
        {
            return "FF" + ToHex();
        }

        public bool Equals(Rgb other)
        {
            return R == other.R && G == other.G && B == other.B;
        }

        public override bool Equals(object obj)
        {
            return obj is Rgb && Equals((Rgb)obj);
        }

        public override int GetHashCode()
        {
            return (R << 16) | (G << 8) | B;
        }

        public override string ToString()
        {
            return "#" + ToHex();
        }

        private static byte Clamp(int v)
        {
            return (byte)(v < 0 ? 0 : v > 255 ? 255 : v);
        }
    }
}
