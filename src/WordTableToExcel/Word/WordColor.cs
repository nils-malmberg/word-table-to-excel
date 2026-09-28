using System;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Décodage des couleurs Word (<c>WdColor</c>, <c>WdColorIndex</c>) : couleurs RVB, « Automatique »,
    /// couleurs de thème (Word 2007+) et couleurs de surlignage.
    /// </summary>
    public static class WordColor
    {
        /// <summary>wdColorAutomatic.</summary>
        public const int Automatic = -16777216;

        private static readonly Rgb?[] Highlights =
        {
            null,                       // 0 wdNoHighlight
            new Rgb(0x00, 0x00, 0x00),  // 1 wdBlack
            new Rgb(0x00, 0x00, 0xFF),  // 2 wdBlue
            new Rgb(0x00, 0xFF, 0xFF),  // 3 wdTurquoise
            new Rgb(0x00, 0xFF, 0x00),  // 4 wdBrightGreen
            new Rgb(0xFF, 0x00, 0xFF),  // 5 wdPink
            new Rgb(0xFF, 0x00, 0x00),  // 6 wdRed
            new Rgb(0xFF, 0xFF, 0x00),  // 7 wdYellow
            new Rgb(0xFF, 0xFF, 0xFF),  // 8 wdWhite
            new Rgb(0x00, 0x00, 0x80),  // 9 wdDarkBlue
            new Rgb(0x00, 0x80, 0x80),  // 10 wdTeal
            new Rgb(0x00, 0x80, 0x00),  // 11 wdGreen
            new Rgb(0x80, 0x00, 0x80),  // 12 wdViolet
            new Rgb(0x80, 0x00, 0x00),  // 13 wdDarkRed
            new Rgb(0x80, 0x80, 0x00),  // 14 wdDarkYellow
            new Rgb(0x80, 0x80, 0x80),  // 15 wdGray50
            new Rgb(0xC0, 0xC0, 0xC0)   // 16 wdGray25
        };

        /// <summary>Couleur de surlignage (<c>Range.HighlightColorIndex</c>) ; null = aucun.</summary>
        public static Rgb? FromHighlightIndex(int index)
        {
            if (index <= 0 || index >= Highlights.Length) return null;
            return Highlights[index];
        }

        /// <summary>Couleur RVB « simple » (octet de poids fort nul).</summary>
        public static bool IsPlainRgb(int value)
        {
            return (value & unchecked((int)0xFF000000)) == 0;
        }

        /// <summary>Couleur de thème Word 2007+ (octet de poids fort 0xD0 à 0xDF).</summary>
        public static bool IsThemeColor(int value)
        {
            return ((uint)value >> 28) == 0xD;
        }

        /// <summary>
        /// Décode une valeur WdColor. Les couleurs de thème sont résolues grâce à
        /// <paramref name="themeLookup"/> (index MsoThemeColorSchemeIndex → couleur), si fourni.
        /// </summary>
        public static Rgb? Decode(int value, Func<int, Rgb?> themeLookup)
        {
            if (value == Automatic || value == WordCom.Undefined) return null;
            if (IsPlainRgb(value)) return Rgb.FromBgr(value);
            if (!IsThemeColor(value) || themeLookup == null) return null;

            int wdThemeIndex = (value >> 24) & 0x0F;
            Rgb? baseColor = themeLookup(ToSchemeIndex(wdThemeIndex));
            if (!baseColor.HasValue) return null;

            // Octets de poids faible : éclaircissement (tint) et assombrissement (shade), 0xFF = aucun.
            int tint = (value >> 8) & 0xFF;
            int shade = value & 0xFF;
            Rgb color = baseColor.Value;
            if (tint > 0 && tint < 0xFF) color = color.Blend(Rgb.White, 1 - tint / 255.0);
            if (shade > 0 && shade < 0xFF) color = Rgb.Black.Blend(color, shade / 255.0);
            return color;
        }

        /// <summary>WdThemeColorIndex → MsoThemeColorSchemeIndex (index de ThemeColorScheme.Colors).</summary>
        public static int ToSchemeIndex(int wdThemeColorIndex)
        {
            switch (wdThemeColorIndex)
            {
                case 12: return 2; // Arrière-plan 1 → Clair 1
                case 13: return 1; // Texte 1 → Foncé 1
                case 14: return 4; // Arrière-plan 2 → Clair 2
                case 15: return 3; // Texte 2 → Foncé 2
                default: return wdThemeColorIndex + 1;
            }
        }
    }
}
