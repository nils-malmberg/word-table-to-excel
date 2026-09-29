using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Couleurs Excel : RVB direct (« FFRRGGBB »), palette indexée (64 couleurs, personnalisable),
    /// couleurs de thème avec nuance (tint), automatique.
    /// </summary>
    public sealed class ExcelColors
    {
        private static readonly string[] DefaultPalette =
        {
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
            "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
            "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
            "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
            "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
            "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
        };

        /// <summary>Thème Office par défaut (2013+), utilisé si le classeur n'a pas de thème.</summary>
        private static readonly string[] DefaultTheme =
        {
            "000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"
        };

        private readonly Rgb[] _palette;
        private readonly Rgb[] _theme; // ordre du XML : dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink

        public ExcelColors(IList<Rgb> customPalette, IList<Rgb> theme)
        {
            _palette = DefaultPalette.Select(h => Rgb.FromHex(h).Value).ToArray();
            if (customPalette != null)
            {
                for (int i = 0; i < customPalette.Count && i < _palette.Length; i++) _palette[i] = customPalette[i];
            }
            _theme = DefaultTheme.Select(h => Rgb.FromHex(h).Value).ToArray();
            if (theme != null)
            {
                for (int i = 0; i < theme.Count && i < _theme.Length; i++) _theme[i] = theme[i];
            }
        }

        public static ExcelColors Default
        {
            get { return new ExcelColors(null, null); }
        }

        /// <summary>Couleur de palette (index 0 à 63) ; 64 = texte système (noir), 65 = fond système (blanc).</summary>
        public Rgb? Indexed(int index)
        {
            if (index >= 0 && index < _palette.Length) return _palette[index];
            if (index == 64) return null;          // couleur système « automatique »
            if (index == 65) return Rgb.White;
            return null;
        }

        /// <summary>
        /// Couleur de thème. Excel numérote 0 = Fond 1 (lt1), 1 = Texte 1 (dk1), 2 = Fond 2 (lt2), 3 = Texte 2 (dk2),
        /// puis Accent 1 à 6, lien hypertexte, lien visité \u2014 alors que le XML du thème commence par dk1.
        /// </summary>
        public Rgb? Theme(int index)
        {
            int xmlIndex = index == 0 ? 1 : index == 1 ? 0 : index == 2 ? 3 : index == 3 ? 2 : index;
            return xmlIndex >= 0 && xmlIndex < _theme.Length ? _theme[xmlIndex] : (Rgb?)null;
        }

        /// <summary>Accent 1 à 6 (1 = accent1).</summary>
        public Rgb Accent(int number)
        {
            int i = 3 + Math.Max(1, Math.Min(6, number));
            return _theme[i];
        }

        public Rgb Text1
        {
            get { return _theme[0]; }
        }

        /// <summary>Lit un élément couleur SpreadsheetML (color, fgColor, bgColor…). null = automatique / absent.</summary>
        public Rgb? Resolve(XElement color)
        {
            if (color == null) return null;
            string auto = OoxmlXml.Attr(color, "auto");
            if (auto != null && OoxmlXml.ParseOnOff(auto)) return null;

            Rgb? baseColor = null;
            string rgb = OoxmlXml.Attr(color, "rgb");
            string theme = OoxmlXml.Attr(color, "theme");
            string indexed = OoxmlXml.Attr(color, "indexed");
            if (rgb != null) baseColor = Rgb.FromHex(rgb);
            else if (theme != null) baseColor = Theme(OoxmlXml.IntAttr(color, "theme", -1));
            else if (indexed != null) baseColor = Indexed(OoxmlXml.IntAttr(color, "indexed", -1));
            if (!baseColor.HasValue) return null;

            string tintText = OoxmlXml.Attr(color, "tint");
            double tint;
            if (tintText != null && double.TryParse(tintText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out tint) && tint != 0)
            {
                return ApplyTint(baseColor.Value, tint);
            }
            return baseColor;
        }

        /// <summary>Nuance Excel : éclaircit (tint &gt; 0) ou assombrit (tint &lt; 0) la luminance TSL.</summary>
        public static Rgb ApplyTint(Rgb color, double tint)
        {
            double h, s, l;
            ToHsl(color, out h, out s, out l);
            if (tint < 0) l = l * (1 + tint);
            else l = l * (1 - tint) + tint;
            return FromHsl(h, s, Math.Max(0, Math.Min(1, l)));
        }

        /// <summary>Lit les couleurs du thème (xl/theme/theme1.xml), dans l'ordre dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink.</summary>
        public static List<Rgb> ReadTheme(XDocument theme)
        {
            var result = new List<Rgb>();
            if (theme == null || theme.Root == null) return result;
            var scheme = theme.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
            if (scheme == null) return result;
            foreach (var name in new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" })
            {
                var slot = scheme.Elements().FirstOrDefault(e => e.Name.LocalName == name);
                Rgb? value = null;
                if (slot != null)
                {
                    var srgb = slot.Elements().FirstOrDefault(e => e.Name.LocalName == "srgbClr");
                    var sys = slot.Elements().FirstOrDefault(e => e.Name.LocalName == "sysClr");
                    if (srgb != null) value = Rgb.FromHex(OoxmlXml.Attr(srgb, "val"));
                    else if (sys != null) value = Rgb.FromHex(OoxmlXml.Attr(sys, "lastClr"));
                }
                if (!value.HasValue) break;
                result.Add(value.Value);
            }
            return result;
        }

        private static void ToHsl(Rgb c, out double h, out double s, out double l)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            l = (max + min) / 2;
            if (max == min)
            {
                h = s = 0;
                return;
            }
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (max == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }

        private static Rgb FromHsl(double h, double s, double l)
        {
            if (s == 0)
            {
                int v = (int)Math.Round(l * 255);
                return new Rgb(v, v, v);
            }
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;
            return new Rgb((int)Math.Round(HueToRgb(p, q, h + 1.0 / 3) * 255),
                           (int)Math.Round(HueToRgb(p, q, h) * 255),
                           (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3) * 255));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }
    }
}
