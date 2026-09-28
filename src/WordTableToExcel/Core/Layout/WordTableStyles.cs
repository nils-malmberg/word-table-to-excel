using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Layout
{
    /// <summary>Zones de mise en forme conditionnelle d'un style de tableau (ordre = priorité croissante).</summary>
    public enum TableRegion
    {
        WholeTable,
        Band1Vertical,
        Band2Vertical,
        Band1Horizontal,
        Band2Horizontal,
        FirstColumn,
        LastColumn,
        FirstRow,
        LastRow,
        NorthEastCell,
        NorthWestCell,
        SouthEastCell,
        SouthWestCell
    }

    /// <summary>Options « Ligne d'en-tête », « Lignes à bandes »… (élément tblLook).</summary>
    public struct TableLook
    {
        public bool FirstRow;
        public bool LastRow;
        public bool FirstColumn;
        public bool LastColumn;
        public bool NoHorizontalBand;
        public bool NoVerticalBand;

        /// <summary>Valeur par défaut de Word pour un nouveau tableau (04A0).</summary>
        public static TableLook Default
        {
            get { return FromHex(0x04A0); }
        }

        public static TableLook FromHex(int v)
        {
            return new TableLook
            {
                FirstRow = (v & 0x0020) != 0,
                LastRow = (v & 0x0040) != 0,
                FirstColumn = (v & 0x0080) != 0,
                LastColumn = (v & 0x0100) != 0,
                NoHorizontalBand = (v & 0x0200) != 0,
                NoVerticalBand = (v & 0x0400) != 0
            };
        }

        public static TableLook Parse(XElement tblLook)
        {
            if (tblLook == null) return Default;
            var look = Default;
            string hex = OoxmlXml.Attr(tblLook, "val");
            int v;
            if (hex != null && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v)) look = FromHex(v);

            bool b;
            if (OoxmlXml.TryBoolAttr(tblLook, "firstRow", out b)) look.FirstRow = b;
            if (OoxmlXml.TryBoolAttr(tblLook, "lastRow", out b)) look.LastRow = b;
            if (OoxmlXml.TryBoolAttr(tblLook, "firstColumn", out b)) look.FirstColumn = b;
            if (OoxmlXml.TryBoolAttr(tblLook, "lastColumn", out b)) look.LastColumn = b;
            if (OoxmlXml.TryBoolAttr(tblLook, "noHBand", out b)) look.NoHorizontalBand = b;
            if (OoxmlXml.TryBoolAttr(tblLook, "noVBand", out b)) look.NoVerticalBand = b;
            return look;
        }
    }

    /// <summary>Accès aux éléments WordprocessingML indépendamment de l'espace de noms (Word 2003 XML ou Open XML).</summary>
    public static class OoxmlXml
    {
        public static bool Is(XElement e, string localName)
        {
            return e != null && string.Equals(e.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase);
        }

        public static XElement Child(XElement parent, string localName)
        {
            if (parent == null) return null;
            foreach (var e in parent.Elements())
            {
                if (Is(e, localName)) return e;
            }
            return null;
        }

        public static XElement Child(XElement parent, params string[] path)
        {
            var current = parent;
            foreach (var name in path)
            {
                current = Child(current, name);
                if (current == null) return null;
            }
            return current;
        }

        public static string Attr(XElement e, string localName)
        {
            if (e == null) return null;
            foreach (var a in e.Attributes())
            {
                if (string.Equals(a.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) return a.Value;
            }
            return null;
        }

        public static int IntAttr(XElement e, string localName, int defaultValue)
        {
            string s = Attr(e, localName);
            int v;
            if (s != null && int.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) return v;
            double d;
            if (s != null && double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return (int)Math.Round(d);
            return defaultValue;
        }

        public static bool TryBoolAttr(XElement e, string localName, out bool value)
        {
            value = false;
            string s = Attr(e, localName);
            if (s == null) return false;
            value = ParseOnOff(s);
            return true;
        }

        /// <summary>Propriété « on/off » : présente sans valeur = vrai.</summary>
        public static bool IsOn(XElement element)
        {
            if (element == null) return false;
            string v = Attr(element, "val");
            return v == null || ParseOnOff(v);
        }

        public static bool ParseOnOff(string s)
        {
            s = s.Trim().ToLowerInvariant();
            return !(s == "0" || s == "false" || s == "off" || s == "none");
        }
    }

    /// <summary>Résultat de lecture d'une propriété : « non définie », « définie à aucune », ou valeur.</summary>
    internal struct Defined<T>
    {
        public bool IsDefined;
        public T Value;

        public static Defined<T> Undefined
        {
            get { return new Defined<T>(); }
        }

        public static Defined<T> Of(T value)
        {
            return new Defined<T> { IsDefined = true, Value = value };
        }
    }

    /// <summary>
    /// Styles de tableau du document : permet de résoudre fond, bordures et alignement vertical
    /// d'une cellule en tenant compte de la mise en forme directe, des zones conditionnelles
    /// (ligne d'en-tête, bandes, première colonne…) et de l'héritage (basedOn).
    /// </summary>
    internal sealed class TableStyleSheet
    {
        private readonly Dictionary<string, XElement> _styles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        private readonly XElement _defaultTableStyle;

        public TableStyleSheet(XContainer root)
        {
            foreach (var style in root.Descendants().Where(e => OoxmlXml.Is(e, "style")))
            {
                if (!string.Equals(OoxmlXml.Attr(style, "type"), "table", StringComparison.OrdinalIgnoreCase)) continue;
                string id = OoxmlXml.Attr(style, "styleId");
                if (string.IsNullOrEmpty(id) || _styles.ContainsKey(id)) continue;
                _styles[id] = style;
                string isDefault = OoxmlXml.Attr(style, "default");
                if (isDefault != null && OoxmlXml.ParseOnOff(isDefault) && _defaultTableStyle == null) _defaultTableStyle = style;
            }
        }

        /// <summary>Chaîne d'héritage du style (le style lui-même d'abord).</summary>
        public List<XElement> Chain(string styleId)
        {
            var chain = new List<XElement>();
            XElement style = null;
            if (!string.IsNullOrEmpty(styleId)) _styles.TryGetValue(styleId, out style);
            if (style == null) style = _defaultTableStyle;
            var seen = new HashSet<XElement>();
            while (style != null && seen.Add(style) && chain.Count < 20)
            {
                chain.Add(style);
                string basedOn = OoxmlXml.Attr(OoxmlXml.Child(style, "basedOn"), "val");
                XElement parent = null;
                if (!string.IsNullOrEmpty(basedOn)) _styles.TryGetValue(basedOn, out parent);
                style = parent;
            }
            return chain;
        }

        /// <summary>Taille des bandes horizontales / verticales (tblStyleRowBandSize / tblStyleColBandSize).</summary>
        public static int BandSize(List<XElement> chain, string elementName)
        {
            foreach (var style in chain)
            {
                var e = OoxmlXml.Child(style, "tblPr", elementName);
                if (e != null)
                {
                    int v = OoxmlXml.IntAttr(e, "val", 1);
                    return v < 1 ? 1 : v;
                }
            }
            return 1;
        }

        /// <summary>Propriétés conditionnelles (tblStylePr) d'une zone, du style le plus dérivé au plus général.</summary>
        public static IEnumerable<XElement> ConditionalFormats(List<XElement> chain, TableRegion region)
        {
            string type = RegionTypeName(region);
            foreach (var style in chain)
            {
                foreach (var pr in style.Elements().Where(e => OoxmlXml.Is(e, "tblStylePr")))
                {
                    if (string.Equals(OoxmlXml.Attr(pr, "type"), type, StringComparison.OrdinalIgnoreCase)) yield return pr;
                }
            }
        }

        public static string RegionTypeName(TableRegion region)
        {
            switch (region)
            {
                case TableRegion.Band1Vertical: return "band1Vert";
                case TableRegion.Band2Vertical: return "band2Vert";
                case TableRegion.Band1Horizontal: return "band1Horz";
                case TableRegion.Band2Horizontal: return "band2Horz";
                case TableRegion.FirstColumn: return "firstCol";
                case TableRegion.LastColumn: return "lastCol";
                case TableRegion.FirstRow: return "firstRow";
                case TableRegion.LastRow: return "lastRow";
                case TableRegion.NorthEastCell: return "neCell";
                case TableRegion.NorthWestCell: return "nwCell";
                case TableRegion.SouthEastCell: return "seCell";
                case TableRegion.SouthWestCell: return "swCell";
                default: return "wholeTable";
            }
        }
    }

    /// <summary>Conversion des propriétés WordprocessingML (trame, bordures) vers le modèle.</summary>
    internal static class OoxmlFormat
    {
        /// <summary>Lit une trame w:shd. « Définie à null » = explicitement sans fond.</summary>
        public static Defined<Rgb?> ReadShading(XElement shd)
        {
            if (shd == null) return Defined<Rgb?>.Undefined;
            string pattern = (OoxmlXml.Attr(shd, "val") ?? "clear").Trim().ToLowerInvariant().Replace("-", string.Empty);
            Rgb? fill = Rgb.FromHex(OoxmlXml.Attr(shd, "fill"));
            Rgb? color = Rgb.FromHex(OoxmlXml.Attr(shd, "color"));

            if (pattern == "nil") return Defined<Rgb?>.Of(null);
            if (pattern == "clear") return Defined<Rgb?>.Of(fill);
            if (pattern == "solid") return Defined<Rgb?>.Of(color ?? Rgb.Black);

            if (pattern.StartsWith("pct", StringComparison.Ordinal))
            {
                int pct;
                if (int.TryParse(pattern.Substring(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out pct))
                {
                    if (fill == null && color == null && pct == 0) return Defined<Rgb?>.Of(null);
                    Rgb background = fill ?? Rgb.White;
                    Rgb foreground = color ?? Rgb.Black;
                    return Defined<Rgb?>.Of(background.Blend(foreground, pct / 100.0));
                }
            }
            // Motifs (rayures, quadrillage…) : Excel ne les reproduit pas fidèlement, on garde le fond.
            if (fill.HasValue) return Defined<Rgb?>.Of(fill);
            return Defined<Rgb?>.Of(color.HasValue ? Rgb.White.Blend(color.Value, 0.25) : (Rgb?)null);
        }

        /// <summary>Lit un trait de bordure (w:top, w:left…). nil/none = explicitement sans bordure.</summary>
        public static Defined<BorderLine> ReadBorder(XElement border)
        {
            if (border == null) return Defined<BorderLine>.Undefined;
            string val = (OoxmlXml.Attr(border, "val") ?? "single").Trim().ToLowerInvariant().Replace("-", string.Empty);
            if (val == "nil" || val == "none" || val.Length == 0) return Defined<BorderLine>.Of(BorderLine.None);

            int eighths = OoxmlXml.IntAttr(border, "sz", 4);
            Rgb? color = Rgb.FromHex(OoxmlXml.Attr(border, "color"));
            return Defined<BorderLine>.Of(new BorderLine(MapBorderStyle(val, eighths), color));
        }

        /// <summary>Correspondance style de trait Word → Excel (épaisseur en 1/8 de point).</summary>
        public static BorderStyle MapBorderStyle(string val, int eighths)
        {
            bool heavy = eighths >= 12;
            switch (val)
            {
                case "single":
                case "wave":
                case "doublewave":
                case "dashdotstroked":
                case "threedemboss":
                case "threedengrave":
                case "outset":
                case "inset":
                    return BySize(eighths);
                case "thick":
                    return eighths >= 18 ? BorderStyle.Thick : BorderStyle.Medium;
                case "double":
                case "triple":
                    return BorderStyle.Double;
                case "dotted":
                    return BorderStyle.Dotted;
                case "dashed":
                case "dashsmallgap":
                    return heavy ? BorderStyle.MediumDashed : BorderStyle.Dashed;
                case "dotdash":
                    return heavy ? BorderStyle.MediumDashDot : BorderStyle.DashDot;
                case "dotdotdash":
                    return heavy ? BorderStyle.MediumDashDotDot : BorderStyle.DashDotDot;
            }
            if (val.Contains("thin") || val.Contains("thick") || val.Contains("gap")) return BorderStyle.Double;
            return BySize(eighths);
        }

        private static BorderStyle BySize(int eighths)
        {
            if (eighths <= 8) return BorderStyle.Thin;    // ≤ 1 pt
            if (eighths <= 16) return BorderStyle.Medium; // ≤ 2 pt
            return BorderStyle.Thick;
        }

        /// <summary>Recherche un bord dans un conteneur de bordures (tcBorders / tblBorders).</summary>
        public static Defined<BorderLine> ReadEdge(XElement borders, string edge)
        {
            if (borders == null) return Defined<BorderLine>.Undefined;
            var e = OoxmlXml.Child(borders, edge);
            if (e == null)
            {
                // Open XML « strict » / Word 2010+ : start / end au lieu de left / right.
                if (edge == "left") e = OoxmlXml.Child(borders, "start");
                else if (edge == "right") e = OoxmlXml.Child(borders, "end");
            }
            return ReadBorder(e);
        }

        public static Defined<VerticalAlignment> ReadVerticalAlignment(XElement tcPr)
        {
            var v = OoxmlXml.Child(tcPr, "vAlign");
            if (v == null) return Defined<VerticalAlignment>.Undefined;
            string val = (OoxmlXml.Attr(v, "val") ?? "top").ToLowerInvariant();
            if (val == "center") return Defined<VerticalAlignment>.Of(VerticalAlignment.Center);
            if (val == "bottom") return Defined<VerticalAlignment>.Of(VerticalAlignment.Bottom);
            return Defined<VerticalAlignment>.Of(VerticalAlignment.Top);
        }

        /// <summary>Orientation du texte (textDirection / textFlow) → rotation Excel.</summary>
        public static int ReadTextRotation(XElement tcPr)
        {
            var e = OoxmlXml.Child(tcPr, "textDirection") ?? OoxmlXml.Child(tcPr, "textFlow");
            if (e == null) return 0;
            string val = (OoxmlXml.Attr(e, "val") ?? string.Empty).ToLowerInvariant().Replace("-", string.Empty);
            switch (val)
            {
                case "btlr":
                case "btl":
                    return 90;   // de bas en haut
                case "tbrl":
                case "tbrlv":
                case "tbv":
                    return 180;  // de haut en bas
                default:
                    return 0;
            }
        }
    }
}
