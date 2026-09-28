using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Xlsx
{
    /// <summary>
    /// Table des styles du classeur (polices, remplissages, bordures, formats numériques, formats de cellule),
    /// dédupliquée pour rester compacte quel que soit le nombre de cellules.
    /// </summary>
    internal sealed class XlsxStyleRegistry
    {
        public const string DefaultFontName = "Calibri";
        public const double DefaultFontSize = 11;

        private struct XfKey : IEquatable<XfKey>
        {
            public int Font, Fill, Border, NumFmt, Rotation;
            public HorizontalAlignment Horizontal;
            public VerticalAlignment Vertical;
            public bool Wrap;

            public bool Equals(XfKey o)
            {
                return Font == o.Font && Fill == o.Fill && Border == o.Border && NumFmt == o.NumFmt
                    && Rotation == o.Rotation && Horizontal == o.Horizontal && Vertical == o.Vertical && Wrap == o.Wrap;
            }

            public override bool Equals(object obj)
            {
                return obj is XfKey && Equals((XfKey)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int h = Font;
                    h = h * 31 + Fill;
                    h = h * 31 + Border;
                    h = h * 31 + NumFmt;
                    h = h * 31 + Rotation;
                    h = h * 31 + (int)Horizontal;
                    h = h * 31 + (int)Vertical;
                    return h * 31 + (Wrap ? 1 : 0);
                }
            }
        }

        private sealed class BorderKey : IEquatable<BorderKey>
        {
            public BorderLine Left, Right, Top, Bottom;

            public bool Equals(BorderKey o)
            {
                return o != null && Left.Equals(o.Left) && Right.Equals(o.Right) && Top.Equals(o.Top) && Bottom.Equals(o.Bottom);
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as BorderKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return ((Left.GetHashCode() * 31 + Right.GetHashCode()) * 31 + Top.GetHashCode()) * 31 + Bottom.GetHashCode();
                }
            }
        }

        private readonly List<RunFormat> _fonts = new List<RunFormat>();
        private readonly Dictionary<RunFormat, int> _fontIndex = new Dictionary<RunFormat, int>();
        private readonly List<Rgb> _fills = new List<Rgb>();
        private readonly Dictionary<Rgb, int> _fillIndex = new Dictionary<Rgb, int>();
        private readonly List<BorderKey> _borders = new List<BorderKey>();
        private readonly Dictionary<BorderKey, int> _borderIndex = new Dictionary<BorderKey, int>();
        private readonly List<KeyValuePair<int, string>> _numFmts = new List<KeyValuePair<int, string>>();
        private readonly Dictionary<string, int> _numFmtIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<XfKey> _xfs = new List<XfKey>();
        private readonly Dictionary<XfKey, int> _xfIndex = new Dictionary<XfKey, int>();

        public XlsxStyleRegistry()
        {
            GetFontId(new RunFormat());
            GetBorderId(new CellBorders());
            GetXfId(0, 0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false, 0);
        }

        public static RunFormat NormalizeFont(RunFormat format)
        {
            var f = format == null ? new RunFormat() : format.Clone();
            if (string.IsNullOrEmpty(f.FontName)) f.FontName = DefaultFontName;
            if (!f.Size.HasValue || f.Size.Value <= 0 || f.Size.Value > 409) f.Size = DefaultFontSize;
            f.Size = Math.Round(f.Size.Value * 2, MidpointRounding.AwayFromZero) / 2; // Excel : demi-points
            return f;
        }

        public int GetFontId(RunFormat format)
        {
            var key = NormalizeFont(format);
            int id;
            if (_fontIndex.TryGetValue(key, out id)) return id;
            id = _fonts.Count;
            _fonts.Add(key);
            _fontIndex[key] = id;
            return id;
        }

        /// <summary>Remplissage : 0 = aucun, 1 = gray125 (réservé par Excel), puis couleurs unies.</summary>
        public int GetFillId(Rgb? color)
        {
            if (!color.HasValue) return 0;
            int id;
            if (_fillIndex.TryGetValue(color.Value, out id)) return id;
            id = _fills.Count + 2;
            _fills.Add(color.Value);
            _fillIndex[color.Value] = id;
            return id;
        }

        public int GetBorderId(CellBorders borders)
        {
            var key = new BorderKey
            {
                Left = borders.Left ?? BorderLine.None,
                Right = borders.Right ?? BorderLine.None,
                Top = borders.Top ?? BorderLine.None,
                Bottom = borders.Bottom ?? BorderLine.None
            };
            int id;
            if (_borderIndex.TryGetValue(key, out id)) return id;
            id = _borders.Count;
            _borders.Add(key);
            _borderIndex[key] = id;
            return id;
        }

        /// <summary>Format numérique personnalisé (identifiants à partir de 164) ; null = Standard.</summary>
        public int GetNumFmtId(string code)
        {
            if (string.IsNullOrEmpty(code)) return 0;
            int id;
            if (_numFmtIndex.TryGetValue(code, out id)) return id;
            id = 164 + _numFmts.Count;
            _numFmts.Add(new KeyValuePair<int, string>(id, code));
            _numFmtIndex[code] = id;
            return id;
        }

        public int GetXfId(int fontId, int fillId, int borderId, int numFmtId, HorizontalAlignment horizontal, VerticalAlignment vertical, bool wrap, int rotation)
        {
            var key = new XfKey
            {
                Font = fontId,
                Fill = fillId,
                Border = borderId,
                NumFmt = numFmtId,
                Horizontal = horizontal,
                Vertical = vertical,
                Wrap = wrap,
                Rotation = rotation
            };
            int id;
            if (_xfIndex.TryGetValue(key, out id)) return id;
            id = _xfs.Count;
            _xfs.Add(key);
            _xfIndex[key] = id;
            return id;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("styleSheet", XlsxNames.Main);

            if (_numFmts.Count > 0)
            {
                w.WriteStartElement("numFmts");
                w.WriteAttributeString("count", Str(_numFmts.Count));
                foreach (var nf in _numFmts)
                {
                    w.WriteStartElement("numFmt");
                    w.WriteAttributeString("numFmtId", Str(nf.Key));
                    w.WriteAttributeString("formatCode", nf.Value);
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }

            w.WriteStartElement("fonts");
            w.WriteAttributeString("count", Str(_fonts.Count));
            foreach (var f in _fonts)
            {
                w.WriteStartElement("font");
                WriteFontProperties(w, f, "name");
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("fills");
            w.WriteAttributeString("count", Str(_fills.Count + 2));
            WritePatternFill(w, "none", null);
            WritePatternFill(w, "gray125", null);
            foreach (var fill in _fills) WritePatternFill(w, "solid", fill);
            w.WriteEndElement();

            w.WriteStartElement("borders");
            w.WriteAttributeString("count", Str(_borders.Count));
            foreach (var b in _borders)
            {
                w.WriteStartElement("border");
                WriteBorderEdge(w, "left", b.Left);
                WriteBorderEdge(w, "right", b.Right);
                WriteBorderEdge(w, "top", b.Top);
                WriteBorderEdge(w, "bottom", b.Bottom);
                w.WriteStartElement("diagonal");
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("cellStyleXfs");
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("xf");
            w.WriteAttributeString("numFmtId", "0");
            w.WriteAttributeString("fontId", "0");
            w.WriteAttributeString("fillId", "0");
            w.WriteAttributeString("borderId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cellXfs");
            w.WriteAttributeString("count", Str(_xfs.Count));
            foreach (var xf in _xfs)
            {
                w.WriteStartElement("xf");
                w.WriteAttributeString("numFmtId", Str(xf.NumFmt));
                w.WriteAttributeString("fontId", Str(xf.Font));
                w.WriteAttributeString("fillId", Str(xf.Fill));
                w.WriteAttributeString("borderId", Str(xf.Border));
                w.WriteAttributeString("xfId", "0");
                if (xf.NumFmt != 0) w.WriteAttributeString("applyNumberFormat", "1");
                if (xf.Font != 0) w.WriteAttributeString("applyFont", "1");
                if (xf.Fill != 0) w.WriteAttributeString("applyFill", "1");
                if (xf.Border != 0) w.WriteAttributeString("applyBorder", "1");
                bool hasAlignment = xf.Horizontal != HorizontalAlignment.General || xf.Vertical != VerticalAlignment.Bottom || xf.Wrap || xf.Rotation != 0;
                if (hasAlignment)
                {
                    w.WriteAttributeString("applyAlignment", "1");
                    w.WriteStartElement("alignment");
                    if (xf.Horizontal != HorizontalAlignment.General) w.WriteAttributeString("horizontal", HorizontalName(xf.Horizontal));
                    if (xf.Vertical != VerticalAlignment.Bottom) w.WriteAttributeString("vertical", xf.Vertical == VerticalAlignment.Top ? "top" : "center");
                    if (xf.Rotation != 0) w.WriteAttributeString("textRotation", Str(xf.Rotation));
                    if (xf.Wrap) w.WriteAttributeString("wrapText", "1");
                    w.WriteEndElement();
                }
                w.WriteEndElement();
            }
            w.WriteEndElement();

            w.WriteStartElement("cellStyles");
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("cellStyle");
            w.WriteAttributeString("name", "Normal");
            w.WriteAttributeString("xfId", "0");
            w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("dxfs");
            w.WriteAttributeString("count", "0");
            w.WriteEndElement();
            w.WriteStartElement("tableStyles");
            w.WriteAttributeString("count", "0");
            w.WriteEndElement();

            w.WriteEndElement();
        }

        /// <summary>Propriétés de police (partagées entre &lt;font&gt; des styles et &lt;rPr&gt; du texte enrichi).</summary>
        public static void WriteFontProperties(XmlWriter w, RunFormat format, string nameElement)
        {
            var f = NormalizeFont(format);
            if (f.Bold) Empty(w, "b");
            if (f.Italic) Empty(w, "i");
            if (f.Strike) Empty(w, "strike");
            if (f.Underline == UnderlineKind.Single) Empty(w, "u");
            else if (f.Underline == UnderlineKind.Double) ValElement(w, "u", "double");
            if (f.Position == VerticalPosition.Superscript) ValElement(w, "vertAlign", "superscript");
            else if (f.Position == VerticalPosition.Subscript) ValElement(w, "vertAlign", "subscript");
            ValElement(w, "sz", f.Size.Value.ToString("0.##", CultureInfo.InvariantCulture));
            if (f.Color.HasValue)
            {
                w.WriteStartElement("color");
                w.WriteAttributeString("rgb", f.Color.Value.ToArgbHex());
                w.WriteEndElement();
            }
            ValElement(w, nameElement, f.FontName);
        }

        private static void WritePatternFill(XmlWriter w, string pattern, Rgb? color)
        {
            w.WriteStartElement("fill");
            w.WriteStartElement("patternFill");
            w.WriteAttributeString("patternType", pattern);
            if (color.HasValue)
            {
                w.WriteStartElement("fgColor");
                w.WriteAttributeString("rgb", color.Value.ToArgbHex());
                w.WriteEndElement();
                w.WriteStartElement("bgColor");
                w.WriteAttributeString("indexed", "64");
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteBorderEdge(XmlWriter w, string name, BorderLine line)
        {
            w.WriteStartElement(name);
            if (line != null && line.IsVisible)
            {
                w.WriteAttributeString("style", BorderStyleName(line.Style));
                w.WriteStartElement("color");
                if (line.Color.HasValue) w.WriteAttributeString("rgb", line.Color.Value.ToArgbHex());
                else w.WriteAttributeString("auto", "1");
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static string BorderStyleName(BorderStyle style)
        {
            switch (style)
            {
                case BorderStyle.Hair: return "hair";
                case BorderStyle.Medium: return "medium";
                case BorderStyle.Thick: return "thick";
                case BorderStyle.Double: return "double";
                case BorderStyle.Dotted: return "dotted";
                case BorderStyle.Dashed: return "dashed";
                case BorderStyle.MediumDashed: return "mediumDashed";
                case BorderStyle.DashDot: return "dashDot";
                case BorderStyle.MediumDashDot: return "mediumDashDot";
                case BorderStyle.DashDotDot: return "dashDotDot";
                case BorderStyle.MediumDashDotDot: return "mediumDashDotDot";
                default: return "thin";
            }
        }

        private static string HorizontalName(HorizontalAlignment h)
        {
            switch (h)
            {
                case HorizontalAlignment.Left: return "left";
                case HorizontalAlignment.Center: return "center";
                case HorizontalAlignment.Right: return "right";
                case HorizontalAlignment.Justify: return "justify";
                case HorizontalAlignment.Distributed: return "distributed";
                default: return "general";
            }
        }

        private static void Empty(XmlWriter w, string name)
        {
            w.WriteStartElement(name);
            w.WriteEndElement();
        }

        private static void ValElement(XmlWriter w, string name, string value)
        {
            w.WriteStartElement(name);
            w.WriteAttributeString("val", value);
            w.WriteEndElement();
        }

        private static string Str(int v)
        {
            return v.ToString(CultureInfo.InvariantCulture);
        }
    }

    internal static class XlsxNames
    {
        public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    }
}
