using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Format de cellule Excel (élément xf de cellXfs).</summary>
    public sealed class XlsxCellFormat
    {
        public static readonly XlsxCellFormat Default = new XlsxCellFormat();

        public int NumberFormatId;
        public int FontId;
        public int FillId;
        public int BorderId;
        /// <summary>general, left, center, right, fill, justify, centerContinuous, distributed (null = general).</summary>
        public string Horizontal;
        /// <summary>top, center, bottom, justify, distributed (null = bottom, valeur par défaut d'Excel).</summary>
        public string Vertical;
        public bool WrapText;
        public int TextRotation;
        public int Indent;
    }

    /// <summary>
    /// Format différentiel (élément dxf) : ne contient que les propriétés qu'il modifie.
    /// Utilisé par les styles de tableau Excel et la mise en forme conditionnelle.
    /// </summary>
    public sealed class XlsxDifferentialFormat
    {
        public bool? Bold;
        public bool? Italic;
        public bool? Strike;
        public UnderlineKind? Underline;
        public bool HasFontColor;
        public Rgb? FontColor;
        public bool HasFill;
        public Rgb? Fill;
        /// <summary>Bordures définies (null = non modifiée).</summary>
        public BorderLine Left, Right, Top, Bottom, InsideHorizontal, InsideVertical;
        public string NumberFormatCode;

        public bool IsEmpty
        {
            get
            {
                return !Bold.HasValue && !Italic.HasValue && !Strike.HasValue && !Underline.HasValue && !HasFontColor && !HasFill
                    && Left == null && Right == null && Top == null && Bottom == null && InsideHorizontal == null && InsideVertical == null
                    && NumberFormatCode == null;
            }
        }
    }

    /// <summary>Style de tableau personnalisé défini dans le classeur (tableStyles).</summary>
    public sealed class XlsxTableStyleDefinition
    {
        public string Name;
        /// <summary>Type d'élément (wholeTable, headerRow, firstRowStripe…) → format différentiel.</summary>
        public readonly Dictionary<string, XlsxDifferentialFormat> Elements = new Dictionary<string, XlsxDifferentialFormat>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Largeur des bandes (firstRowStripe, secondColumnStripe…), 1 par défaut.</summary>
        public readonly Dictionary<string, int> StripeSizes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Feuille de styles d'un classeur (xl/styles.xml), couleurs déjà résolues.</summary>
    public sealed class XlsxStyles
    {
        private readonly List<XlsxDifferentialFormat> _dxfs = new List<XlsxDifferentialFormat>();
        private readonly Dictionary<string, XlsxTableStyleDefinition> _tableStyles = new Dictionary<string, XlsxTableStyleDefinition>(StringComparer.OrdinalIgnoreCase);
        private readonly List<RunFormat> _fonts = new List<RunFormat>();
        private readonly List<Rgb?> _fills = new List<Rgb?>();
        private readonly List<CellBorders> _borders = new List<CellBorders>();
        private readonly List<XlsxCellFormat> _cellFormats = new List<XlsxCellFormat>();
        private readonly Dictionary<int, string> _numberFormats = new Dictionary<int, string>();

        public XlsxStyles(ExcelColors colors)
        {
            Colors = colors ?? ExcelColors.Default;
        }

        public ExcelColors Colors { get; private set; }

        public int CellFormatCount
        {
            get { return _cellFormats.Count; }
        }

        /// <summary>Police par défaut du classeur (police 0 : style Normal).</summary>
        public RunFormat DefaultFont
        {
            get { return _fonts.Count > 0 ? _fonts[0] : new RunFormat { FontName = "Calibri", Size = 11 }; }
        }

        public XlsxCellFormat CellFormat(int index)
        {
            return index >= 0 && index < _cellFormats.Count ? _cellFormats[index] : (_cellFormats.Count > 0 ? _cellFormats[0] : XlsxCellFormat.Default);
        }

        public RunFormat Font(int id)
        {
            return (id >= 0 && id < _fonts.Count ? _fonts[id] : DefaultFont).Clone();
        }

        public Rgb? Fill(int id)
        {
            return id >= 0 && id < _fills.Count ? _fills[id] : null;
        }

        public CellBorders Border(int id)
        {
            return (id >= 0 && id < _borders.Count ? _borders[id] : new CellBorders()).Clone();
        }

        /// <summary>Code du format de nombre : personnalisé (numFmts) ou intégré ; null = Standard.</summary>
        public string NumberFormatCode(int id)
        {
            string code;
            if (_numberFormats.TryGetValue(id, out code)) return code;
            return null;
        }

        public bool HasCustomNumberFormat(int id)
        {
            return _numberFormats.ContainsKey(id);
        }

        /// <summary>Format différentiel n° <paramref name="id"/> (null s'il n'existe pas).</summary>
        public XlsxDifferentialFormat DifferentialFormat(int id)
        {
            return id >= 0 && id < _dxfs.Count ? _dxfs[id] : null;
        }

        /// <summary>Style de tableau personnalisé portant ce nom (null si c'est un style intégré d'Excel).</summary>
        public XlsxTableStyleDefinition TableStyle(string name)
        {
            XlsxTableStyleDefinition style;
            return name != null && _tableStyles.TryGetValue(name, out style) ? style : null;
        }

        /// <summary>Style appliqué aux tableaux qui n'en précisent pas.</summary>
        public string DefaultTableStyle { get; private set; }

        /// <summary>Lit la palette personnalisée (styles.xml › colors › indexedColors), à connaître avant de résoudre les couleurs.</summary>
        public static List<Rgb> ReadCustomPalette(XDocument styles)
        {
            var palette = new List<Rgb>();
            if (styles == null || styles.Root == null) return palette;
            var indexed = OoxmlXml.Child(styles.Root, "colors", "indexedColors");
            if (indexed == null) return palette;
            foreach (var c in indexed.Elements().Where(e => OoxmlXml.Is(e, "rgbColor")))
            {
                var rgb = Rgb.FromHex(OoxmlXml.Attr(c, "rgb"));
                palette.Add(rgb ?? Rgb.Black);
            }
            return palette;
        }

        public static XlsxStyles Parse(XDocument styles, ExcelColors colors)
        {
            var result = new XlsxStyles(colors);
            if (styles == null || styles.Root == null)
            {
                result._fonts.Add(new RunFormat { FontName = "Calibri", Size = 11 });
                return result;
            }
            var root = styles.Root;

            var numFmts = OoxmlXml.Child(root, "numFmts");
            if (numFmts != null)
            {
                foreach (var nf in numFmts.Elements().Where(e => OoxmlXml.Is(e, "numFmt")))
                {
                    int id = OoxmlXml.IntAttr(nf, "numFmtId", -1);
                    string code = OoxmlXml.Attr(nf, "formatCode");
                    if (id >= 0 && code != null) result._numberFormats[id] = code;
                }
            }

            var fonts = OoxmlXml.Child(root, "fonts");
            if (fonts != null)
            {
                foreach (var f in fonts.Elements().Where(e => OoxmlXml.Is(e, "font"))) result._fonts.Add(result.ReadFont(f));
            }
            if (result._fonts.Count == 0) result._fonts.Add(new RunFormat { FontName = "Calibri", Size = 11 });
            // Police sans nom ou sans taille : Excel utilise la police Normal du classeur.
            var normal = result._fonts[0];
            if (string.IsNullOrEmpty(normal.FontName)) normal.FontName = "Calibri";
            if (!normal.Size.HasValue) normal.Size = 11;
            foreach (var f in result._fonts)
            {
                if (string.IsNullOrEmpty(f.FontName)) f.FontName = normal.FontName;
                if (!f.Size.HasValue) f.Size = normal.Size;
            }

            var fills = OoxmlXml.Child(root, "fills");
            if (fills != null)
            {
                foreach (var f in fills.Elements().Where(e => OoxmlXml.Is(e, "fill"))) result._fills.Add(result.ReadFill(f));
            }

            var borders = OoxmlXml.Child(root, "borders");
            if (borders != null)
            {
                foreach (var b in borders.Elements().Where(e => OoxmlXml.Is(e, "border"))) result._borders.Add(result.ReadBorder(b));
            }

            var xfs = OoxmlXml.Child(root, "cellXfs");
            if (xfs != null)
            {
                foreach (var xf in xfs.Elements().Where(e => OoxmlXml.Is(e, "xf")))
                {
                    var format = new XlsxCellFormat
                    {
                        NumberFormatId = OoxmlXml.IntAttr(xf, "numFmtId", 0),
                        FontId = OoxmlXml.IntAttr(xf, "fontId", 0),
                        FillId = OoxmlXml.IntAttr(xf, "fillId", 0),
                        BorderId = OoxmlXml.IntAttr(xf, "borderId", 0)
                    };
                    var alignment = OoxmlXml.Child(xf, "alignment");
                    if (alignment != null)
                    {
                        format.Horizontal = OoxmlXml.Attr(alignment, "horizontal");
                        format.Vertical = OoxmlXml.Attr(alignment, "vertical");
                        bool wrap;
                        if (OoxmlXml.TryBoolAttr(alignment, "wrapText", out wrap)) format.WrapText = wrap;
                        format.TextRotation = OoxmlXml.IntAttr(alignment, "textRotation", 0);
                        format.Indent = Math.Max(0, OoxmlXml.IntAttr(alignment, "indent", 0));
                    }
                    result._cellFormats.Add(format);
                }
            }

            var dxfs = OoxmlXml.Child(root, "dxfs");
            if (dxfs != null)
            {
                foreach (var dxf in dxfs.Elements().Where(e => OoxmlXml.Is(e, "dxf"))) result._dxfs.Add(result.ReadDifferentialFormat(dxf));
            }

            var tableStyles = OoxmlXml.Child(root, "tableStyles");
            if (tableStyles != null)
            {
                result.DefaultTableStyle = OoxmlXml.Attr(tableStyles, "defaultTableStyle");
                foreach (var ts in tableStyles.Elements().Where(e => OoxmlXml.Is(e, "tableStyle")))
                {
                    string name = OoxmlXml.Attr(ts, "name");
                    if (string.IsNullOrEmpty(name)) continue;
                    bool forTables;
                    if (OoxmlXml.TryBoolAttr(ts, "table", out forTables) && !forTables) continue; // style réservé aux tableaux croisés
                    var definition = new XlsxTableStyleDefinition { Name = name };
                    foreach (var element in ts.Elements().Where(e => OoxmlXml.Is(e, "tableStyleElement")))
                    {
                        string type = OoxmlXml.Attr(element, "type");
                        var format = result.DifferentialFormat(OoxmlXml.IntAttr(element, "dxfId", -1));
                        if (type == null) continue;
                        if (format != null) definition.Elements[type] = format;
                        int size = OoxmlXml.IntAttr(element, "size", 1);
                        definition.StripeSizes[type] = Math.Max(1, size);
                    }
                    result._tableStyles[name] = definition;
                }
            }
            return result;
        }

        private XlsxDifferentialFormat ReadDifferentialFormat(XElement dxf)
        {
            var result = new XlsxDifferentialFormat();
            var font = OoxmlXml.Child(dxf, "font");
            if (font != null)
            {
                var b = OoxmlXml.Child(font, "b");
                if (b != null) result.Bold = OoxmlXml.IsOn(b);
                var i = OoxmlXml.Child(font, "i");
                if (i != null) result.Italic = OoxmlXml.IsOn(i);
                var strike = OoxmlXml.Child(font, "strike");
                if (strike != null) result.Strike = OoxmlXml.IsOn(strike);
                var u = OoxmlXml.Child(font, "u");
                if (u != null)
                {
                    string val = (OoxmlXml.Attr(u, "val") ?? "single").ToLowerInvariant();
                    result.Underline = val == "none" ? UnderlineKind.None : val.StartsWith("double", StringComparison.Ordinal) ? UnderlineKind.Double : UnderlineKind.Single;
                }
                var color = OoxmlXml.Child(font, "color");
                if (color != null)
                {
                    result.HasFontColor = true;
                    result.FontColor = Colors.Resolve(color);
                }
            }

            var fill = OoxmlXml.Child(dxf, "fill");
            if (fill != null)
            {
                var pattern = OoxmlXml.Child(fill, "patternFill");
                if (pattern != null)
                {
                    string type = (OoxmlXml.Attr(pattern, "patternType") ?? "solid").ToLowerInvariant();
                    result.HasFill = true;
                    if (type != "none")
                    {
                        // Particularité des dxf : la couleur d'un remplissage uni est portée par bgColor.
                        result.Fill = Colors.Resolve(OoxmlXml.Child(pattern, "bgColor")) ?? Colors.Resolve(OoxmlXml.Child(pattern, "fgColor"));
                        if (!result.Fill.HasValue) result.HasFill = false;
                    }
                }
                else
                {
                    var solid = ReadFill(fill);
                    if (solid.HasValue)
                    {
                        result.HasFill = true;
                        result.Fill = solid;
                    }
                }
            }

            var border = OoxmlXml.Child(dxf, "border");
            if (border != null)
            {
                result.Left = ReadDefinedEdge(OoxmlXml.Child(border, "left") ?? OoxmlXml.Child(border, "start"));
                result.Right = ReadDefinedEdge(OoxmlXml.Child(border, "right") ?? OoxmlXml.Child(border, "end"));
                result.Top = ReadDefinedEdge(OoxmlXml.Child(border, "top"));
                result.Bottom = ReadDefinedEdge(OoxmlXml.Child(border, "bottom"));
                result.InsideHorizontal = ReadDefinedEdge(OoxmlXml.Child(border, "horizontal"));
                result.InsideVertical = ReadDefinedEdge(OoxmlXml.Child(border, "vertical"));
            }

            var numFmt = OoxmlXml.Child(dxf, "numFmt");
            if (numFmt != null) result.NumberFormatCode = OoxmlXml.Attr(numFmt, "formatCode");
            return result;
        }

        /// <summary>Bordure d'un dxf : null si l'élément est absent ou sans style (non modifiée).</summary>
        private BorderLine ReadDefinedEdge(XElement edge)
        {
            if (edge == null) return null;
            if (OoxmlXml.Attr(edge, "style") == null) return null;
            return ReadEdge(edge);
        }

        private RunFormat ReadFont(XElement font)
        {
            var f = new RunFormat();
            var name = OoxmlXml.Child(font, "name") ?? OoxmlXml.Child(font, "rFont");
            f.FontName = OoxmlXml.Attr(name, "val");
            var sz = OoxmlXml.Child(font, "sz");
            double size;
            if (sz != null && double.TryParse(OoxmlXml.Attr(sz, "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out size) && size > 0) f.Size = size;
            f.Bold = OoxmlXml.IsOn(OoxmlXml.Child(font, "b"));
            f.Italic = OoxmlXml.IsOn(OoxmlXml.Child(font, "i"));
            f.Strike = OoxmlXml.IsOn(OoxmlXml.Child(font, "strike"));
            var u = OoxmlXml.Child(font, "u");
            if (u != null)
            {
                string val = (OoxmlXml.Attr(u, "val") ?? "single").ToLowerInvariant();
                f.Underline = val == "none" ? UnderlineKind.None : val.StartsWith("double", StringComparison.Ordinal) ? UnderlineKind.Double : UnderlineKind.Single;
            }
            var va = OoxmlXml.Attr(OoxmlXml.Child(font, "vertAlign"), "val");
            if (va == "superscript") f.Position = VerticalPosition.Superscript;
            else if (va == "subscript") f.Position = VerticalPosition.Subscript;
            f.Color = Colors.Resolve(OoxmlXml.Child(font, "color"));
            return f;
        }

        /// <summary>Police d'un segment de texte enrichi (élément rPr des chaînes partagées).</summary>
        public RunFormat ReadRunFont(XElement rPr, RunFormat cellFont)
        {
            if (rPr == null) return cellFont.Clone();
            var f = ReadFont(rPr);
            // Un segment sans taille ni police hérite de la police de la cellule.
            if (string.IsNullOrEmpty(f.FontName)) f.FontName = cellFont.FontName;
            if (!f.Size.HasValue) f.Size = cellFont.Size;
            if (!f.Color.HasValue && OoxmlXml.Child(rPr, "color") == null) f.Color = cellFont.Color;
            return f;
        }

        private Rgb? ReadFill(XElement fill)
        {
            var pattern = OoxmlXml.Child(fill, "patternFill");
            if (pattern != null)
            {
                string type = (OoxmlXml.Attr(pattern, "patternType") ?? "none").ToLowerInvariant();
                if (type == "none") return null;
                Rgb? fg = Colors.Resolve(OoxmlXml.Child(pattern, "fgColor"));
                Rgb? bg = Colors.Resolve(OoxmlXml.Child(pattern, "bgColor"));
                if (type == "solid") return fg ?? bg ?? Rgb.Black;
                // Motifs : couleur moyenne du motif (Word n'a pas d'équivalent exact dans un tableau).
                double density;
                switch (type)
                {
                    case "gray0625": density = 0.0625; break;
                    case "gray125": density = 0.125; break;
                    case "lightgray": density = 0.25; break;
                    case "mediumgray": density = 0.5; break;
                    case "darkgray": density = 0.75; break;
                    default: density = type.StartsWith("light", StringComparison.Ordinal) ? 0.25 : 0.5; break;
                }
                return (bg ?? Rgb.White).Blend(fg ?? Rgb.Black, density);
            }
            var gradient = OoxmlXml.Child(fill, "gradientFill");
            if (gradient != null)
            {
                var stops = gradient.Elements().Where(e => OoxmlXml.Is(e, "stop")).Select(s => Colors.Resolve(OoxmlXml.Child(s, "color"))).Where(c => c.HasValue).Select(c => c.Value).ToList();
                if (stops.Count > 0) return stops.First().Blend(stops.Last(), 0.5);
            }
            return null;
        }

        private CellBorders ReadBorder(XElement border)
        {
            return new CellBorders
            {
                Left = ReadEdge(OoxmlXml.Child(border, "left") ?? OoxmlXml.Child(border, "start")),
                Right = ReadEdge(OoxmlXml.Child(border, "right") ?? OoxmlXml.Child(border, "end")),
                Top = ReadEdge(OoxmlXml.Child(border, "top")),
                Bottom = ReadEdge(OoxmlXml.Child(border, "bottom"))
            };
        }

        private BorderLine ReadEdge(XElement edge)
        {
            if (edge == null) return BorderLine.None;
            BorderStyle style;
            switch ((OoxmlXml.Attr(edge, "style") ?? "none").ToLowerInvariant())
            {
                case "thin": style = BorderStyle.Thin; break;
                case "medium": style = BorderStyle.Medium; break;
                case "thick": style = BorderStyle.Thick; break;
                case "double": style = BorderStyle.Double; break;
                case "dotted": style = BorderStyle.Dotted; break;
                case "dashed": style = BorderStyle.Dashed; break;
                case "hair": style = BorderStyle.Hair; break;
                case "mediumdashed": style = BorderStyle.MediumDashed; break;
                case "dashdot": style = BorderStyle.DashDot; break;
                case "mediumdashdot": style = BorderStyle.MediumDashDot; break;
                case "slantdashdot": style = BorderStyle.MediumDashDot; break;
                case "dashdotdot": style = BorderStyle.DashDotDot; break;
                case "mediumdashdotdot": style = BorderStyle.MediumDashDotDot; break;
                default: return BorderLine.None;
            }
            return new BorderLine(style, Colors.Resolve(OoxmlXml.Child(edge, "color")));
        }
    }
}
