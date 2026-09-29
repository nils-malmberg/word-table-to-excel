using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Layout
{
    /// <summary>Contenu d'une cellule lu dans le XML : segments de texte mis en forme et alignement.</summary>
    public sealed class XmlCellContent
    {
        /// <summary>Segments bruts (non nettoyés), paragraphes terminés par « \r », comme le texte renvoyé par Word.</summary>
        public readonly List<TextRun> Runs = new List<TextRun>();

        /// <summary>Alignement des paragraphes (valeur WdParagraphAlignment : 0 gauche, 1 centré, 2 droite, 3 justifié…).</summary>
        public int Alignment;
    }

    /// <summary>
    /// Lit le texte et la mise en forme effective des caractères directement dans le XML Open XML d'un tableau
    /// (<c>Range.WordOpenXML</c>), au lieu d'interroger Word caractère par caractère : indispensable quand Word est
    /// piloté depuis un autre programme, où chaque question posée à Word coûte un aller-retour entre processus.
    /// <para>
    /// La mise en forme est résolue comme le fait Word : valeurs par défaut du document, style de tableau (zones
    /// conditionnelles comprises : ligne d'en-tête, bandes…), style de paragraphe, style de caractère, puis mise en
    /// forme directe ; propriétés « bascule » (gras, italique…) combinées entre styles. Le texte suit l'affichage :
    /// résultats des champs (sans leur code), texte masqué et suppressions suivies exclus, insertions incluses.
    /// </para>
    /// </summary>
    internal sealed class WordXmlContentReader
    {
        private const string W2006 = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string WStrict = "http://purl.oclc.org/ooxml/wordprocessingml/main";

        private readonly Dictionary<string, XElement> _paragraphStyles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, XElement> _characterStyles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        private readonly XElement _defaultParagraphStyle;
        private readonly RunProps _defaults;
        private readonly XElement _defaultParagraphProperties;
        private readonly string _majorFont;
        private readonly string _minorFont;
        private readonly Dictionary<string, Rgb> _themeColors = new Dictionary<string, Rgb>(StringComparer.OrdinalIgnoreCase);
        private readonly Func<bool, string> _themeFontFallback;
        private readonly Dictionary<XElement, RunProps> _styleRunProps = new Dictionary<XElement, RunProps>();
        private readonly Dictionary<string, List<XElement>> _paragraphChains = new Dictionary<string, List<XElement>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<RunProps>> _characterChains = new Dictionary<string, List<RunProps>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Lecteur pour le document XML <paramref name="root"/> (paquet « Flat OPC » de Word ou document seul) ;
        /// null si le tableau n'est pas au format Open XML (XML Word 2003 : lecture par Word).
        /// </summary>
        /// <param name="themeFontFallback">Polices du thème (vrai = titres) si le XML ne contient pas le thème.</param>
        public static WordXmlContentReader TryCreate(XElement root, XElement tbl, Func<bool, string> themeFontFallback)
        {
            if (root == null || tbl == null) return null;
            string ns = tbl.Name.NamespaceName;
            if (ns != W2006 && ns != WStrict) return null;
            return new WordXmlContentReader(root, themeFontFallback);
        }

        private WordXmlContentReader(XElement root, Func<bool, string> themeFontFallback)
        {
            _themeFontFallback = themeFontFallback;

            XElement styles = FindPart(root, "/word/styles.xml", "styles");
            if (styles != null)
            {
                foreach (var style in styles.Elements().Where(e => OoxmlXml.Is(e, "style")))
                {
                    string id = OoxmlXml.Attr(style, "styleId");
                    if (string.IsNullOrEmpty(id)) continue;
                    string type = OoxmlXml.Attr(style, "type") ?? "paragraph";
                    bool isDefault = OoxmlXml.Attr(style, "default") != null && OoxmlXml.ParseOnOff(OoxmlXml.Attr(style, "default"));
                    if (string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!_paragraphStyles.ContainsKey(id)) _paragraphStyles[id] = style;
                        if (isDefault && _defaultParagraphStyle == null) _defaultParagraphStyle = style;
                    }
                    else if (string.Equals(type, "character", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!_characterStyles.ContainsKey(id)) _characterStyles[id] = style;
                    }
                }
                var docDefaults = OoxmlXml.Child(styles, "docDefaults");
                _defaults = RunProps.Parse(OoxmlXml.Child(docDefaults, "rPrDefault", "rPr"));
                _defaultParagraphProperties = OoxmlXml.Child(docDefaults, "pPrDefault", "pPr");
            }
            if (_defaults == null)
            {
                // Sans valeurs par défaut dans le document, Word utilise Times New Roman 10 points.
                _defaults = new RunProps { Font = "Times New Roman", Size = 10 };
            }

            XElement theme = FindPart(root, "/word/theme/theme1.xml", "theme");
            var themeElements = OoxmlXml.Child(theme, "themeElements");
            var fontScheme = OoxmlXml.Child(themeElements, "fontScheme");
            _majorFont = NonEmpty(OoxmlXml.Attr(OoxmlXml.Child(fontScheme, "majorFont", "latin"), "typeface"));
            _minorFont = NonEmpty(OoxmlXml.Attr(OoxmlXml.Child(fontScheme, "minorFont", "latin"), "typeface"));
            var colorScheme = OoxmlXml.Child(themeElements, "clrScheme");
            if (colorScheme != null)
            {
                foreach (var entry in colorScheme.Elements())
                {
                    var srgb = OoxmlXml.Child(entry, "srgbClr");
                    Rgb? color = srgb != null ? Rgb.FromHex(OoxmlXml.Attr(srgb, "val")) : Rgb.FromHex(OoxmlXml.Attr(OoxmlXml.Child(entry, "sysClr"), "lastClr"));
                    if (color.HasValue) _themeColors[entry.Name.LocalName] = color.Value;
                }
            }
        }

        private static string NonEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? null : s;
        }

        /// <summary>Partie du paquet (Flat OPC) portant ce nom, sinon premier élément de ce nom dans le document.</summary>
        private static XElement FindPart(XElement root, string partName, string rootLocalName)
        {
            foreach (var part in root.Descendants().Where(e => OoxmlXml.Is(e, "part")))
            {
                if (!string.Equals(OoxmlXml.Attr(part, "name"), partName, StringComparison.OrdinalIgnoreCase)) continue;
                var found = part.Descendants().FirstOrDefault(e => OoxmlXml.Is(e, rootLocalName));
                if (found != null) return found;
            }
            // Pas de paquet (document seul) ou partie nommée autrement : premier élément correspondant hors
            // « stylesWithEffects » (Word 2010), qui double styles.xml.
            foreach (var e in root.DescendantsAndSelf().Where(x => OoxmlXml.Is(x, rootLocalName)))
            {
                var part = e.Ancestors().FirstOrDefault(a => OoxmlXml.Is(a, "part"));
                string name = part == null ? null : OoxmlXml.Attr(part, "name");
                if (name != null && name.IndexOf("WithEffects", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                return e;
            }
            return null;
        }

        // ------------------------------------------------------------------ lecture

        /// <summary>Contenu d'une cellule (<c>w:tc</c>).</summary>
        /// <param name="tableFormats">Mise en forme du style de tableau qui s'applique à la cellule (priorité décroissante).</param>
        public XmlCellContent ReadCell(XElement tc, IList<XElement> tableFormats)
        {
            var state = new WalkState(this, tableFormats ?? new XElement[0], true);
            Block(tc, state);
            var content = new XmlCellContent();
            content.Runs.AddRange(state.Runs);
            // Comme Word : alignement commun des paragraphes, sinon celui du premier.
            content.Alignment = state.Alignments.Count == 0 ? 0 : state.Alignments[0];
            return content;
        }

        /// <summary>Texte visible du tableau entier, dans l'ordre du document (pour vérification avec le texte de Word).</summary>
        public string TableText(XElement tbl)
        {
            var state = new WalkState(this, new XElement[0], false);
            NestedTable(tbl, state);
            return state.Raw.ToString();
        }

        private sealed class WalkState
        {
            public readonly WordXmlContentReader Reader;
            public IList<XElement> TableFormats;
            public readonly bool CollectRuns;
            public readonly List<TextRun> Runs = new List<TextRun>();
            public readonly StringBuilder Raw = new StringBuilder();
            public readonly List<int> Alignments = new List<int>();
            /// <summary>Pile des champs ouverts : vrai tant qu'on est dans le code du champ.</summary>
            public readonly List<bool> Fields = new List<bool>();
            public ParagraphContext Paragraph;

            public WalkState(WordXmlContentReader reader, IList<XElement> tableFormats, bool collectRuns)
            {
                Reader = reader;
                TableFormats = tableFormats;
                CollectRuns = collectRuns;
            }

            public bool InFieldCode
            {
                get
                {
                    for (int i = 0; i < Fields.Count; i++)
                    {
                        if (Fields[i]) return true;
                    }
                    return false;
                }
            }

            public void Add(string text, RunFormat format, Rgb? highlight)
            {
                if (!CollectRuns || string.IsNullOrEmpty(text)) return;
                var last = Runs.Count > 0 ? Runs[Runs.Count - 1] : null;
                var run = new TextRun(text, format, highlight);
                if (last != null && last.HasSameFormatting(run)) last.Text += text;
                else Runs.Add(run);
            }

            /// <summary>Fin de paragraphe : « \r », avec la mise en forme du texte qui précède.</summary>
            public void EndParagraph()
            {
                Raw.Append('\r');
                if (!CollectRuns) return;
                if (Runs.Count > 0) Runs[Runs.Count - 1].Text += "\r";
                else Runs.Add(new TextRun("\r", Reader.Resolve(Paragraph, null).Format, null));
            }
        }

        private void Block(XElement container, WalkState state)
        {
            foreach (var child in container.Elements())
            {
                if (OoxmlXml.Is(child, "p")) Paragraph(child, state);
                else if (OoxmlXml.Is(child, "tbl"))
                {
                    // Tableau imbriqué : son texte fait partie de la cellule (comme dans le texte renvoyé par Word).
                    var outer = state.TableFormats;
                    state.TableFormats = new XElement[0];
                    NestedTable(child, state);
                    state.TableFormats = outer;
                }
                else if (OoxmlXml.Is(child, "sdt"))
                {
                    var content = OoxmlXml.Child(child, "sdtContent");
                    if (content != null) Block(content, state);
                }
                else if (OoxmlXml.Is(child, "customXml") || OoxmlXml.Is(child, "ins") || OoxmlXml.Is(child, "moveTo"))
                {
                    Block(child, state);
                }
                // tcPr, signets, suppressions (del, moveFrom)… : rien d'affiché.
            }
        }

        private void NestedTable(XElement tbl, WalkState state)
        {
            foreach (var tr in WordXmlTableParser.StructuralChildren(tbl, "tr"))
            {
                foreach (var tc in WordXmlTableParser.StructuralChildren(tr, "tc")) Block(tc, state);
                state.EndParagraph(); // marque de fin de ligne
            }
        }

        private void Paragraph(XElement p, WalkState state)
        {
            var pPr = OoxmlXml.Child(p, "pPr");
            var chain = ParagraphChain(OoxmlXml.Attr(OoxmlXml.Child(pPr, "pStyle"), "val"));
            state.Paragraph = new ParagraphContext
            {
                Styles = chain.Select(StyleRunProps).ToList(),
                Table = state.TableFormats.Select(f => StyleRunProps(f)).ToList()
            };
            if (state.CollectRuns) state.Alignments.Add(Alignment(pPr, chain, state.TableFormats));
            Inline(p, state);
            state.EndParagraph();
        }

        private void Inline(XElement parent, WalkState state)
        {
            foreach (var child in parent.Elements())
            {
                if (OoxmlXml.Is(child, "r"))
                {
                    Run(child, state);
                }
                else if (OoxmlXml.Is(child, "hyperlink") || OoxmlXml.Is(child, "smartTag") || OoxmlXml.Is(child, "customXml")
                    || OoxmlXml.Is(child, "ins") || OoxmlXml.Is(child, "moveTo") || OoxmlXml.Is(child, "fldSimple")
                    || OoxmlXml.Is(child, "bdo") || OoxmlXml.Is(child, "dir"))
                {
                    Inline(child, state);
                }
                else if (OoxmlXml.Is(child, "sdt"))
                {
                    var content = OoxmlXml.Child(child, "sdtContent");
                    if (content != null) Inline(content, state);
                }
                else if (OoxmlXml.Is(child, "oMath") || OoxmlXml.Is(child, "oMathPara"))
                {
                    if (state.InFieldCode) continue;
                    // Équation : son texte linéaire, avec la mise en forme du paragraphe.
                    string text = string.Concat(child.Descendants().Where(e => OoxmlXml.Is(e, "t")).Select(e => e.Value).ToArray());
                    state.Raw.Append(text);
                    var effective = Resolve(state.Paragraph, null);
                    state.Add(text, effective.Format, effective.Highlight);
                }
                // del, moveFrom (supprimés), pPr, signets, commentaires, AlternateContent (dessins)… : ignorés.
            }
        }

        private void Run(XElement r, WalkState state)
        {
            RunProps direct = null;
            bool directParsed = false;
            Effective effective = null;
            foreach (var child in r.Elements())
            {
                if (OoxmlXml.Is(child, "rPr")) continue;
                if (OoxmlXml.Is(child, "fldChar"))
                {
                    string type = (OoxmlXml.Attr(child, "fldCharType") ?? string.Empty).ToLowerInvariant();
                    if (type == "begin") state.Fields.Add(true);
                    else if (type == "separate" && state.Fields.Count > 0) state.Fields[state.Fields.Count - 1] = false;
                    else if (type == "end" && state.Fields.Count > 0) state.Fields.RemoveAt(state.Fields.Count - 1);
                    continue;
                }
                if (state.InFieldCode) continue;

                string text = null;
                string rawText = null;
                string symbolFont = null;
                string name = child.Name.LocalName;
                switch (name)
                {
                    case "t":
                        text = child.Value;
                        break;
                    case "tab":
                    case "ptab":
                        text = "\t";
                        break;
                    case "br":
                    {
                        string type = (OoxmlXml.Attr(child, "type") ?? string.Empty).ToLowerInvariant();
                        text = type == "page" ? "\f" : type == "column" ? "\u000E" : "\u000B";
                        break;
                    }
                    case "cr":
                        text = "\u000B";
                        break;
                    case "noBreakHyphen":
                        text = "\u001E";
                        break;
                    case "softHyphen":
                        text = "\u001F";
                        break;
                    case "sym":
                        text = Symbol(child);
                        symbolFont = OoxmlXml.Attr(child, "font");
                        rawText = "("; // Word renvoie « ( » pour un symbole dans le texte d'une plage
                        break;
                }
                if (string.IsNullOrEmpty(text)) continue;

                if (!directParsed)
                {
                    direct = RunProps.Parse(OoxmlXml.Child(r, "rPr"));
                    directParsed = true;
                }
                if (effective == null) effective = Resolve(state.Paragraph, direct);
                if (effective.Hidden) continue;

                state.Raw.Append(rawText ?? text);
                if (!state.CollectRuns) continue;
                string shown = effective.AllCaps ? text.ToUpper(CultureInfo.CurrentCulture) : text;
                RunFormat format = effective.Format;
                if (!string.IsNullOrEmpty(symbolFont))
                {
                    format = format.Clone();
                    format.FontName = symbolFont;
                }
                state.Add(shown, format, effective.Highlight);
            }
        }

        /// <summary>Caractère d'un symbole (w:sym) : les polices de symboles utilisent la zone F000-F0FF.</summary>
        private static string Symbol(XElement sym)
        {
            string hex = OoxmlXml.Attr(sym, "char");
            int code;
            if (string.IsNullOrEmpty(hex) || !int.TryParse(hex.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code)) return null;
            if (code >= 0xF000 && code <= 0xF0FF) code -= 0xF000;
            if (code <= 0 || code > 0x10FFFF || (code >= 0xD800 && code <= 0xDFFF)) return null;
            return char.ConvertFromUtf32(code);
        }

        // ------------------------------------------------------------------ styles

        private sealed class ParagraphContext
        {
            /// <summary>Style de paragraphe et ses parents (le plus dérivé d'abord).</summary>
            public List<RunProps> Styles;
            /// <summary>Style de tableau (priorité décroissante).</summary>
            public List<RunProps> Table;
        }

        private sealed class Effective
        {
            public RunFormat Format;
            public Rgb? Highlight;
            public bool Hidden;
            public bool AllCaps;
        }

        private List<XElement> ParagraphChain(string styleId)
        {
            string key = styleId ?? string.Empty;
            List<XElement> chain;
            if (_paragraphChains.TryGetValue(key, out chain)) return chain;
            XElement style = null;
            if (!string.IsNullOrEmpty(styleId)) _paragraphStyles.TryGetValue(styleId, out style);
            if (style == null) style = _defaultParagraphStyle;
            chain = Chain(style, _paragraphStyles);
            _paragraphChains[key] = chain;
            return chain;
        }

        private List<RunProps> CharacterChain(string styleId)
        {
            if (string.IsNullOrEmpty(styleId)) return new List<RunProps>();
            List<RunProps> chain;
            if (_characterChains.TryGetValue(styleId, out chain)) return chain;
            XElement style;
            _characterStyles.TryGetValue(styleId, out style);
            chain = Chain(style, _characterStyles).Select(StyleRunProps).ToList();
            _characterChains[styleId] = chain;
            return chain;
        }

        private static List<XElement> Chain(XElement style, Dictionary<string, XElement> styles)
        {
            var chain = new List<XElement>();
            var seen = new HashSet<XElement>();
            while (style != null && seen.Add(style) && chain.Count < 20)
            {
                chain.Add(style);
                string basedOn = OoxmlXml.Attr(OoxmlXml.Child(style, "basedOn"), "val");
                XElement parent = null;
                if (!string.IsNullOrEmpty(basedOn)) styles.TryGetValue(basedOn, out parent);
                style = parent;
            }
            return chain;
        }

        /// <summary>Propriétés de caractères d'un style (w:style) ou d'une zone de style de tableau (w:tblStylePr).</summary>
        private RunProps StyleRunProps(XElement style)
        {
            RunProps props;
            if (_styleRunProps.TryGetValue(style, out props)) return props;
            props = RunProps.Parse(OoxmlXml.Child(style, "rPr")) ?? RunProps.Empty;
            _styleRunProps[style] = props;
            return props;
        }

        /// <summary>
        /// Mise en forme effective d'un segment : valeurs par défaut &lt; style de tableau &lt; style de paragraphe &lt;
        /// style de caractère &lt; mise en forme directe. Propriétés « bascule » : chaque niveau de style qui les
        /// active les inverse ; la mise en forme directe impose sa valeur.
        /// </summary>
        private Effective Resolve(ParagraphContext paragraph, RunProps direct)
        {
            var levels = new List<RunProps>();
            var characterChain = CharacterChain(direct == null ? null : direct.StyleId);
            if (direct != null) levels.Add(direct);
            levels.AddRange(characterChain);
            if (paragraph != null) levels.AddRange(paragraph.Styles);
            if (paragraph != null) levels.AddRange(paragraph.Table);
            levels.Add(_defaults);

            var groups = new List<List<RunProps>>();
            if (paragraph != null) groups.Add(paragraph.Table);
            if (paragraph != null) groups.Add(paragraph.Styles);
            groups.Add(characterChain);

            var format = new RunFormat
            {
                FontName = Font(levels),
                Size = First(levels, p => p.Size),
                Bold = Toggle(direct, groups, p => p.Bold),
                Italic = Toggle(direct, groups, p => p.Italic),
                Strike = Toggle(direct, groups, p => p.Strike) || Toggle(direct, groups, p => p.DoubleStrike),
                Underline = Underline(FirstText(levels, p => p.Underline)),
                Position = Position(FirstText(levels, p => p.VertAlign)),
                Color = Color(levels.FirstOrDefault(p => p.Color != null))
            };
            return new Effective
            {
                Format = format,
                Highlight = Highlight(FirstText(levels, p => p.Highlight)),
                Hidden = Toggle(direct, groups, p => p.Vanish),
                AllCaps = Toggle(direct, groups, p => p.Caps)
            };
        }

        private bool Toggle(RunProps direct, List<List<RunProps>> groups, Func<RunProps, bool?> get)
        {
            if (direct != null && get(direct).HasValue) return get(direct).Value;
            bool value = get(_defaults) == true;
            foreach (var group in groups)
            {
                foreach (var props in group)
                {
                    bool? v = get(props);
                    if (!v.HasValue) continue;
                    if (v.Value) value = !value;
                    break; // dans une chaîne de styles, le plus dérivé l'emporte
                }
            }
            return value;
        }

        private static double? First(List<RunProps> levels, Func<RunProps, double?> get)
        {
            foreach (var p in levels)
            {
                var v = get(p);
                if (v.HasValue) return v;
            }
            return null;
        }

        private static string FirstText(List<RunProps> levels, Func<RunProps, string> get)
        {
            foreach (var p in levels)
            {
                var v = get(p);
                if (v != null) return v;
            }
            return null;
        }

        private string Font(List<RunProps> levels)
        {
            foreach (var p in levels)
            {
                if (p.ThemeFont != null)
                {
                    string theme = ThemeFont(p.ThemeFont);
                    if (theme != null) return theme;
                    if (p.Font != null) return p.Font;
                    continue;
                }
                if (p.Font != null) return p.Font;
            }
            return null;
        }

        private string ThemeFont(string themeFont)
        {
            bool major = themeFont.StartsWith("major", StringComparison.OrdinalIgnoreCase);
            string name = major ? _majorFont : _minorFont;
            if (name == null && _themeFontFallback != null)
            {
                try
                {
                    name = NonEmpty(_themeFontFallback(major));
                }
                catch (Exception)
                {
                    name = null;
                }
            }
            return name;
        }

        private static UnderlineKind Underline(string value)
        {
            if (value == null) return UnderlineKind.None;
            switch (value.ToLowerInvariant())
            {
                case "none":
                    return UnderlineKind.None;
                case "double":
                case "wavydouble":
                    return UnderlineKind.Double;
                default:
                    return UnderlineKind.Single;
            }
        }

        private static VerticalPosition Position(string value)
        {
            if (value == null) return VerticalPosition.Baseline;
            switch (value.ToLowerInvariant())
            {
                case "superscript": return VerticalPosition.Superscript;
                case "subscript": return VerticalPosition.Subscript;
                default: return VerticalPosition.Baseline;
            }
        }

        /// <summary>Couleur du texte : valeur RVB enregistrée par Word, sinon couleur du thème (éclaircie / assombrie).</summary>
        private Rgb? Color(RunProps props)
        {
            if (props == null) return null;
            var color = props.Color;
            string val = OoxmlXml.Attr(color, "val");
            if (val != null && string.Equals(val.Trim(), "auto", StringComparison.OrdinalIgnoreCase)) return null;
            Rgb? rgb = Rgb.FromHex(val);
            if (rgb.HasValue) return rgb;

            string themeColor = OoxmlXml.Attr(color, "themeColor");
            if (themeColor == null) return null;
            Rgb baseColor;
            if (!_themeColors.TryGetValue(SchemeName(themeColor), out baseColor)) return null;
            int tint = HexByte(OoxmlXml.Attr(color, "themeTint"));
            int shade = HexByte(OoxmlXml.Attr(color, "themeShade"));
            if (tint >= 0 && tint < 0xFF) baseColor = baseColor.Blend(Rgb.White, 1 - tint / 255.0);
            if (shade >= 0 && shade < 0xFF) baseColor = Rgb.Black.Blend(baseColor, shade / 255.0);
            return baseColor;
        }

        private static int HexByte(string hex)
        {
            int v;
            return hex != null && int.TryParse(hex.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v) ? v : -1;
        }

        /// <summary>Nom de couleur de thème Word → entrée du jeu de couleurs (correspondance par défaut).</summary>
        private static string SchemeName(string themeColor)
        {
            switch (themeColor.ToLowerInvariant())
            {
                case "dark1":
                case "text1": return "dk1";
                case "light1":
                case "background1": return "lt1";
                case "dark2":
                case "text2": return "dk2";
                case "light2":
                case "background2": return "lt2";
                case "hyperlink": return "hlink";
                case "followedhyperlink": return "folHlink";
                default: return themeColor; // accent1…accent6
            }
        }

        private static Rgb? Highlight(string name)
        {
            if (name == null) return null;
            switch (name.ToLowerInvariant())
            {
                case "black": return new Rgb(0x00, 0x00, 0x00);
                case "blue": return new Rgb(0x00, 0x00, 0xFF);
                case "cyan": return new Rgb(0x00, 0xFF, 0xFF);
                case "green": return new Rgb(0x00, 0xFF, 0x00);
                case "magenta": return new Rgb(0xFF, 0x00, 0xFF);
                case "red": return new Rgb(0xFF, 0x00, 0x00);
                case "yellow": return new Rgb(0xFF, 0xFF, 0x00);
                case "white": return new Rgb(0xFF, 0xFF, 0xFF);
                case "darkblue": return new Rgb(0x00, 0x00, 0x80);
                case "darkcyan": return new Rgb(0x00, 0x80, 0x80);
                case "darkgreen": return new Rgb(0x00, 0x80, 0x00);
                case "darkmagenta": return new Rgb(0x80, 0x00, 0x80);
                case "darkred": return new Rgb(0x80, 0x00, 0x00);
                case "darkyellow": return new Rgb(0x80, 0x80, 0x00);
                case "darkgray": return new Rgb(0x80, 0x80, 0x80);
                case "lightgray": return new Rgb(0xC0, 0xC0, 0xC0);
                default: return null;
            }
        }

        /// <summary>Alignement d'un paragraphe : direct, style de paragraphe, style de tableau, valeur par défaut.</summary>
        private int Alignment(XElement pPr, List<XElement> paragraphChain, IList<XElement> tableFormats)
        {
            string jc = OoxmlXml.Attr(OoxmlXml.Child(pPr, "jc"), "val");
            foreach (var style in paragraphChain)
            {
                if (jc != null) break;
                jc = OoxmlXml.Attr(OoxmlXml.Child(style, "pPr", "jc"), "val");
            }
            foreach (var format in tableFormats)
            {
                if (jc != null) break;
                jc = OoxmlXml.Attr(OoxmlXml.Child(format, "pPr", "jc"), "val");
            }
            if (jc == null) jc = OoxmlXml.Attr(OoxmlXml.Child(_defaultParagraphProperties, "jc"), "val");
            switch ((jc ?? "left").ToLowerInvariant())
            {
                case "center": return 1;
                case "right":
                case "end": return 2;
                case "both":
                case "justify": return 3;
                case "distribute": return 4;
                case "mediumkashida": return 5;
                case "highkashida": return 7;
                case "lowkashida": return 8;
                case "thaidistribute": return 9;
                default: return 0;
            }
        }

        /// <summary>Propriétés de caractères d'un niveau (null = non définie à ce niveau).</summary>
        private sealed class RunProps
        {
            public static readonly RunProps Empty = new RunProps();

            public bool? Bold, Italic, Caps, Strike, DoubleStrike, Vanish;
            public string Font;
            public string ThemeFont;
            public double? Size;
            public XElement Color;
            public string Highlight;
            public string Underline;
            public string VertAlign;
            public string StyleId;

            public static RunProps Parse(XElement rPr)
            {
                if (rPr == null) return null;
                var p = new RunProps();
                foreach (var e in rPr.Elements())
                {
                    switch (e.Name.LocalName)
                    {
                        case "b": p.Bold = OoxmlXml.IsOn(e); break;
                        case "i": p.Italic = OoxmlXml.IsOn(e); break;
                        case "caps": p.Caps = OoxmlXml.IsOn(e); break;
                        case "strike": p.Strike = OoxmlXml.IsOn(e); break;
                        case "dstrike": p.DoubleStrike = OoxmlXml.IsOn(e); break;
                        case "vanish": p.Vanish = OoxmlXml.IsOn(e); break;
                        case "rFonts":
                            p.ThemeFont = NonEmpty(OoxmlXml.Attr(e, "asciiTheme"));
                            p.Font = NonEmpty(OoxmlXml.Attr(e, "ascii"));
                            break;
                        case "sz":
                        {
                            int halfPoints = OoxmlXml.IntAttr(e, "val", -1);
                            if (halfPoints > 0) p.Size = halfPoints / 2.0;
                            break;
                        }
                        case "color": p.Color = e; break;
                        case "highlight": p.Highlight = OoxmlXml.Attr(e, "val") ?? "none"; break;
                        case "u": p.Underline = OoxmlXml.Attr(e, "val") ?? "single"; break;
                        case "vertAlign": p.VertAlign = OoxmlXml.Attr(e, "val") ?? "baseline"; break;
                        case "rStyle": p.StyleId = OoxmlXml.Attr(e, "val"); break;
                    }
                }
                return p;
            }
        }
    }
}
