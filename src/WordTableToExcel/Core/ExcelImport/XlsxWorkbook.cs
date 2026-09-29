using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.ExcelImport
{
    public enum XlsxSheetState
    {
        Visible,
        Hidden,
        VeryHidden
    }

    public enum XlsxSheetKind
    {
        Worksheet,
        Chartsheet,
        Dialogsheet,
        Macrosheet
    }

    /// <summary>Feuille déclarée dans le classeur (xl/workbook.xml).</summary>
    public sealed class XlsxSheetInfo
    {
        /// <summary>Position dans le classeur (0 = première feuille).</summary>
        public int Index;
        public string Name;
        public XlsxSheetState State;
        public XlsxSheetKind Kind;
        /// <summary>Partie du paquet (« xl/worksheets/sheet1.xml ») ; null si introuvable.</summary>
        public string PartName;
        /// <summary>Zone d'impression définie dans Excel (première zone si plusieurs).</summary>
        public CellRange? PrintArea;
        /// <summary>true si la zone d'impression est composée de plusieurs plages.</summary>
        public bool PrintAreaHasSeveralRanges;
        /// <summary>Lignes à répéter en haut de chaque page (titres d'impression).</summary>
        public int? PrintTitleFirstRow;
        public int? PrintTitleLastRow;

        public bool IsVisible
        {
            get { return State == XlsxSheetState.Visible; }
        }

        public override string ToString()
        {
            return Name;
        }
    }

    /// <summary>Segment de texte enrichi (chaîne partagée ou chaîne en ligne).</summary>
    public sealed class XlsxTextRun
    {
        public string Text;
        /// <summary>Propriétés de police du segment (rPr) ; null = police de la cellule.</summary>
        public XElement Properties;
    }

    /// <summary>Texte d'une cellule, éventuellement composé de segments de mises en forme différentes.</summary>
    public sealed class XlsxRichText
    {
        public static readonly XlsxRichText Empty = new XlsxRichText(string.Empty, null);

        public XlsxRichText(string text, List<XlsxTextRun> runs)
        {
            Text = text ?? string.Empty;
            Runs = runs != null && runs.Count > 0 ? runs : null;
        }

        public readonly string Text;
        /// <summary>null si le texte est homogène (mise en forme de la cellule).</summary>
        public readonly List<XlsxTextRun> Runs;

        public bool IsRich
        {
            get { return Runs != null; }
        }
    }

    /// <summary>
    /// Classeur Excel (.xlsx, .xlsm, .xltx, .xltm) lu sans Excel : feuilles, chaînes partagées,
    /// styles, thème. Le fichier n'est jamais modifié.
    /// </summary>
    public sealed class XlsxWorkbook
    {
        /// <summary>Taille maximale acceptée pour un fichier (octets).</summary>
        public const long MaxFileSize = 200L * 1024 * 1024;

        private readonly ZipReader _zip;
        private readonly List<XlsxSheetInfo> _sheets = new List<XlsxSheetInfo>();
        private readonly List<XlsxRichText> _sharedStrings = new List<XlsxRichText>();

        private XlsxWorkbook(ZipReader zip)
        {
            _zip = zip;
            Warnings = new List<string>();
        }

        public IList<XlsxSheetInfo> Sheets
        {
            get { return _sheets; }
        }

        public IList<XlsxRichText> SharedStrings
        {
            get { return _sharedStrings; }
        }

        /// <summary>Calendrier 1904 (classeurs Mac anciens).</summary>
        public bool Date1904 { get; private set; }

        /// <summary>true si Excel doit recalculer les formules à l'ouverture : les valeurs enregistrées peuvent être périmées.</summary>
        public bool FullCalculationOnLoad { get; private set; }

        public XlsxStyles Styles { get; private set; }

        /// <summary>Remarques non bloquantes sur le classeur.</summary>
        public List<string> Warnings { get; private set; }

        /// <summary>Lit un classeur sur disque, même s'il est ouvert dans Excel (lecture partagée).</summary>
        public static XlsxWorkbook Load(string path)
        {
            byte[] data;
            try
            {
                var info = new FileInfo(path);
                if (!info.Exists) throw new ExcelImportException("Le fichier est introuvable :\n" + path);
                if (info.Length > MaxFileSize) throw new ExcelImportException("Le fichier est trop volumineux pour être importé (" + (info.Length / (1024 * 1024)) + " Mo).");
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    data = new byte[fs.Length];
                    int offset = 0;
                    while (offset < data.Length)
                    {
                        int read = fs.Read(data, offset, data.Length - offset);
                        if (read <= 0) break;
                        offset += read;
                    }
                    if (offset != data.Length) Array.Resize(ref data, offset);
                }
            }
            catch (ExcelImportException)
            {
                throw;
            }
            catch (IOException ex)
            {
                throw new ExcelImportException("Impossible de lire le fichier :\n" + path + "\n\n" + ex.Message, ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new ExcelImportException("Accès refusé au fichier :\n" + path, ex);
            }
            return Load(data);
        }

        public static XlsxWorkbook Load(byte[] data)
        {
            bool convertible;
            string problem = XlsxPackage.DiagnoseNonZip(data, out convertible);
            if (problem != null) throw new ExcelImportException(problem, convertible);

            ZipReader zip;
            try
            {
                zip = new ZipReader(data);
            }
            catch (InvalidDataException ex)
            {
                throw new ExcelImportException("Ce fichier n'est pas un classeur Excel valide ou il est endommagé.", ex);
            }

            try
            {
                var workbook = new XlsxWorkbook(zip);
                workbook.Read();
                return workbook;
            }
            catch (InvalidDataException ex)
            {
                throw new ExcelImportException("Le classeur est endommagé : " + ex.Message, ex);
            }
        }

        /// <summary>Lit le contenu d'une feuille de calcul.</summary>
        public XlsxSheet ReadSheet(XlsxSheetInfo info)
        {
            if (info == null) throw new ArgumentNullException("info");
            if (info.Kind != XlsxSheetKind.Worksheet || info.PartName == null || !_zip.Contains(info.PartName))
            {
                throw new ExcelImportException("La feuille « " + info.Name + " » n'est pas une feuille de calcul (graphique, macro ou boîte de dialogue).");
            }
            try
            {
                return XlsxSheet.Read(this, _zip, info);
            }
            catch (InvalidDataException ex)
            {
                throw new ExcelImportException("La feuille « " + info.Name + " » est endommagée : " + ex.Message, ex);
            }
            catch (XmlException ex)
            {
                throw new ExcelImportException("La feuille « " + info.Name + " » est endommagée (XML invalide, ligne " + ex.LineNumber + ").", ex);
            }
        }

        internal XlsxRichText SharedString(int index)
        {
            return index >= 0 && index < _sharedStrings.Count ? _sharedStrings[index] : null;
        }

        // ------------------------------------------------------------------ lecture

        private void Read()
        {
            string workbookPart = FindWorkbookPart();
            var workbook = XlsxPackage.LoadXml(_zip, workbookPart);
            if (workbook == null || workbook.Root == null || !OoxmlXml.Is(workbook.Root, "workbook"))
            {
                throw new ExcelImportException("Le classeur est endommagé (description du classeur illisible).");
            }
            var root = workbook.Root;
            var rels = XlsxPackage.Relationships(_zip, workbookPart);

            var workbookPr = OoxmlXml.Child(root, "workbookPr");
            bool date1904;
            if (OoxmlXml.TryBoolAttr(workbookPr, "date1904", out date1904)) Date1904 = date1904;
            bool fullCalc;
            if (OoxmlXml.TryBoolAttr(OoxmlXml.Child(root, "calcPr"), "fullCalcOnLoad", out fullCalc)) FullCalculationOnLoad = fullCalc;

            // Thème et palette avant les styles (les couleurs en dépendent).
            List<Rgb> theme = null;
            var themeRel = rels.FirstOrDefault(r => !r.External && r.IsOfType("theme"));
            if (themeRel != null) theme = ExcelColors.ReadTheme(SafeLoad(themeRel.Target));
            var stylesRel = rels.FirstOrDefault(r => !r.External && r.IsOfType("styles"));
            XDocument stylesXml = stylesRel != null ? SafeLoad(stylesRel.Target) : null;
            var colors = new ExcelColors(XlsxStyles.ReadCustomPalette(stylesXml), theme);
            Styles = XlsxStyles.Parse(stylesXml, colors);

            var sharedRel = rels.FirstOrDefault(r => !r.External && r.IsOfType("sharedStrings"));
            if (sharedRel != null) ReadSharedStrings(sharedRel.Target);

            ReadSheets(root, rels);
            ReadDefinedNames(root);
        }

        private string FindWorkbookPart()
        {
            foreach (var rel in XlsxPackage.Relationships(_zip, "_rels/.rels"))
            {
                if (!rel.External && rel.IsOfType("officeDocument") && _zip.Contains(rel.Target)) return rel.Target;
            }
            if (_zip.Contains("xl/workbook.xml")) return "xl/workbook.xml";

            if (_zip.Contains("xl/workbook.bin"))
            {
                throw new ExcelImportException("Ce classeur est au format binaire Excel (.xlsb).\n\nOuvrez-le dans Excel et enregistrez-le au format « Classeur Excel (.xlsx) », puis recommencez.", true);
            }
            if (_zip.Contains("content.xml") && _zip.Contains("mimetype"))
            {
                throw new ExcelImportException("Ce fichier est au format OpenDocument (.ods).\n\nOuvrez-le dans Excel ou LibreOffice et enregistrez-le au format « Classeur Excel (.xlsx) », puis recommencez.", true);
            }
            if (_zip.Contains("word/document.xml"))
            {
                throw new ExcelImportException("Ce fichier est un document Word, pas un classeur Excel.");
            }
            if (_zip.Contains("ppt/presentation.xml"))
            {
                throw new ExcelImportException("Ce fichier est une présentation PowerPoint, pas un classeur Excel.");
            }
            throw new ExcelImportException("Ce fichier n'est pas un classeur Excel (.xlsx) valide.");
        }

        /// <summary>Partie facultative : une partie absente ou illisible n'empêche pas l'import.</summary>
        private XDocument SafeLoad(string partName)
        {
            try
            {
                return XlsxPackage.LoadXml(_zip, partName);
            }
            catch (Exception ex)
            {
                if (ex is ExcelImportException || ex is InvalidDataException)
                {
                    Warnings.Add("Partie « " + partName + " » illisible, mise en forme par défaut utilisée.");
                    return null;
                }
                throw;
            }
        }

        private void ReadSheets(XElement root, List<OpcRelationship> rels)
        {
            var sheets = OoxmlXml.Child(root, "sheets");
            if (sheets == null) return;
            int index = 0;
            foreach (var sheet in sheets.Elements())
            {
                if (!OoxmlXml.Is(sheet, "sheet")) continue;
                var info = new XlsxSheetInfo
                {
                    Index = index++,
                    Name = OoxmlXml.Attr(sheet, "name") ?? ("Feuille " + index.ToString(CultureInfo.InvariantCulture))
                };
                string state = (OoxmlXml.Attr(sheet, "state") ?? "visible").ToLowerInvariant();
                info.State = state == "hidden" ? XlsxSheetState.Hidden : state == "veryhidden" ? XlsxSheetState.VeryHidden : XlsxSheetState.Visible;

                string relId = RelationshipId(sheet);
                var rel = relId == null ? null : rels.FirstOrDefault(r => string.Equals(r.Id, relId, StringComparison.Ordinal));
                if (rel != null && !rel.External)
                {
                    info.PartName = rel.Target;
                    if (rel.IsOfType("worksheet")) info.Kind = XlsxSheetKind.Worksheet;
                    else if (rel.IsOfType("chartsheet")) info.Kind = XlsxSheetKind.Chartsheet;
                    else if (rel.IsOfType("dialogsheet")) info.Kind = XlsxSheetKind.Dialogsheet;
                    else info.Kind = XlsxSheetKind.Macrosheet;
                }
                else
                {
                    info.Kind = XlsxSheetKind.Worksheet;
                    info.PartName = null;
                }
                _sheets.Add(info);
            }
        }

        private static string RelationshipId(XElement sheet)
        {
            foreach (var a in sheet.Attributes())
            {
                if (a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None) return a.Value;
            }
            return OoxmlXml.Attr(sheet, "id");
        }

        private void ReadDefinedNames(XElement root)
        {
            var names = OoxmlXml.Child(root, "definedNames");
            if (names == null) return;
            foreach (var dn in names.Elements())
            {
                if (!OoxmlXml.Is(dn, "definedName")) continue;
                string name = OoxmlXml.Attr(dn, "name");
                if (name == null) continue;
                bool printArea = string.Equals(name, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase);
                bool printTitles = string.Equals(name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase);
                if (!printArea && !printTitles) continue;

                int localSheet = OoxmlXml.IntAttr(dn, "localSheetId", -1);
                if (localSheet < 0 || localSheet >= _sheets.Count) continue;
                var info = _sheets[localSheet];
                var areas = SplitAreas(dn.Value);

                if (printArea)
                {
                    foreach (var area in areas)
                    {
                        CellRange range;
                        if (!CellRange.TryParse(area, out range)) continue;
                        if (info.PrintArea == null) info.PrintArea = range;
                        else info.PrintAreaHasSeveralRanges = true;
                    }
                }
                else
                {
                    foreach (var area in areas)
                    {
                        CellRange range;
                        if (!CellRange.TryParse(area, out range)) continue;
                        // Titres de lignes (« $1:$2 ») seulement ; les colonnes répétées n'ont pas d'équivalent Word.
                        if (range.FirstColumn == 0 && range.LastColumn == CellReference.MaxColumns - 1)
                        {
                            info.PrintTitleFirstRow = range.FirstRow;
                            info.PrintTitleLastRow = range.LastRow;
                        }
                    }
                }
            }
        }

        /// <summary>Sépare « 'Feuil 1'!$A$1:$B$2,'Feuil 1'!$D$1:$E$2 » en zones (virgules hors apostrophes).</summary>
        internal static List<string> SplitAreas(string formula)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(formula)) return result;
            var sb = new StringBuilder();
            bool quoted = false;
            foreach (char c in formula.Trim().TrimStart('='))
            {
                if (c == '\'') quoted = !quoted;
                if (c == ',' && !quoted)
                {
                    if (sb.Length > 0) result.Add(sb.ToString().Trim());
                    sb.Length = 0;
                    continue;
                }
                sb.Append(c);
            }
            if (sb.Length > 0) result.Add(sb.ToString().Trim());
            return result;
        }

        private void ReadSharedStrings(string partName)
        {
            XmlReader reader;
            try
            {
                reader = XlsxPackage.OpenReader(_zip, partName);
            }
            catch (InvalidDataException ex)
            {
                throw new ExcelImportException("Le classeur est endommagé (textes des cellules illisibles).", ex);
            }
            if (reader == null) return;
            try
            {
                using (reader)
                {
                    reader.Read();
                    while (!reader.EOF)
                    {
                        // ReadStringItem avance le lecteur après l'élément : pas de Read() supplémentaire.
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si") _sharedStrings.Add(ReadStringItem(reader));
                        else reader.Read();
                    }
                }
            }
            catch (XmlException ex)
            {
                throw new ExcelImportException("Le classeur est endommagé (textes des cellules illisibles).", ex);
            }
        }

        /// <summary>
        /// Lit un élément « si » ou « is » (le lecteur est positionné dessus) : texte simple (t)
        /// ou segments enrichis (r). Le texte phonétique (rPh) est ignoré, comme à l'affichage dans Excel.
        /// Le lecteur est ensuite positionné sur le nœud qui suit l'élément.
        /// </summary>
        internal static XlsxRichText ReadStringItem(XmlReader reader)
        {
            var element = (XElement)XNode.ReadFrom(reader);
            return ReadStringItem(element);
        }

        internal static XlsxRichText ReadStringItem(XElement element)
        {
            var text = new StringBuilder();
            List<XlsxTextRun> runs = null;
            foreach (var child in element.Elements())
            {
                string name = child.Name.LocalName;
                if (name == "t")
                {
                    text.Append(XlsxPackage.DecodeEscapes(child.Value));
                }
                else if (name == "r")
                {
                    var t = OoxmlXml.Child(child, "t");
                    string value = XlsxPackage.DecodeEscapes(t == null ? string.Empty : t.Value);
                    if (runs == null) runs = new List<XlsxTextRun>();
                    runs.Add(new XlsxTextRun { Text = value, Properties = OoxmlXml.Child(child, "rPr") });
                    text.Append(value);
                }
            }
            if (runs != null && runs.All(r => r.Properties == null)) runs = null;
            return new XlsxRichText(text.ToString(), runs);
        }
    }
}
