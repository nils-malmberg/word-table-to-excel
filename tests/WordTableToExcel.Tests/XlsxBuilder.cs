using System.Collections.Generic;
using System.IO;
using System.Text;
using WordTableToExcel.Core.Xlsx;

namespace WordTableToExcel.Tests
{
    /// <summary>Construit un petit classeur .xlsx en mémoire à partir de fragments XML (tests de lecture).</summary>
    internal sealed class XlsxBuilder
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private readonly List<KeyValuePair<string, string>> _sheets = new List<KeyValuePair<string, string>>();
        private readonly List<string> _sheetExtras = new List<string>();

        /// <summary>Contenu de &lt;cellXfs&gt;, &lt;fonts&gt;… (sans l'élément racine styleSheet).</summary>
        public string Styles = "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                             + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                             + "<borders count=\"1\"><border/></borders>"
                             + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>";

        public List<string> SharedStrings = new List<string>();
        public string DefinedNames;
        public bool Date1904;

        /// <summary>Ajoute une feuille ; <paramref name="body"/> = contenu de &lt;worksheet&gt; (sheetData, mergeCells…).</summary>
        public XlsxBuilder Sheet(string name, string body, string extraAttributes = "")
        {
            _sheets.Add(new KeyValuePair<string, string>(name, body));
            _sheetExtras.Add(extraAttributes);
            return this;
        }

        public byte[] Build()
        {
            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipWriter(ms))
                {
                    zip.AddEntry("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                        + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

                    var wb = new StringBuilder();
                    wb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"" + Main + "\" xmlns:r=\"" + R + "\">");
                    if (Date1904) wb.Append("<workbookPr date1904=\"1\"/>");
                    wb.Append("<sheets>");
                    var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    for (int i = 0; i < _sheets.Count; i++)
                    {
                        wb.Append("<sheet name=\"" + _sheets[i].Key + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\" " + _sheetExtras[i] + "/>");
                        rels.Append("<Relationship Id=\"rId" + (i + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + (i + 1) + ".xml\"/>");
                        zip.AddEntry("xl/worksheets/sheet" + (i + 1) + ".xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"" + Main + "\" xmlns:r=\"" + R + "\">" + _sheets[i].Value + "</worksheet>");
                    }
                    wb.Append("</sheets>");
                    if (DefinedNames != null) wb.Append("<definedNames>" + DefinedNames + "</definedNames>");
                    wb.Append("</workbook>");
                    rels.Append("<Relationship Id=\"rIdS\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                    rels.Append("<Relationship Id=\"rIdT\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
                    rels.Append("</Relationships>");
                    zip.AddEntry("xl/workbook.xml", wb.ToString());
                    zip.AddEntry("xl/_rels/workbook.xml.rels", rels.ToString());
                    zip.AddEntry("xl/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"" + Main + "\">" + Styles + "</styleSheet>");
                    var sst = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"" + Main + "\">");
                    foreach (var s in SharedStrings) sst.Append(s.StartsWith("<") ? "<si>" + s + "</si>" : "<si><t xml:space=\"preserve\">" + System.Security.SecurityElement.Escape(s) + "</t></si>");
                    sst.Append("</sst>");
                    zip.AddEntry("xl/sharedStrings.xml", sst.ToString());
                    zip.Finish();
                }
                return ms.ToArray();
            }
        }
    }
}
