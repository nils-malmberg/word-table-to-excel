using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Xlsx;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Écrit un <see cref="TableModel"/> sous forme de tableau WordprocessingML, avec une mise en forme
    /// entièrement explicite (polices, tailles, couleurs, bordures, trames, espacements) pour que le résultat
    /// ne dépende pas des styles du document de destination.
    /// </summary>
    public static class WordTableXmlWriter
    {
        public const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string OfficeDocumentType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string DocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

        /// <summary>Marge intérieure gauche et droite des cellules (twips).</summary>
        public const int CellMarginTwips = 30;
        /// <summary>Hauteur maximale d'une ligne sans risque de texte caché en bas de page (points).</summary>
        private const double MaxUnsplittableRowPt = 300;

        /// <summary>Paquet « Flat OPC » accepté par Range.InsertXML (Word 2007 et suivants).</summary>
        public static string BuildFlatOpc(TableModel table)
        {
            var sb = new StringBuilder();
            var settings = new XmlWriterSettings { OmitXmlDeclaration = false, Indent = false, Encoding = Encoding.UTF8, CheckCharacters = false };
            using (var w = XmlWriter.Create(new StringWriterUtf8(sb), settings))
            {
                w.WriteStartDocument(true);
                w.WriteProcessingInstruction("mso-application", "progid=\"Word.Document\"");
                w.WriteStartElement("pkg", "package", PackageNamespace);

                w.WriteStartElement("pkg", "part", PackageNamespace);
                w.WriteAttributeString("pkg", "name", PackageNamespace, "/_rels/.rels");
                w.WriteAttributeString("pkg", "contentType", PackageNamespace, "application/vnd.openxmlformats-package.relationships+xml");
                w.WriteStartElement("pkg", "xmlData", PackageNamespace);
                WriteRootRelationships(w);
                w.WriteEndElement();
                w.WriteEndElement();

                w.WriteStartElement("pkg", "part", PackageNamespace);
                w.WriteAttributeString("pkg", "name", PackageNamespace, "/word/document.xml");
                w.WriteAttributeString("pkg", "contentType", PackageNamespace, DocumentContentType);
                w.WriteStartElement("pkg", "xmlData", PackageNamespace);
                WriteDocument(w, table);
                w.WriteEndElement();
                w.WriteEndElement();

                w.WriteEndElement();
                w.WriteEndDocument();
            }
            return sb.ToString();
        }

        /// <summary>Document .docx minimal contenant le tableau (insertion de secours par Range.InsertFile).</summary>
        public static byte[] BuildDocx(TableModel table)
        {
            using (var ms = new MemoryStream())
            {
                using (var zip = new ZipWriter(ms))
                {
                    zip.AddEntry("[Content_Types].xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                        + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                        + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                        + "<Override PartName=\"/word/document.xml\" ContentType=\"" + DocumentContentType + "\"/>"
                        + "</Types>");
                    zip.AddEntry("_rels/.rels", Xml(WriteRootRelationships));
                    zip.AddEntry("word/document.xml", Xml(w => WriteDocument(w, table)));
                    zip.Finish();
                }
                return ms.ToArray();
            }
        }

        /// <summary>Élément w:document complet (corps : tableau puis paragraphe vide).</summary>
        public static string BuildDocumentXml(TableModel table)
        {
            return Xml(w => WriteDocument(w, table));
        }

        private static string Xml(Action<XmlWriter> write)
        {
            var sb = new StringBuilder();
            var settings = new XmlWriterSettings { Indent = false, Encoding = Encoding.UTF8, CheckCharacters = false };
            using (var w = XmlWriter.Create(new StringWriterUtf8(sb), settings))
            {
                w.WriteStartDocument(true);
                write(w);
                w.WriteEndDocument();
            }
            return sb.ToString();
        }

        private static void WriteRootRelationships(XmlWriter w)
        {
            w.WriteStartElement("Relationships", RelationshipsNamespace);
            w.WriteStartElement("Relationship", RelationshipsNamespace);
            w.WriteAttributeString("Id", "rId1");
            w.WriteAttributeString("Type", OfficeDocumentType);
            w.WriteAttributeString("Target", "word/document.xml");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteDocument(XmlWriter w, TableModel table)
        {
            w.WriteStartElement("w", "document", WordNamespace);
            w.WriteStartElement("w", "body", WordNamespace);
            WriteTable(w, table);
            // Un corps de document se termine par un paragraphe ; Word le fusionne avec le paragraphe d'insertion.
            w.WriteStartElement("w", "p", WordNamespace);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        // ------------------------------------------------------------------ tableau

        public static void WriteTable(XmlWriter w, TableModel table)
        {
            int[] grid = GridTwips(table);
            int totalWidth = grid.Sum();

            w.WriteStartElement("w", "tbl", WordNamespace);
            w.WriteStartElement("w", "tblPr", WordNamespace);
            if (table.RightToLeft) Element(w, "bidiVisual");
            Element(w, "tblW", "w", Int(totalWidth), "type", "dxa");
            Element(w, "tblInd", "w", "0", "type", "dxa");
            w.WriteStartElement("w", "tblBorders", WordNamespace);
            foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" }) Element(w, edge, "val", "nil");
            w.WriteEndElement();
            Element(w, "tblLayout", "type", "fixed");
            w.WriteStartElement("w", "tblCellMar", WordNamespace);
            Element(w, "top", "w", "0", "type", "dxa");
            Element(w, "left", "w", Int(CellMarginTwips), "type", "dxa");
            Element(w, "bottom", "w", "0", "type", "dxa");
            Element(w, "right", "w", Int(CellMarginTwips), "type", "dxa");
            w.WriteEndElement();
            w.WriteStartElement("w", "tblLook", WordNamespace);
            Attr(w, "val", "0000");
            Attr(w, "firstRow", "0");
            Attr(w, "lastRow", "0");
            Attr(w, "firstColumn", "0");
            Attr(w, "lastColumn", "0");
            Attr(w, "noHBand", "1");
            Attr(w, "noVBand", "1");
            w.WriteEndElement();
            w.WriteEndElement(); // tblPr

            w.WriteStartElement("w", "tblGrid", WordNamespace);
            foreach (var width in grid) Element(w, "gridCol", "w", Int(width));
            w.WriteEndElement();

            var anchors = new Dictionary<long, CellModel>();
            var owner = new CellModel[table.RowCount, table.ColumnCount];
            foreach (var cell in table.Cells)
            {
                if (cell.Row < 0 || cell.Column < 0 || cell.Row >= table.RowCount || cell.Column >= table.ColumnCount) continue;
                anchors[Key(cell.Row, cell.Column)] = cell;
                for (int i = 0; i < cell.RowSpan && cell.Row + i < table.RowCount; i++)
                {
                    for (int j = 0; j < cell.ColumnSpan && cell.Column + j < table.ColumnCount; j++) owner[cell.Row + i, cell.Column + j] = cell;
                }
            }

            for (int r = 0; r < table.RowCount; r++)
            {
                w.WriteStartElement("w", "tr", WordNamespace);
                w.WriteStartElement("w", "trPr", WordNamespace);
                // Ligne insécable (comme dans Excel), sauf si elle risque de dépasser une page : Word masquerait la fin du texte.
                if (EstimatedRowHeight(table, r, grid) <= MaxUnsplittableRowPt) Element(w, "cantSplit");
                double height = table.RowHeightsPt != null && r < table.RowHeightsPt.Length ? table.RowHeightsPt[r] : 0;
                if (height > 0)
                {
                    bool exact = table.RowHeightExact != null && r < table.RowHeightExact.Length && table.RowHeightExact[r];
                    Element(w, "trHeight", "val", Int(Math.Min(31680, (int)Math.Round(height * 20))), "hRule", exact ? "exact" : "atLeast");
                }
                if (r < table.HeaderRowCount) Element(w, "tblHeader");
                w.WriteEndElement();

                int c = 0;
                while (c < table.ColumnCount)
                {
                    var cell = owner[r, c];
                    if (cell == null)
                    {
                        WriteEmptyCell(w, grid[c]);
                        c++;
                        continue;
                    }
                    int span = Math.Min(cell.ColumnSpan, table.ColumnCount - cell.Column);
                    int width = 0;
                    for (int j = 0; j < span; j++) width += grid[cell.Column + j];
                    if (cell.Row == r) WriteCell(w, table, cell, width, span, r);
                    else WriteContinuation(w, table, cell, width, span, r);
                    c = cell.Column + span;
                }
                w.WriteEndElement(); // tr
            }
            w.WriteEndElement(); // tbl
        }

        private static long Key(int row, int column)
        {
            return ((long)row << 16) | (uint)column;
        }

        /// <summary>Largeurs de grille en twips (minimum raisonnable pour une colonne).</summary>
        public static int[] GridTwips(TableModel table)
        {
            var grid = new int[table.ColumnCount];
            for (int i = 0; i < grid.Length; i++)
            {
                double pt = table.ColumnWidthsPt != null && i < table.ColumnWidthsPt.Length ? table.ColumnWidthsPt[i] : 0;
                if (pt <= 0) pt = 48;
                grid[i] = Math.Max(2 * CellMarginTwips + 20, (int)Math.Round(pt * 20));
            }
            return grid;
        }

        private static void WriteCell(XmlWriter w, TableModel table, CellModel cell, int width, int span, int row)
        {
            bool merged = cell.RowSpan > 1 && row + 1 < table.RowCount;
            var borders = cell.Borders ?? new CellBorders();
            w.WriteStartElement("w", "tc", WordNamespace);
            WriteCellProperties(w, cell, width, span, merged ? "restart" : null,
                borders.Top, borders.Left, merged ? BorderLine.None : borders.Bottom, borders.Right);
            WriteParagraphs(w, table, cell);
            w.WriteEndElement();
        }

        /// <summary>Cellule de continuation d'une fusion verticale : mêmes bordures latérales et même fond.</summary>
        private static void WriteContinuation(XmlWriter w, TableModel table, CellModel cell, int width, int span, int row)
        {
            var borders = cell.Borders ?? new CellBorders();
            bool last = row == Math.Min(table.RowCount - 1, cell.Row + cell.RowSpan - 1);
            w.WriteStartElement("w", "tc", WordNamespace);
            WriteCellProperties(w, cell, width, span, "continue", BorderLine.None, borders.Left, last ? borders.Bottom : BorderLine.None, borders.Right);
            w.WriteStartElement("w", "p", WordNamespace);
            WriteParagraphProperties(w, table, cell, cell.PrimaryFormat);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteEmptyCell(XmlWriter w, int width)
        {
            w.WriteStartElement("w", "tc", WordNamespace);
            w.WriteStartElement("w", "tcPr", WordNamespace);
            Element(w, "tcW", "w", Int(width), "type", "dxa");
            w.WriteEndElement();
            w.WriteStartElement("w", "p", WordNamespace);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteCellProperties(XmlWriter w, CellModel cell, int width, int span, string vMerge,
            BorderLine top, BorderLine left, BorderLine bottom, BorderLine right)
        {
            w.WriteStartElement("w", "tcPr", WordNamespace);
            Element(w, "tcW", "w", Int(width), "type", "dxa");
            if (span > 1) Element(w, "gridSpan", "val", Int(span));
            if (vMerge == "restart") Element(w, "vMerge", "val", "restart");
            else if (vMerge == "continue") Element(w, "vMerge");

            w.WriteStartElement("w", "tcBorders", WordNamespace);
            WriteBorder(w, "top", top);
            WriteBorder(w, "left", left);
            WriteBorder(w, "bottom", bottom);
            WriteBorder(w, "right", right);
            w.WriteEndElement();

            Element(w, "shd", "val", "clear", "color", "auto", "fill", cell.Fill.HasValue ? cell.Fill.Value.ToHex() : "auto");
            if (cell.TextRotation == 90) Element(w, "textDirection", "val", "btLr");
            else if (cell.TextRotation == 180) Element(w, "textDirection", "val", "tbRl");
            Element(w, "vAlign", "val", cell.VerticalAlignment == VerticalAlignment.Center ? "center" : cell.VerticalAlignment == VerticalAlignment.Bottom ? "bottom" : "top");
            w.WriteEndElement();
        }

        private static void WriteBorder(XmlWriter w, string edge, BorderLine line)
        {
            if (line == null || !line.IsVisible)
            {
                Element(w, edge, "val", "nil");
                return;
            }
            string val;
            int size;
            switch (line.Style)
            {
                case BorderStyle.Hair: val = "single"; size = 2; break;
                case BorderStyle.Medium: val = "single"; size = 12; break;
                case BorderStyle.Thick: val = "single"; size = 18; break;
                case BorderStyle.Double: val = "double"; size = 4; break;
                case BorderStyle.Dotted: val = "dotted"; size = 4; break;
                case BorderStyle.Dashed: val = "dashed"; size = 4; break;
                case BorderStyle.MediumDashed: val = "dashed"; size = 12; break;
                case BorderStyle.DashDot: val = "dotDash"; size = 4; break;
                case BorderStyle.MediumDashDot: val = "dotDash"; size = 12; break;
                case BorderStyle.DashDotDot: val = "dotDotDash"; size = 4; break;
                case BorderStyle.MediumDashDotDot: val = "dotDotDash"; size = 12; break;
                default: val = "single"; size = 4; break;
            }
            Element(w, edge, "val", val, "sz", Int(size), "space", "0", "color", line.Color.HasValue ? line.Color.Value.ToHex() : "auto");
        }

        // ------------------------------------------------------------------ paragraphes

        private static void WriteParagraphs(XmlWriter w, TableModel table, CellModel cell)
        {
            var mark = cell.PrimaryFormat ?? new RunFormat();
            w.WriteStartElement("w", "p", WordNamespace);
            WriteParagraphProperties(w, table, cell, mark);
            foreach (var run in cell.Runs)
            {
                if (string.IsNullOrEmpty(run.Text)) continue;
                WriteRun(w, run);
            }
            w.WriteEndElement();
        }

        private static void WriteParagraphProperties(XmlWriter w, TableModel table, CellModel cell, RunFormat mark)
        {
            w.WriteStartElement("w", "pPr", WordNamespace);
            Element(w, "keepNext", "val", "0");
            Element(w, "keepLines", "val", "0");
            Element(w, "pageBreakBefore", "val", "0");
            Element(w, "widowControl", "val", "0");
            if (table.RightToLeft) Element(w, "bidi");
            Element(w, "snapToGrid", "val", "0");
            Element(w, "spacing", "before", "0", "beforeAutospacing", "0", "after", "0", "afterAutospacing", "0", "line", "240", "lineRule", "auto");

            int indent = cell.IndentLevel > 0 ? (int)Math.Round(cell.IndentLevel * SheetConverter.IndentStepPt(mark) * 20) : 0;
            string left = "0", right = "0";
            if (indent > 0)
            {
                if (cell.HorizontalAlignment == HorizontalAlignment.Right) right = Int(indent);
                else if (cell.HorizontalAlignment == HorizontalAlignment.Distributed)
                {
                    left = Int(indent);
                    right = Int(indent);
                }
                else if (cell.HorizontalAlignment != HorizontalAlignment.Center) left = Int(indent);
            }
            Element(w, "ind", "left", left, "right", right, "firstLine", "0");
            Element(w, "contextualSpacing", "val", "0");

            string jc;
            switch (cell.HorizontalAlignment)
            {
                case HorizontalAlignment.Center: jc = "center"; break;
                case HorizontalAlignment.Right: jc = table.RightToLeft ? "left" : "right"; break;
                case HorizontalAlignment.Justify: jc = "both"; break;
                case HorizontalAlignment.Distributed: jc = "distribute"; break;
                default: jc = table.RightToLeft ? "right" : "left"; break;
            }
            Element(w, "jc", "val", jc);
            Element(w, "outlineLvl", "val", "9");
            w.WriteStartElement("w", "rPr", WordNamespace);
            WriteRunProperties(w, mark);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteRun(XmlWriter w, TextRun run)
        {
            w.WriteStartElement("w", "r", WordNamespace);
            w.WriteStartElement("w", "rPr", WordNamespace);
            WriteRunProperties(w, run.Format);
            w.WriteEndElement();

            var text = new StringBuilder();
            Action flush = () =>
            {
                if (text.Length == 0) return;
                w.WriteStartElement("w", "t", WordNamespace);
                w.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                w.WriteString(text.ToString());
                w.WriteEndElement();
                text.Length = 0;
            };
            foreach (char c in run.Text)
            {
                if (c == '\n')
                {
                    flush();
                    Element(w, "br");
                }
                else if (c == '\t')
                {
                    flush();
                    Element(w, "tab");
                }
                else
                {
                    text.Append(c);
                }
            }
            flush();
            w.WriteEndElement();
        }

        private static void WriteRunProperties(XmlWriter w, RunFormat f)
        {
            if (!string.IsNullOrEmpty(f.FontName))
            {
                Element(w, "rFonts", "ascii", f.FontName, "hAnsi", f.FontName, "eastAsia", f.FontName, "cs", f.FontName);
            }
            Element(w, "b", "val", f.Bold ? "1" : "0");
            Element(w, "bCs", "val", f.Bold ? "1" : "0");
            Element(w, "i", "val", f.Italic ? "1" : "0");
            Element(w, "iCs", "val", f.Italic ? "1" : "0");
            Element(w, "caps", "val", "0");
            Element(w, "smallCaps", "val", "0");
            Element(w, "strike", "val", f.Strike ? "1" : "0");
            Element(w, "dstrike", "val", "0");
            Element(w, "vanish", "val", "0");
            // Couleur automatique d'Excel = noir, même sur fond foncé (la couleur « auto » de Word deviendrait blanche).
            Element(w, "color", "val", f.Color.HasValue ? f.Color.Value.ToHex() : "000000");
            Element(w, "spacing", "val", "0");
            Element(w, "w", "val", "100");
            Element(w, "position", "val", "0");
            if (f.Size.HasValue && f.Size.Value > 0)
            {
                string half = Int((int)Math.Round(Math.Min(1638, f.Size.Value) * 2));
                Element(w, "sz", "val", half);
                Element(w, "szCs", "val", half);
            }
            Element(w, "u", "val", f.Underline == UnderlineKind.Double ? "double" : f.Underline == UnderlineKind.Single ? "single" : "none");
            Element(w, "vertAlign", "val", f.Position == VerticalPosition.Superscript ? "superscript" : f.Position == VerticalPosition.Subscript ? "subscript" : "baseline");
        }

        /// <summary>Estimation de la hauteur d'une ligne (points), pour décider si elle peut rester d'un seul tenant.</summary>
        private static double EstimatedRowHeight(TableModel table, int row, int[] grid)
        {
            double tallest = table.RowHeightsPt != null && row < table.RowHeightsPt.Length ? table.RowHeightsPt[row] : 0;
            foreach (var cell in table.Cells)
            {
                if (cell.Row != row || cell.RowSpan > 1) continue;
                double width = 0;
                for (int j = 0; j < cell.ColumnSpan && cell.Column + j < grid.Length; j++) width += grid[cell.Column + j] / 20.0;
                width = Math.Max(10, width - 2 * CellMarginTwips / 20.0);
                var format = cell.PrimaryFormat ?? new RunFormat();
                double lineHeight = (format.Size ?? 11) * 1.25;
                double lines = 0;
                foreach (var paragraph in cell.PlainText.Split('\n'))
                {
                    double textWidth = SheetConverter.EstimateTextWidth(paragraph, format);
                    lines += Math.Max(1, Math.Ceiling(textWidth / width));
                }
                tallest = Math.Max(tallest, lines * lineHeight);
            }
            return tallest;
        }

        // ------------------------------------------------------------------ utilitaires XML

        private static void Element(XmlWriter w, string name, params string[] attributes)
        {
            w.WriteStartElement("w", name, WordNamespace);
            for (int i = 0; i + 1 < attributes.Length; i += 2) Attr(w, attributes[i], attributes[i + 1]);
            w.WriteEndElement();
        }

        private static void Attr(XmlWriter w, string name, string value)
        {
            w.WriteAttributeString("w", name, WordNamespace, value);
        }

        private static string Int(int value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>StringWriter déclarant UTF-8 (sinon la déclaration XML annoncerait UTF-16).</summary>
        private sealed class StringWriterUtf8 : StringWriter
        {
            public StringWriterUtf8(StringBuilder sb)
                : base(sb, CultureInfo.InvariantCulture)
            {
            }

            public override Encoding Encoding
            {
                get { return Encoding.UTF8; }
            }
        }
    }
}
