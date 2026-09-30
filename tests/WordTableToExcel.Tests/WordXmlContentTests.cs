using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Text;
using Xunit;
using static WordTableToExcel.Tests.TestUtil;

namespace WordTableToExcel.Tests
{
    /// <summary>Paquet « Flat OPC » (comme Range.WordOpenXML) construit à partir d'un .docx.</summary>
    internal static class FlatPackage
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        public static XElement FromDocx(string path)
        {
            var root = new XElement(Pkg + "package", new XAttribute(XNamespace.Xmlns + "pkg", Pkg.NamespaceName));
            using (var zip = ZipFile.OpenRead(path))
            {
                // stylesWithEffects d'abord : le lecteur doit bien prendre styles.xml.
                foreach (var name in new[] { "word/stylesWithEffects.xml", "word/document.xml", "word/styles.xml", "word/theme/theme1.xml" })
                {
                    var entry = zip.GetEntry(name);
                    if (entry == null) continue;
                    XDocument part;
                    using (var stream = entry.Open()) part = XDocument.Load(stream);
                    root.Add(new XElement(Pkg + "part", new XAttribute(Pkg + "name", "/" + name), new XElement(Pkg + "xmlData", part.Root)));
                }
            }
            return root;
        }

        /// <summary>Tableaux du corps du document, dans l'ordre (ceux de Document.Tables dans Word).</summary>
        public static List<XElement> BodyTables(XElement root)
        {
            var body = root.Descendants().First(e => OoxmlXml.Is(e, "body"));
            return WordXmlTableParser.StructuralChildren(body, "tbl").ToList();
        }
    }

    /// <summary>Lecture du texte et de la mise en forme des cellules dans le XML (export rapide de l'application).</summary>
    public class WordXmlContentTests
    {
        private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static string FixturePath(string name)
        {
            return Path.Combine(AppContext.BaseDirectory, "fixtures", name);
        }

        private static List<TextRun> Runs(LayoutCell cell)
        {
            return CellTextSanitizer.Clean(cell.Content.Runs);
        }

        private static string Text(LayoutCell cell)
        {
            return string.Concat(Runs(cell).Select(r => r.Text));
        }

        private static TextRun RunWith(LayoutCell cell, string text)
        {
            return Runs(cell).First(r => r.Text.Contains(text));
        }

        // ------------------------------------------------------------------ document réel

        /// <summary>
        /// Document de test (13 tableaux et un tableau imbriqué, styles de tableau intégrés de Word, champs, fusions, tableau imbriqué) :
        /// la mise en forme lue a été comparée à celle qu'applique LibreOffice au même document (2 582 caractères,
        /// gras, italique, couleur, police, taille, soulignement, barré, exposant/indice et 284 alignements identiques).
        /// </summary>
        [Fact]
        public void ComplexDocument_AllTablesAreReadFromXml()
        {
            var root = FlatPackage.FromDocx(FixturePath("rapport_test_complexe.docx"));
            var tables = FlatPackage.BodyTables(root);
            Assert.Equal(13, tables.Count); // + 1 tableau imbriqué
            var layouts = tables.Select(t => WordXmlTableParser.Parse(root, t)).ToList();
            Assert.All(layouts, l => Assert.True(l.HasContent));
            Assert.All(layouts, l => Assert.All(l.Cells, c => Assert.NotNull(c.Content)));

            // Tableau 1, style « Medium Shading 1 - Accent 1 » : en-tête blanc en gras, première colonne en gras.
            var t1 = layouts[0];
            var header = Runs(At(t1, 0, 0)).Single();
            Assert.Equal("Secteur", header.Text);
            Assert.True(header.Format.Bold);
            Assert.Equal(Hex("FFFFFF"), header.Format.Color);
            Assert.Equal("Calibri", header.Format.FontName); // police du thème (minorHAnsi)
            Assert.Equal(11, header.Format.Size);
            Assert.True(Runs(At(t1, 1, 0)).Single().Format.Bold);
            var value = Runs(At(t1, 1, 1)).Single();
            Assert.Equal("12 540,3", value.Text);
            Assert.False(value.Format.Bold);
            Assert.Equal(Hex("C00000"), Runs(At(t1, 1, 4)).Single().Format.Color);
            Assert.Equal(2, At(t1, 1, 1).Content.Alignment); // aligné à droite

            // Tableau 2 : indices et exposants.
            var t2 = layouts[1];
            Assert.Equal(VerticalPosition.Subscript, RunWith(At(t2, 3, 0), "ref").Format.Position);
            Assert.Equal(VerticalPosition.Superscript, Runs(At(t2, 5, 3)).Last().Format.Position);

            // Tableau 3 : cellule de plusieurs paragraphes, le second en petits caractères italiques.
            var t3 = layouts[2];
            Assert.Equal("Rénovations annuelles\ndont 60 % en maisons individuelles\net 40 % en logements collectifs", Text(At(t3, 3, 0)));
            Assert.Equal(9, RunWith(At(t3, 3, 0), "dont").Format.Size);
            Assert.True(RunWith(At(t3, 3, 0), "dont").Format.Italic);

            // Tableau 4 : en-tête sur deux lignes fusionnées, surlignage partiel.
            var t4 = layouts[3];
            Assert.Equal("2024", Text(At(t4, 0, 2)));
            Assert.Equal(2, At(t4, 0, 2).ColumnSpan);
            Assert.Equal(Hex("FFFF00"), RunWith(At(t4, 8, 5), "record").Highlight);

            // Tableau 10 : tableau imbriqué, son texte suit celui de la cellule (comme dans le texte renvoyé par Word).
            Assert.Equal("Pôles :\nConsommation\n6\n\nProduction\n5", Text(At(layouts[9], 2, 2)));

            // Texte de vérification : celui de toutes les cellules, sans code de champ.
            foreach (var layout in layouts)
            {
                string cells = string.Concat(layout.Cells.Select(c => string.Concat(c.Content.Runs.Select(r => r.Text))));
                Assert.Equal(CellTextSanitizer.Comparable(cells.ToUpperInvariant()), CellTextSanitizer.Comparable(layout.XmlText).ToUpperInvariant());
            }
        }

        // ------------------------------------------------------------------ cas particuliers

        private const string Theme = @"<a:theme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""Office""><a:themeElements>
<a:clrScheme name=""Office""><a:dk1><a:sysClr val=""windowText"" lastClr=""000000""/></a:dk1><a:lt1><a:sysClr val=""window"" lastClr=""FFFFFF""/></a:lt1>
<a:dk2><a:srgbClr val=""1F497D""/></a:dk2><a:lt2><a:srgbClr val=""EEECE1""/></a:lt2><a:accent1><a:srgbClr val=""4F81BD""/></a:accent1></a:clrScheme>
<a:fontScheme name=""Office""><a:majorFont><a:latin typeface=""Cambria""/></a:majorFont><a:minorFont><a:latin typeface=""Calibri""/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>";

        private const string Styles = @"
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=""minorHAnsi"" w:hAnsiTheme=""minorHAnsi""/><w:sz w:val=""22""/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal""><w:name w:val=""Normal""/></w:style>
<w:style w:type=""paragraph"" w:styleId=""Titre""><w:name w:val=""Titre""/><w:basedOn w:val=""Normal""/><w:pPr><w:jc w:val=""right""/></w:pPr><w:rPr><w:rFonts w:asciiTheme=""majorHAnsi""/><w:b/><w:sz w:val=""28""/></w:rPr></w:style>
<w:style w:type=""character"" w:styleId=""Strong""><w:name w:val=""Strong""/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type=""character"" w:styleId=""Emph""><w:name w:val=""Emphase""/><w:basedOn w:val=""Strong""/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type=""character"" w:styleId=""Hyperlink""><w:name w:val=""Hyperlink""/><w:rPr><w:color w:val=""0000FF"" w:themeColor=""hyperlink""/><w:u w:val=""single""/></w:rPr></w:style>
<w:style w:type=""table"" w:styleId=""Grille""><w:name w:val=""Grille""/>
  <w:tblStylePr w:type=""firstRow""><w:pPr><w:jc w:val=""center""/></w:pPr><w:rPr><w:b/><w:color w:val=""FFFFFF"" w:themeColor=""background1""/></w:rPr></w:tblStylePr>
</w:style>";

        private static XElement Package(string cellsXml, string styles = Styles, string theme = Theme, string ns = W)
        {
            string xml = @"<pkg:package xmlns:pkg=""http://schemas.microsoft.com/office/2006/xmlPackage"">"
                + @"<pkg:part pkg:name=""/word/document.xml""><pkg:xmlData><w:document xmlns:w=""" + ns + @"""><w:body><w:tbl><w:tblPr><w:tblStyle w:val=""Grille""/><w:tblLook w:val=""04A0""/></w:tblPr>"
                + cellsXml + "</w:tbl><w:p/></w:body></w:document></pkg:xmlData></pkg:part>"
                + @"<pkg:part pkg:name=""/word/styles.xml""><pkg:xmlData><w:styles xmlns:w=""" + ns + @""">" + styles + "</w:styles></pkg:xmlData></pkg:part>"
                + (theme == null ? string.Empty : @"<pkg:part pkg:name=""/word/theme/theme1.xml""><pkg:xmlData>" + theme + "</pkg:xmlData></pkg:part>")
                + "</pkg:package>";
            return XElement.Parse(xml);
        }

        private static TableLayout ParseCells(string rowsXml, Func<bool, string> themeFont = null, string styles = Styles, string theme = Theme, string ns = W)
        {
            var root = Package(rowsXml, styles, theme, ns);
            var tbl = root.Descendants().First(e => e.Name.LocalName == "tbl");
            return WordXmlTableParser.Parse(root, tbl, themeFont);
        }

        private static string Row(params string[] cells)
        {
            return "<w:tr>" + string.Concat(cells.Select(c => "<w:tc>" + c + "</w:tc>")) + "</w:tr>";
        }

        [Fact]
        public void StylePrecedence_TableStyleParagraphStyleCharacterStyleAndDirect()
        {
            var layout = ParseCells(
                Row("<w:p><w:r><w:t>En-tête</w:t></w:r></w:p>",
                    @"<w:p><w:r><w:rPr><w:rStyle w:val=""Strong""/></w:rPr><w:t>Fort</w:t></w:r></w:p>")
                + Row(@"<w:p><w:pPr><w:pStyle w:val=""Titre""/></w:pPr><w:r><w:rPr><w:b w:val=""0""/></w:rPr><w:t>Maigre</w:t></w:r><w:r><w:t>Titre</w:t></w:r></w:p>",
                      @"<w:p><w:pPr><w:pStyle w:val=""Titre""/></w:pPr><w:r><w:rPr><w:rStyle w:val=""Emph""/></w:rPr><w:t>Emphase</w:t></w:r></w:p>"));

            // Style de tableau (ligne d'en-tête) : gras blanc, centré.
            var header = Runs(At(layout, 0, 0)).Single();
            Assert.True(header.Format.Bold);
            Assert.Equal(Hex("FFFFFF"), header.Format.Color);
            Assert.Equal("Calibri", header.Format.FontName);
            Assert.Equal(11, header.Format.Size);
            Assert.Equal(1, At(layout, 0, 0).Content.Alignment);

            // Propriété « bascule » : gras du style de tableau inversé par le style de caractère gras.
            Assert.False(Runs(At(layout, 0, 1)).Single().Format.Bold);

            // Style de paragraphe « Titre » : police des titres, 14 pt, gras, aligné à droite ; la mise en forme directe l'emporte.
            var maigre = RunWith(At(layout, 1, 0), "Maigre");
            Assert.False(maigre.Format.Bold);
            Assert.Equal("Cambria", maigre.Format.FontName);
            Assert.Equal(14, maigre.Format.Size);
            Assert.True(RunWith(At(layout, 1, 0), "Titre").Format.Bold);
            Assert.Equal(2, At(layout, 1, 0).Content.Alignment);

            // Style de caractère hérité (Emph basé sur Strong) dans un paragraphe gras : gras inversé, italique.
            var emphase = Runs(At(layout, 1, 1)).Single();
            Assert.False(emphase.Format.Bold);
            Assert.True(emphase.Format.Italic);
        }

        [Fact]
        public void Text_FieldsRevisionsHiddenTextAndSpecialCharacters()
        {
            var layout = ParseCells(Row(
                @"<w:p><w:r><w:t xml:space=""preserve"">Tableau </w:t></w:r><w:r><w:fldChar w:fldCharType=""begin""/></w:r><w:r><w:instrText> SEQ Tableau \* ARABIC </w:instrText></w:r>"
                + @"<w:r><w:fldChar w:fldCharType=""separate""/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType=""end""/></w:r>"
                + @"<w:fldSimple w:instr="" PAGE ""><w:r><w:t>12</w:t></w:r></w:fldSimple></w:p>",
                @"<w:p><w:r><w:t>Texte</w:t></w:r><w:del w:id=""1"" w:author=""A""><w:r><w:delText>supprimé</w:delText></w:r></w:del><w:ins w:id=""2"" w:author=""A""><w:r><w:t>ajouté</w:t></w:r></w:ins>"
                + @"<w:r><w:rPr><w:vanish/></w:rPr><w:t>caché</w:t></w:r><w:r><w:rPr><w:caps/></w:rPr><w:t>maj</w:t></w:r></w:p>",
                @"<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:noBreakHyphen/><w:t>d</w:t><w:softHyphen/><w:t>e</w:t></w:r><w:r><w:sym w:font=""Wingdings"" w:char=""F0FC""/></w:r></w:p>"
                + @"<w:p><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>contrôle</w:t></w:r></w:sdtContent></w:sdt><w:hyperlink r:id=""rId1"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><w:r><w:rPr><w:rStyle w:val=""Hyperlink""/></w:rPr><w:t>lien</w:t></w:r></w:hyperlink></w:p>"));

            Assert.Equal("Tableau 312", Text(At(layout, 0, 0)));
            Assert.Equal("TexteajoutéMAJ", Text(At(layout, 0, 1)));
            Assert.Equal("a b\nc-deü\ncontrôlelien", Text(At(layout, 0, 2)));

            var symbol = RunWith(At(layout, 0, 2), "ü");
            Assert.Equal("Wingdings", symbol.Format.FontName);
            var link = RunWith(At(layout, 0, 2), "lien");
            Assert.Equal(UnderlineKind.Single, link.Format.Underline);
            Assert.Equal(Hex("0000FF"), link.Format.Color);

            // Texte de vérification : ni code de champ, ni suppression, ni texte masqué ; « ( » pour un symbole.
            string check = CellTextSanitizer.Comparable(layout.XmlText);
            Assert.DoesNotContain("SEQ", check);
            Assert.DoesNotContain("supprimé", check);
            Assert.DoesNotContain("caché", check);
            Assert.Contains("Tableau312", check);
            Assert.Contains("c-de(", check);
        }

        [Fact]
        public void Colors_ValueThemeTintShadeAutoAndHighlight()
        {
            var layout = ParseCells(Row(
                @"<w:p><w:r><w:rPr><w:color w:themeColor=""accent1"" w:themeShade=""BF""/></w:rPr><w:t>ombré</w:t></w:r>"
                + @"<w:r><w:rPr><w:color w:themeColor=""accent1"" w:themeTint=""99""/></w:rPr><w:t>clair</w:t></w:r>"
                + @"<w:r><w:rPr><w:color w:val=""auto""/></w:rPr><w:t>auto</w:t></w:r>"
                + @"<w:r><w:rPr><w:color w:val=""00B050"" w:themeColor=""accent1""/><w:highlight w:val=""cyan""/></w:rPr><w:t>vert</w:t></w:r></w:p>"));
            var cell = At(layout, 0, 0);
            Assert.Equal(Rgb.Black.Blend(Hex("4F81BD"), 0xBF / 255.0), RunWith(cell, "ombré").Format.Color);
            Assert.Equal(Hex("4F81BD").Blend(Rgb.White, 1 - 0x99 / 255.0), RunWith(cell, "clair").Format.Color);
            // Ligne d'en-tête du style (blanc) remplacée par la couleur automatique explicite.
            Assert.Null(RunWith(cell, "auto").Format.Color);
            var green = RunWith(cell, "vert");
            Assert.Equal(Hex("00B050"), green.Format.Color);
            Assert.Equal(Hex("00FFFF"), green.Highlight);
        }

        [Fact]
        public void ThemeFontsFromWordWhenTheXmlHasNoTheme()
        {
            var asked = new List<bool>();
            var layout = ParseCells(Row("<w:p><w:r><w:t>x</w:t></w:r></w:p>"), major => { asked.Add(major); return major ? "Titres" : "Aptos"; }, theme: null);
            Assert.Equal("Aptos", Runs(At(layout, 0, 0)).Single().Format.FontName);
            Assert.Contains(false, asked);
        }

        [Fact]
        public void DocumentWithoutStyles_UsesWordDefaults()
        {
            var layout = ParseCells(Row("<w:p><w:r><w:t>x</w:t></w:r></w:p>"), styles: string.Empty, theme: null);
            var run = Runs(At(layout, 0, 0)).Single();
            Assert.Equal("Times New Roman", run.Format.FontName);
            Assert.Equal(10, run.Format.Size);
        }

        [Fact]
        public void Word2003Xml_ContentIsLeftToWord()
        {
            var layout = ParseCells(Row("<w:p><w:r><w:t>x</w:t></w:r></w:p>"), ns: "http://schemas.microsoft.com/office/word/2003/wordml");
            Assert.NotNull(layout);
            Assert.False(layout.HasContent);
            Assert.Null(At(layout, 0, 0).Content);
        }

        [Fact]
        public void ComparableText_IgnoresLayoutFieldCodesAndMarkers()
        {
            Assert.Equal("Tableau3", CellTextSanitizer.Comparable("Tableau \u0013 SEQ Tableau \u00143\u0015\r\a"));
            Assert.Equal("a-b", CellTextSanitizer.Comparable("a\u001Eb\u001F\t\u000B "));
            Assert.Equal(string.Empty, CellTextSanitizer.Comparable(null));
            // Pas de limite de longueur (contrairement au texte d'une cellule Excel).
            Assert.Equal(40000, CellTextSanitizer.Comparable(new string('x', 40000)).Length);
        }

        // ------------------------------------------------------------------ suivi des modifications

        private static bool HasVariant(TableLayout layout, string comparable)
        {
            return layout.XmlTextVariants.Any(v => CellTextSanitizer.Comparable(v) == comparable);
        }

        [Fact]
        public void TrackedDeletions_AreNeverExported_EvenIfNotAccepted()
        {
            var layout = ParseCells(
                Row("<w:p><w:r><w:t>En-tête</w:t></w:r></w:p>")
                + Row(@"<w:p><w:r><w:t xml:space=""preserve"">Prix </w:t></w:r><w:del w:id=""1"" w:author=""A""><w:r><w:delText>120</w:delText></w:r></w:del>"
                    + @"<w:ins w:id=""2"" w:author=""A""><w:r><w:t>150</w:t></w:r></w:ins>"
                    + @"<w:moveFrom w:id=""3"" w:author=""A""><w:r><w:t>déplacé</w:t></w:r></w:moveFrom></w:p>")
                // Marque de paragraphe supprimée : les deux paragraphes n'en font plus qu'un.
                + Row(@"<w:p><w:pPr><w:rPr><w:del w:id=""4"" w:author=""A""/></w:rPr></w:pPr><w:r><w:t>ab</w:t></w:r></w:p><w:p><w:r><w:t>cd</w:t></w:r></w:p>"));

            Assert.Equal("Prix 150", Text(At(layout, 1, 0)));
            Assert.Equal("abcd", Text(At(layout, 2, 0)));

            // Vérification avec le texte de Word : acceptée qu'il contienne le texte supprimé ou non.
            Assert.True(HasVariant(layout, "En-têtePrix150abcd"));
            Assert.True(HasVariant(layout, "En-têtePrix120150déplacéabcd"));
            Assert.Equal("En-têtePrix150abcd", CellTextSanitizer.Comparable(layout.XmlText));
        }

        [Fact]
        public void TrackedDeletedRows_AreRemovedFromTheTable()
        {
            const string deletedRow = @"<w:tr><w:trPr><w:del w:id=""9"" w:author=""A""/></w:trPr><w:tc><w:p><w:del w:id=""10"" w:author=""A""><w:r><w:delText>Supprimée</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p><w:del w:id=""11"" w:author=""A""><w:r><w:delText>0</w:delText></w:r></w:del></w:p></w:tc></w:tr>";
            var layout = ParseCells(
                Row("<w:p><w:r><w:t>Région</w:t></w:r></w:p>", "<w:p><w:r><w:t>Valeur</w:t></w:r></w:p>")
                + deletedRow
                + Row("<w:p><w:r><w:t>Nord</w:t></w:r></w:p>", "<w:p><w:r><w:t>12</w:t></w:r></w:p>"));

            Assert.Equal(2, layout.RowCount);
            Assert.Equal(3, layout.SourceRowCount);
            Assert.Contains(1, layout.DeletedRows);
            Assert.Equal("Nord", Text(At(layout, 1, 0)));
            Assert.Equal(2, At(layout, 1, 0).SourceRow); // rang dans Word, pour la lecture de secours
            Assert.DoesNotContain(layout.Cells, c => Text(c).Contains("Supprimée"));
            Assert.True(HasVariant(layout, "RégionValeurSupprimée0Nord12"));
            Assert.True(HasVariant(layout, "RégionValeurNord12"));
        }

        [Fact]
        public void TableWhoseRowsAreAllDeleted_IsEmpty()
        {
            var layout = ParseCells(@"<w:tr><w:trPr><w:del w:id=""1"" w:author=""A""/></w:trPr><w:tc><w:p><w:del w:id=""2"" w:author=""A""><w:r><w:delText>x</w:delText></w:r></w:del></w:p></w:tc></w:tr>");
            Assert.Equal(0, layout.RowCount);
            Assert.Empty(layout.Cells);
        }

        // ------------------------------------------------------------------ renvois et appels de notes

        private const string NoteStyles = Styles + @"
<w:style w:type=""character"" w:styleId=""Appelnotedebasdep""><w:name w:val=""footnote reference""/><w:rPr><w:vertAlign w:val=""superscript""/></w:rPr></w:style>
<w:style w:type=""character"" w:styleId=""AppelPerso""><w:name w:val=""Appel perso""/><w:basedOn w:val=""Appelnotedebasdep""/></w:style>
<w:style w:type=""character"" w:styleId=""Appelnotedefin""><w:name w:val=""endnote reference""/><w:rPr><w:vertAlign w:val=""superscript""/></w:rPr></w:style>";

        [Fact]
        public void NoteReferences_AreNotExported()
        {
            var layout = ParseCells(
                Row("<w:p><w:r><w:t>En-tête</w:t></w:r></w:p>")
                // Renvoi (Insertion › Renvoi › Note de bas de page) : champ NOTEREF dont le résultat est un chiffre.
                + Row(@"<w:p><w:r><w:t>12,5</w:t></w:r><w:r><w:fldChar w:fldCharType=""begin""/></w:r><w:r><w:instrText xml:space=""preserve""> NOTEREF _Ref4521 \h </w:instrText></w:r>"
                    + @"<w:r><w:fldChar w:fldCharType=""separate""/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType=""end""/></w:r></w:p>")
                + Row(@"<w:p><w:r><w:t>48</w:t></w:r><w:fldSimple w:instr="" NOTEREF _Ref4522 \f \h ""><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p>")
                // Appels de note : marque automatique (sans texte) ou texte au style « Appel de note ».
                + Row(@"<w:p><w:r><w:t>7</w:t></w:r><w:r><w:rPr><w:rStyle w:val=""Appelnotedebasdep""/></w:rPr><w:footnoteReference w:id=""2""/></w:r>"
                    + @"<w:r><w:rPr><w:rStyle w:val=""AppelPerso""/></w:rPr><w:t>5</w:t></w:r><w:r><w:rPr><w:rStyle w:val=""Appelnotedefin""/></w:rPr><w:t>i</w:t></w:r></w:p>")
                // Un exposant ordinaire (m²) reste, en exposant.
                + Row(@"<w:p><w:r><w:t>m</w:t></w:r><w:r><w:rPr><w:vertAlign w:val=""superscript""/></w:rPr><w:t>2</w:t></w:r></w:p>"),
                styles: NoteStyles);

            Assert.Equal("12,5", Text(At(layout, 1, 0)));
            Assert.Single(Runs(At(layout, 1, 0))); // valeur de mise en forme uniforme : convertible en nombre
            Assert.Equal("48", Text(At(layout, 2, 0)));
            Assert.Equal("7", Text(At(layout, 3, 0)));
            Assert.Equal("m2", Text(At(layout, 4, 0)));
            Assert.Equal(VerticalPosition.Superscript, Runs(At(layout, 4, 0)).Last().Format.Position);

            // Le texte de Word contient le résultat des renvois : la vérification en tient compte.
            Assert.Equal("En-tête12,5348475im2", CellTextSanitizer.Comparable(layout.XmlText));
        }
    }

    public class WordRevisionsTests
    {
        [Fact]
        public void Intervals_NormalizeSubtractCover()
        {
            var merged = WordTableToExcel.Word.WordRevisions.Normalize(new List<int[]> { new[] { 10, 12 }, new[] { 3, 5 }, new[] { 4, 8 }, new[] { 20, 20 } });
            Assert.Equal(new[] { "3-8", "10-12" }, merged.Select(i => i[0] + "-" + i[1]));

            var pieces = WordTableToExcel.Word.WordRevisions.Subtract(0, 15, merged);
            Assert.Equal(new[] { "0-3", "8-10", "12-15" }, pieces.Select(i => i[0] + "-" + i[1]));
            Assert.Equal(new[] { "0-15" }, WordTableToExcel.Word.WordRevisions.Subtract(0, 15, null).Select(i => i[0] + "-" + i[1]));
            Assert.Empty(WordTableToExcel.Word.WordRevisions.Subtract(4, 7, merged));

            Assert.True(WordTableToExcel.Word.WordRevisions.Covers(merged, 3, 8));
            Assert.False(WordTableToExcel.Word.WordRevisions.Covers(merged, 3, 11));
            Assert.False(WordTableToExcel.Word.WordRevisions.Covers(merged, 5, 5));
        }
    }
}
