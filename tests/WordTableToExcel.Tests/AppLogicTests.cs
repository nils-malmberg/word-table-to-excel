using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WordTableToExcel.App;
using Xunit;

namespace WordTableToExcel.Tests
{
    /// <summary>Logique de l'application autonome indépendante de Word et de l'interface.</summary>
    public class AppLogicTests
    {
        // ------------------------------------------------------------------ liste des documents

        private static DocumentEntry Open(string fullName, int pid = 10, bool saved = true)
        {
            return new DocumentEntry
            {
                State = DocumentState.Open,
                ProcessId = pid,
                Name = fullName.Substring(fullName.LastIndexOf('\\') + 1),
                FullName = fullName,
                Folder = saved ? fullName.Substring(0, Math.Max(0, fullName.LastIndexOf('\\'))) : string.Empty, // chemins Windows, tests exécutés aussi sous Linux
                TableCount = 2
            };
        }

        [Fact]
        public void FileAlreadyOpenInWordIsListedOnceAsOpenDocument()
        {
            var open = new[] { Open(@"C:\Docs\Rapport.docx") };
            var list = DocumentList.Merge(open, new[] { @"c:\docs\RAPPORT.docx", @"C:\Docs\Budget.docx" });
            Assert.Equal(2, list.Count);
            Assert.Equal(DocumentState.Open, list[0].State);
            Assert.Equal(DocumentState.File, list[1].State);
            Assert.EndsWith("Budget.docx", list[1].Name); // nom seul sous Windows (chemin complet sous Linux)
            Assert.Equal(-1, list[1].TableCount);
        }

        [Fact]
        public void UnsavedDocumentNeverMatchesAFile()
        {
            var open = new[] { Open("Document1", saved: false) };
            var list = DocumentList.Merge(open, new[] { "Document1" });
            Assert.Equal(2, list.Count);
        }

        [Fact]
        public void DuplicateFilesAreListedOnce()
        {
            var list = DocumentList.Merge(new DocumentEntry[0], new[] { @"C:\A\x.docx", @"C:\A\X.DOCX", @"C:\A\y.docx" });
            Assert.Equal(2, list.Count);
        }

        [Fact]
        public void SameDocumentInTwoWordInstancesIsKeptTwice()
        {
            var list = DocumentList.Merge(new[] { Open(@"C:\A\x.docx", 1), Open(@"C:\A\x.docx", 2), Open(@"C:\A\x.docx", 1) }, new string[0]);
            Assert.Equal(2, list.Count);
        }

        [Fact]
        public void KeysDistinguishOpenProtectedAndFileEntries()
        {
            var open = Open(@"C:\A\x.docx");
            var pv = Open(@"C:\A\x.docx");
            pv.State = DocumentState.ProtectedView;
            var file = DocumentEntry.ForFile(@"C:\A\x.docx");
            Assert.NotEqual(open.Key, pv.Key);
            Assert.NotEqual(open.Key, file.Key);
            Assert.Equal(open.Key, Open(@"c:\a\X.DOCX").Key);
        }

        [Fact]
        public void OneDriveAddressesAreComparedAsUrls()
        {
            Assert.True(DocumentList.SamePath("https://d.docs.live.net/abc/Documents/Rapport.docx", "https://d.docs.live.net/abc/Documents/rapport.docx"));
            Assert.False(DocumentList.SamePath("https://d.docs.live.net/abc/Rapport.docx", @"C:\Rapport.docx"));
            Assert.False(DocumentList.SamePath(null, "x"));
        }

        [Theory]
        [InlineData(@"C:\a\b.docx", true, false)]
        [InlineData(@"C:\a\b.DOC", true, false)]
        [InlineData(@"C:\a\b.rtf", true, false)]
        [InlineData(@"C:\a\b.xlsx", false, true)]
        [InlineData(@"C:\a\b.xls", false, true)]
        [InlineData(@"C:\a\b.csv", false, true)]
        [InlineData(@"C:\a\b.pdf", false, false)]
        [InlineData(@"C:\a\docx", false, false)]
        [InlineData("", false, false)]
        public void FileKinds(string path, bool word, bool excel)
        {
            Assert.Equal(word, DocumentList.IsWordFile(path));
            Assert.Equal(excel, DocumentList.IsExcelFile(path));
        }

        // ------------------------------------------------------------------ point d'insertion

        [Fact]
        public void DescribesInsertionAfterParagraph()
        {
            var info = new InsertionInfo { DocumentName = "Rapport.docx", Page = 3, ParagraphText = "Les ventes progressent.\r" };
            Assert.Equal("Insertion dans « Rapport.docx », page 3 : juste après le paragraphe « Les ventes progressent. ».", InsertionInfo.Describe(info, true));
        }

        [Fact]
        public void DescribesInsertionBeforeParagraphWhenCursorAtItsStart()
        {
            var info = new InsertionInfo { DocumentName = "R.docx", ParagraphText = "Titre\r", AtParagraphStart = true };
            Assert.Equal("Insertion dans « R.docx » : juste avant le paragraphe « Titre ».", InsertionInfo.Describe(info, true));
        }

        [Fact]
        public void DescribesEmptyLineTableSelectionAndHeader()
        {
            Assert.EndsWith("sur la ligne vide où se trouve le curseur.",
                InsertionInfo.Describe(new InsertionInfo { DocumentName = "R.docx", ParagraphText = "\r" }, true));
            Assert.EndsWith("juste après le tableau où se trouve le curseur.",
                InsertionInfo.Describe(new InsertionInfo { DocumentName = "R.docx", InTable = true }, true));

            string text = InsertionInfo.Describe(new InsertionInfo { DocumentName = "R.docx", ParagraphText = "Texte\r", AtParagraphStart = true, HasSelection = true, MainStory = false }, true);
            Assert.Contains("juste après le paragraphe « Texte ».", text); // une sélection est conservée : insertion après elle
            Assert.Contains("Le texte sélectionné dans Word est conservé.", text);
            Assert.Contains("en-tête, un pied de page", text);
        }

        [Fact]
        public void DescribesMissingWordOrDocumentAndProblems()
        {
            Assert.StartsWith("Word n'est pas ouvert.", InsertionInfo.Describe(null, false));
            Assert.StartsWith("Aucun document n'est ouvert dans Word.", InsertionInfo.Describe(null, true));
            var info = new InsertionInfo { DocumentName = "R.docx", Problem = "Mode protégé." };
            Assert.False(info.CanImport);
            Assert.Equal("Mode protégé.", InsertionInfo.Describe(info, true));
        }

        [Fact]
        public void ExcerptIsSingleLineAndShortened()
        {
            Assert.Equal("Titre du chapitre", InsertionInfo.Excerpt("  Titre\tdu\u000Bchapitre\r\a"));
            Assert.Equal(string.Empty, InsertionInfo.Excerpt("\r"));
            string longText = string.Join(" ", new[] { "Les", "résultats", "du", "troisième", "trimestre", "confirment", "la", "progression", "des", "ventes", "dans", "toutes", "les", "régions" });
            string excerpt = InsertionInfo.Excerpt(longText);
            Assert.EndsWith("\u2026", excerpt);
            Assert.True(excerpt.Length <= 60);
            Assert.StartsWith("Les résultats du troisième trimestre", excerpt);
            Assert.DoesNotContain(" \u2026", excerpt);
        }

        // ------------------------------------------------------------------ erreurs de communication avec Word

        [Fact]
        public void RejectedCallsAreRetriedUntilTimeout()
        {
            Assert.Equal(200, ComErrors.RetryDelay(ComErrors.ServerCallRetryLater, 0, 1500));
            Assert.Equal(200, ComErrors.RetryDelay(ComErrors.ServerCallRetryLater, 1499, 1500));
            Assert.Equal(-1, ComErrors.RetryDelay(ComErrors.ServerCallRetryLater, 1500, 1500));
            Assert.Equal(-1, ComErrors.RetryDelay(1, 0, 1500)); // SERVERCALL_REJECTED : abandon immédiat
        }

        [Fact]
        public void FriendlyMessagesForBusyOrClosedWord()
        {
            var busy = new InvalidOperationException("enveloppe", new COMException("rejeté", ComErrors.CallRejected));
            Assert.Equal(ComErrors.CallRejected, ComErrors.HResultOf(busy));
            Assert.Contains("Word est occupé", ComErrors.FriendlyMessage(busy));
            Assert.Contains("Word a été fermé", ComErrors.FriendlyMessage(new COMException("x", ComErrors.ServerUnavailable)));
            Assert.Contains("Word a été fermé", ComErrors.FriendlyMessage(new COMException("x", ComErrors.Disconnected)));
            Assert.Null(ComErrors.FriendlyMessage(new COMException("autre", unchecked((int)0x800A1066))));
            Assert.Null(ComErrors.FriendlyMessage(new InvalidOperationException("x")));
            Assert.Equal(0, ComErrors.HResultOf(null));
        }
    }
}
