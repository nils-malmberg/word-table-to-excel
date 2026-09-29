using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordTableToExcel.AddIn;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;
using WordTableToExcel.Word;

namespace WordTableToExcel.Import
{
    /// <summary>Feuille proposée à l'import (résultat de l'analyse préalable).</summary>
    internal sealed class ImportCandidate
    {
        public XlsxSheetInfo Info;
        public XlsxSheet Sheet;
        /// <summary>Plage du tableau (après la légende éventuelle) ; null si la feuille est vide ou illisible.</summary>
        public CellRange? Range;
        /// <summary>Plage proposée à l'origine (pour savoir si l'utilisateur l'a modifiée).</summary>
        public CellRange? DetectedRange;
        /// <summary>Légende trouvée en tête de feuille (texte complet).</summary>
        public string DetectedCaption;
        /// <summary>Texte de légende à placer après « Tableau N » (modifiable dans la boîte de dialogue).</summary>
        public string CaptionTitle;
        /// <summary>Raison qui empêche l'import (feuille vide, graphique, illisible…), ou null.</summary>
        public string Problem;
        public bool Selected;

        public bool CanImport
        {
            get { return Sheet != null && Range.HasValue; }
        }
    }

    /// <summary>
    /// Déroulement d'un import : choix du classeur, analyse des feuilles, choix de l'utilisateur, conversion,
    /// insertion dans le document (annulable en une fois par Ctrl+Z), compte rendu.
    /// </summary>
    internal sealed class ImportService
    {
        private const int WdNoProtection = -1;
        private const int MsoLanguageIdUi = 2;

        private readonly dynamic _application;

        public ImportService(object application)
        {
            _application = application;
        }

        public void Run()
        {
            IWin32Window owner = WindowOwner.FromWord((object)_application);
            object documentObject = GetEditableDocument(owner);
            if (documentObject == null) return;

            var settings = Settings.Load();
            string path = AskSourcePath(owner, settings);
            if (path == null) return;
            settings.ImportLastFolder = Path.GetDirectoryName(path);
            settings.Save();

            XlsxWorkbook workbook = LoadWorkbook(owner, path);
            if (workbook == null) return;

            var culture = CultureInfo.CurrentCulture;
            var format = new ExcelFormatSettings(culture, OfficeLanguage(), workbook.Date1904, workbook.Styles.Colors);
            int wordVersion = Connect.WordMajorVersion((object)_application);
            var inserter = new WordTableInserter((object)_application, documentObject, wordVersion);
            var matcher = new CaptionMatcher(new List<string>(settings.ExtraLabels()) { inserter.CaptionLabel });

            List<ImportCandidate> candidates;
            using (new WaitCursor())
            {
                candidates = Analyze(workbook, format, matcher);
            }
            if (!candidates.Any(c => c.CanImport))
            {
                var details = string.Join("\n", candidates.Select(c => "• " + c.Info.Name + " : " + (c.Problem ?? "vide")).ToArray());
                Messages.Info(owner, "Aucune feuille de ce classeur ne contient de tableau à importer.\n\n" + details);
                return;
            }

            bool below = settings.ImportCaptionBelow ?? DocumentUsesCaptionsBelow(documentObject, matcher);
            ImportChoices choices;
            using (var dialog = new ImportDialog(Path.GetFileName(path), candidates, settings, below))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK) return;
                choices = dialog.Choices;
            }
            settings.ImportAddCaption = choices.AddCaption;
            settings.ImportCaptionBelow = choices.CaptionBelow;
            settings.ImportFitToPage = choices.FitToPage;
            settings.ImportSkipHidden = choices.SkipHidden;
            settings.ImportGridlines = choices.AddGridlines;
            settings.Save();

            var report = new ImportReport();
            var jobs = Convert(workbook, format, matcher, inserter, candidates.Where(c => c.Selected && c.CanImport).ToList(), choices, report);
            if (jobs.Count == 0)
            {
                ShowReport(owner, report);
                return;
            }

            long cells = jobs.Sum(j => (long)j.Table.RowCount * j.Table.ColumnCount);
            if (cells > 5000 && !Messages.Ask(owner, "L'import va créer " + WordTableInserter.Describe(jobs.Count) + " totalisant "
                + cells.ToString("N0", CultureInfo.CurrentCulture) + " cellules : l'opération peut prendre du temps.\n\nContinuer ?"))
            {
                return;
            }

            try
            {
                ProgressDialog.Run(owner, "Import des tableaux Excel", progress => Insert(inserter, jobs, choices, progress, report));
            }
            catch (ExportFailedException ex)
            {
                if (!(ex.InnerException is OperationCanceledException))
                {
                    Log.Error("Insertion des tableaux", ex.InnerException ?? ex);
                    report.Errors.Add("L'insertion s'est interrompue : " + (ex.InnerException ?? ex).Message);
                }
            }
            ShowReport(owner, report);
        }

        // ------------------------------------------------------------------ document

        private object GetEditableDocument(IWin32Window owner)
        {
            object document = null;
            try
            {
                if (WordCom.AsInt(_application.Documents.Count) > 0) document = _application.ActiveDocument;
            }
            catch (Exception ex)
            {
                Log.Info("Pas de document actif : " + ex.Message);
            }
            if (document == null)
            {
                bool protectedView = false;
                try
                {
                    protectedView = _application.ActiveProtectedViewWindow != null;
                }
                catch (Exception)
                {
                    // Word 2007 : pas de mode protégé.
                }
                Messages.Info(owner, protectedView
                    ? "Le document est ouvert en mode protégé.\n\nCliquez sur « Activer la modification », puis relancez l'import."
                    : "Aucun document n'est ouvert.\n\nOuvrez ou créez un document Word, placez le curseur à l'endroit voulu, puis relancez l'import.");
                return null;
            }
            try
            {
                if (WordCom.AsInt(((dynamic)document).ProtectionType) != WdNoProtection)
                {
                    Messages.Warning(owner, "Le document est protégé contre les modifications (Révision › Restreindre la modification).\n\n"
                        + "Retirez la protection, puis relancez l'import.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Protection du document inconnue : " + ex.Message);
            }
            return document;
        }

        private CultureInfo OfficeLanguage()
        {
            try
            {
                int lcid = WordCom.AsInt(_application.LanguageSettings.LanguageID(MsoLanguageIdUi));
                if (lcid > 0) return CultureInfo.GetCultureInfo(lcid);
            }
            catch (Exception ex)
            {
                Log.Info("Langue d'Office inconnue : " + ex.Message);
            }
            return CultureInfo.CurrentUICulture;
        }

        /// <summary>Position des légendes déjà présentes dans le document (sous les tableaux ou au-dessus).</summary>
        private static bool DocumentUsesCaptionsBelow(object documentObject, CaptionMatcher matcher)
        {
            try
            {
                dynamic document = documentObject;
                var scanner = new WordCaptionScanner(documentObject, matcher, Log.Info);
                var contexts = new List<TableCaptionContext>();
                int index = 0;
                foreach (dynamic table in document.Tables)
                {
                    index++;
                    contexts.Add(scanner.Scan((object)table, index));
                    if (index >= 40) break;
                }
                return CaptionAssigner.DetectConvention(contexts) == CaptionPosition.Below;
            }
            catch (Exception ex)
            {
                Log.Info("Convention de légende inconnue : " + ex.Message);
                return false;
            }
        }

        // ------------------------------------------------------------------ classeur

        private static string AskSourcePath(IWin32Window owner, Settings settings)
        {
            string folder = !string.IsNullOrEmpty(settings.ImportLastFolder) && Directory.Exists(settings.ImportLastFolder)
                ? settings.ImportLastFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            using (var dialog = new OpenFileDialog
            {
                Title = "Importer des tableaux depuis un classeur Excel",
                Filter = "Classeurs Excel (*.xlsx;*.xlsm;*.xltx;*.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm"
                       + "|Autres classeurs, convertis par Excel (*.xls;*.xlsb;*.ods;*.csv)|*.xls;*.xlsb;*.ods;*.csv"
                       + "|Tous les fichiers (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false,
                InitialDirectory = folder
            })
            {
                return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
            }
        }

        private static XlsxWorkbook LoadWorkbook(IWin32Window owner, string path)
        {
            string converted = null;
            try
            {
                using (new WaitCursor())
                {
                    if (!ExcelFileConverter.IsOpenXml(path) && !LooksLikeZip(path))
                    {
                        converted = ExcelFileConverter.ConvertToXlsx(path);
                        return XlsxWorkbook.Load(converted);
                    }
                    try
                    {
                        return XlsxWorkbook.Load(path);
                    }
                    catch (ExcelImportException ex)
                    {
                        if (!ex.ConvertibleByExcel || !ExcelFileConverter.IsExcelInstalled()) throw;
                        Log.Info("Conversion par Excel : " + ex.Message);
                        converted = ExcelFileConverter.ConvertToXlsx(path);
                        return XlsxWorkbook.Load(converted);
                    }
                }
            }
            catch (ExcelImportException ex)
            {
                Log.Error("Lecture du classeur " + path, ex);
                Messages.Warning(owner, "Impossible de lire le classeur :\n" + path + "\n\n" + ex.Message);
                return null;
            }
            finally
            {
                if (converted != null)
                {
                    try
                    {
                        File.Delete(converted);
                    }
                    catch (Exception)
                    {
                        // Fichier temporaire : sans conséquence.
                    }
                }
            }
        }

        private static bool LooksLikeZip(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    return fs.ReadByte() == 0x50 && fs.ReadByte() == 0x4B;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static List<ImportCandidate> Analyze(XlsxWorkbook workbook, ExcelFormatSettings format, CaptionMatcher matcher)
        {
            var converter = new SheetConverter(workbook, format, new ExcelImportOptions(), matcher);
            var result = new List<ImportCandidate>();
            foreach (var info in workbook.Sheets)
            {
                var candidate = new ImportCandidate { Info = info };
                result.Add(candidate);
                if (info.Kind != XlsxSheetKind.Worksheet || info.PartName == null)
                {
                    candidate.Problem = info.Kind == XlsxSheetKind.Chartsheet ? "Feuille graphique" : "Pas une feuille de calcul";
                    continue;
                }
                try
                {
                    candidate.Sheet = workbook.ReadSheet(info);
                    var range = converter.DetectRange(candidate.Sheet);
                    if (!range.HasValue)
                    {
                        candidate.Problem = "Feuille vide";
                        continue;
                    }
                    string caption;
                    CellRange tableRange;
                    if (converter.TryDetectCaption(candidate.Sheet, range.Value, out caption, out tableRange))
                    {
                        candidate.DetectedCaption = caption;
                        candidate.CaptionTitle = SheetConverter.CaptionTitle(caption);
                        range = tableRange;
                    }
                    candidate.Range = range;
                    candidate.DetectedRange = range;
                    candidate.Selected = info.IsVisible;
                }
                catch (ExcelImportException ex)
                {
                    Log.Error("Lecture de la feuille " + info.Name, ex);
                    candidate.Sheet = null;
                    candidate.Problem = ex.Message;
                }
            }
            return result;
        }

        // ------------------------------------------------------------------ conversion et insertion

        private sealed class ImportJob
        {
            public ImportCandidate Candidate;
            public TableModel Table;
            public List<string> Warnings;
        }

        private sealed class ImportReport
        {
            public readonly List<string> Inserted = new List<string>();
            public readonly List<string> Errors = new List<string>();
            public readonly List<string> Warnings = new List<string>();
            public bool Cancelled;
        }

        private List<ImportJob> Convert(XlsxWorkbook workbook, ExcelFormatSettings format, CaptionMatcher matcher, WordTableInserter inserter,
            List<ImportCandidate> selected, ImportChoices choices, ImportReport report)
        {
            var jobs = new List<ImportJob>();
            using (new WaitCursor())
            using (var measurer = new TextMeasurer())
            {
                var options = new ExcelImportOptions
                {
                    SkipHiddenRowsAndColumns = choices.SkipHidden,
                    AddGridlines = choices.AddGridlines,
                    DetectCaption = false,
                    AvailableWidthPt = choices.FitToPage ? inserter.AvailableWidth() : 0,
                    MeasureText = measurer.Measure
                };
                var converter = new SheetConverter(workbook, format, options, matcher);
                foreach (var candidate in selected)
                {
                    SheetImport result;
                    try
                    {
                        result = converter.Convert(candidate.Sheet, candidate.Range);
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Conversion de la feuille " + candidate.Info.Name, ex);
                        report.Errors.Add("« " + candidate.Info.Name + " » : conversion impossible (" + ex.Message + ").");
                        continue;
                    }
                    if (!result.Succeeded)
                    {
                        report.Errors.Add("« " + candidate.Info.Name + " » : " + result.Error);
                        continue;
                    }
                    jobs.Add(new ImportJob { Candidate = candidate, Table = result.Table, Warnings = result.Warnings });
                }
            }
            return jobs;
        }

        private void Insert(WordTableInserter inserter, List<ImportJob> jobs, ImportChoices choices, IExportProgress progress, ImportReport report)
        {
            bool undoRecord = false;
            bool screenUpdating = true;
            try
            {
                try
                {
                    _application.UndoRecord.StartCustomRecord("Importer des tableaux Excel");
                    undoRecord = true;
                }
                catch (Exception)
                {
                    // Word 2007 : pas d'enregistrement d'annulation groupé.
                }
                try
                {
                    screenUpdating = WordCom.IsTrue(_application.ScreenUpdating);
                    _application.ScreenUpdating = false;
                }
                catch (Exception)
                {
                    // Sans conséquence.
                }

                for (int i = 0; i < jobs.Count; i++)
                {
                    if (progress.IsCancellationRequested)
                    {
                        report.Cancelled = true;
                        break;
                    }
                    var job = jobs[i];
                    progress.Report("Insertion de « " + job.Candidate.Info.Name + " » (" + (i + 1) + " sur " + jobs.Count + ")…", (double)i / jobs.Count);
                    CaptionRequest caption = choices.AddCaption ? new CaptionRequest { Title = job.Candidate.CaptionTitle, Below = choices.CaptionBelow } : null;
                    inserter.Insert(job.Table, caption);
                    report.Inserted.Add(job.Candidate.Info.Name);
                    foreach (var w in job.Warnings) report.Warnings.Add("« " + job.Candidate.Info.Name + " » : " + w);
                }
                if (choices.AddCaption && report.Inserted.Count > 0)
                {
                    progress.Report("Mise à jour de la numérotation des légendes…", 1);
                    inserter.UpdateCaptionNumbers();
                }
            }
            finally
            {
                try
                {
                    _application.ScreenUpdating = screenUpdating;
                }
                catch (Exception)
                {
                    // Sans conséquence.
                }
                if (undoRecord)
                {
                    try
                    {
                        _application.UndoRecord.EndCustomRecord();
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Fin de l'enregistrement d'annulation", ex);
                    }
                }
                inserter.SelectEnd();
            }
        }

        private static void ShowReport(IWin32Window owner, ImportReport report)
        {
            var sb = new StringBuilder();
            if (report.Inserted.Count > 0)
            {
                sb.Append(WordTableInserter.Describe(report.Inserted.Count)).Append(report.Inserted.Count > 1 ? " insérés" : " inséré")
                  .Append(" dans le document.");
                sb.AppendLine().Append("Pour tout annuler : Ctrl+Z (Annuler).");
            }
            else
            {
                sb.Append("Aucun tableau n'a été inséré.");
            }
            if (report.Cancelled) sb.AppendLine().AppendLine().Append("Import interrompu à votre demande.");
            AppendList(sb, "Problèmes :", report.Errors);
            AppendList(sb, "Remarques :", report.Warnings);
            string text = sb.ToString();
            if (report.Errors.Count > 0 && report.Inserted.Count == 0) Messages.Warning(owner, text);
            else Messages.Info(owner, text);
        }

        private static void AppendList(StringBuilder sb, string title, List<string> items)
        {
            if (items.Count == 0) return;
            sb.AppendLine().AppendLine().AppendLine(title);
            const int max = 12;
            foreach (var item in items.Take(max)) sb.AppendLine("• " + item);
            if (items.Count > max) sb.AppendLine("• … (" + (items.Count - max) + " de plus, voir le journal)");
            foreach (var item in items.Skip(max)) Log.Info(item);
        }

        /// <summary>Curseur d'attente pendant une opération courte.</summary>
        private sealed class WaitCursor : IDisposable
        {
            private readonly Cursor _previous;

            public WaitCursor()
            {
                _previous = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
            }

            public void Dispose()
            {
                Cursor.Current = _previous;
            }
        }
    }
}
