using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Core.Xlsx;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;
using WordTableToExcel.Word;

namespace WordTableToExcel.Export
{
    /// <summary>
    /// Déroulement complet d'un export : détection des tableaux et légendes, choix de l'utilisateur,
    /// choix du fichier, lecture des tableaux, écriture du classeur, compte rendu.
    /// Le document Word n'est jamais modifié.
    /// </summary>
    internal sealed class ExportService
    {
        private readonly dynamic _application;

        public ExportService(object application)
        {
            _application = application;
        }

        public void Run()
        {
            IWin32Window owner = WindowOwner.FromWord((object)_application);
            object documentObject = GetActiveDocument();
            if (documentObject == null)
            {
                Messages.Info(owner, "Aucun document n'est ouvert.\n\nOuvrez un document Word contenant des tableaux, puis relancez l'export.\n"
                    + "(Si le document est en mode protégé, cliquez d'abord sur « Activer la modification ».)");
                return;
            }

            dynamic document = documentObject;
            using (new WordDocumentGuard(documentObject))
            {
                int tableCount = WordCom.AsInt(document.Tables.Count);
                if (tableCount == 0)
                {
                    Messages.Info(owner, "Ce document ne contient aucun tableau.");
                    return;
                }

                var settings = Settings.Load();
                CaptionPosition convention;
                List<TableEntry> tables;
                Cursor previousCursor = Cursor.Current;
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    tables = ScanTables(documentObject, tableCount, settings, out convention);
                }
                finally
                {
                    Cursor.Current = previousCursor;
                }

                string documentName = WordCom.AsString(document.Name);
                bool allTables, includeCaption, convertNumbers;
                using (var dialog = new ExportDialog(documentName, tables, convention, settings.AllTables, settings.IncludeCaptionRow, settings.ConvertNumbers))
                {
                    if (dialog.ShowDialog(owner) != DialogResult.OK) return;
                    allTables = dialog.AllTables;
                    includeCaption = dialog.IncludeCaptionRow;
                    convertNumbers = dialog.ConvertNumbers;
                }
                settings.AllTables = allTables;
                settings.IncludeCaptionRow = includeCaption;
                settings.ConvertNumbers = convertNumbers;

                var plan = ExportPlan.Build(tables, allTables);
                if (plan.Count == 0)
                {
                    Messages.Info(owner, "Aucun tableau ne correspond au choix effectué.");
                    return;
                }

                string path = AskTargetPath(owner, documentObject, settings);
                if (path == null) return;
                settings.LastFolder = Path.GetDirectoryName(path);
                settings.Save();

                var options = new XlsxExportOptions
                {
                    IncludeCaptionRow = includeCaption,
                    ConvertNumbers = convertNumbers,
                    NumberCulture = CultureInfo.CurrentCulture,
                    Title = Path.GetFileNameWithoutExtension(documentName)
                };

                var report = new ExportReport();
                try
                {
                    ProgressDialog.Run(owner, "Export des tableaux vers Excel", progress => Export(documentObject, plan, options, path, progress, report));
                }
                catch (ExportFailedException ex)
                {
                    if (ex.InnerException is OperationCanceledException)
                    {
                        Messages.Info(owner, "Export annulé. Aucun fichier n'a été créé ni modifié.");
                        return;
                    }
                    if (ex.InnerException is IOException || ex.InnerException is UnauthorizedAccessException)
                    {
                        Log.Error("Écriture du fichier impossible", ex.InnerException);
                        Messages.Warning(owner, "Impossible d'enregistrer le fichier :\n" + path + "\n\n" + ex.InnerException.Message
                            + "\n\nSi ce classeur est ouvert dans Excel, fermez-le puis recommencez.\nLe document Word n'a pas été modifié.");
                        return;
                    }
                    throw;
                }

                ShowReport(owner, path, report);
            }
        }

        private object GetActiveDocument()
        {
            try
            {
                if (WordCom.AsInt(_application.Documents.Count) > 0) return _application.ActiveDocument;
            }
            catch (Exception ex)
            {
                Log.Info("Pas de document actif : " + ex.Message);
            }
            try
            {
                // Word 2010+ : document ouvert en mode protégé (lecture seule).
                return _application.ActiveProtectedViewWindow.Document;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // ------------------------------------------------------------------ détection

        private static List<TableEntry> ScanTables(object documentObject, int tableCount, Settings settings, out CaptionPosition convention)
        {
            dynamic document = documentObject;
            var labels = new List<string>(settings.ExtraLabels());
            try
            {
                // Libellé « Tableau » de la langue d'interface de Word (wdCaptionTable = -2).
                labels.Add(WordCom.AsString(document.Application.CaptionLabels.Item(-2).Name));
            }
            catch (Exception)
            {
                // Collection absente ou libellé supprimé : les libellés intégrés suffisent.
            }

            var scanner = new WordCaptionScanner(documentObject, new CaptionMatcher(labels), Log.Info);
            var contexts = new List<TableCaptionContext>();
            int index = 0;
            foreach (dynamic table in document.Tables)
            {
                index++;
                contexts.Add(scanner.Scan((object)table, index));
            }
            if (index != tableCount) Log.Info("Nombre de tableaux : " + tableCount + " annoncés, " + index + " parcourus.");

            var assignments = CaptionAssigner.Assign(contexts, out convention);
            var entries = new List<TableEntry>();
            for (int i = 0; i < contexts.Count; i++)
            {
                entries.Add(new TableEntry
                {
                    Index = contexts[i].TableIndex,
                    Caption = assignments[i].Caption == null ? null : assignments[i].Caption.Text,
                    CaptionPosition = assignments[i].Position
                });
            }
            Log.Info(string.Format("{0} tableau(x), {1} avec légende, convention : {2}.", entries.Count, entries.Count(e => e.HasCaption), convention));
            return entries;
        }

        // ------------------------------------------------------------------ fichier cible

        private static string AskTargetPath(IWin32Window owner, object documentObject, Settings settings)
        {
            dynamic document = documentObject;
            string documentPath = string.Empty, documentFullName = string.Empty, documentName = "Document";
            try
            {
                documentPath = WordCom.AsString(document.Path);
                documentFullName = WordCom.AsString(document.FullName);
                documentName = Path.GetFileNameWithoutExtension(WordCom.AsString(document.Name));
            }
            catch (Exception)
            {
                // Document jamais enregistré : valeurs par défaut.
            }

            string initialFolder = !string.IsNullOrEmpty(documentPath) && Directory.Exists(documentPath) ? documentPath
                : !string.IsNullOrEmpty(settings.LastFolder) && Directory.Exists(settings.LastFolder) ? settings.LastFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using (var dialog = new SaveFileDialog
            {
                Title = "Enregistrer le classeur Excel",
                Filter = "Classeur Excel (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                OverwritePrompt = true,
                CheckPathExists = true,
                InitialDirectory = initialFolder,
                FileName = SafeFileName(documentName) + " - tableaux.xlsx"
            })
            {
                while (true)
                {
                    if (dialog.ShowDialog(owner) != DialogResult.OK) return null;
                    string path = dialog.FileName;
                    if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) path += ".xlsx";

                    if (!string.IsNullOrEmpty(documentFullName) && string.Equals(Path.GetFullPath(path), SafeFullPath(documentFullName), StringComparison.OrdinalIgnoreCase))
                    {
                        Messages.Warning(owner, "Ce fichier est le document Word lui-même. Choisissez un autre nom.");
                        continue;
                    }
                    if (!string.Equals(path, dialog.FileName, StringComparison.Ordinal) && File.Exists(path)
                        && !Messages.Ask(owner, "Le fichier existe déjà :\n" + path + "\n\nVoulez-vous le remplacer ?"))
                    {
                        continue;
                    }
                    return path;
                }
            }
        }

        private static string SafeFullPath(string path)
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch (Exception)
            {
                return path;
            }
        }

        private static string SafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in name ?? string.Empty) sb.Append(invalid.Contains(c) ? '_' : c);
            string result = sb.ToString().Trim();
            return result.Length == 0 ? "Document" : result;
        }

        // ------------------------------------------------------------------ export

        private static void Export(object documentObject, List<PlannedSheet> plan, XlsxExportOptions options, string path, IExportProgress progress, ExportReport report)
        {
            dynamic document = documentObject;
            var reader = new WordTableReader(documentObject, Log.Info);
            var writer = new XlsxWorkbookWriter(options);
            var stopwatch = Stopwatch.StartNew();

            for (int i = 0; i < plan.Count; i++)
            {
                var item = plan[i];
                if (progress.IsCancellationRequested) throw new OperationCanceledException();

                string label = string.Format(CultureInfo.CurrentCulture, "Lecture du tableau {0} ({1} sur {2}) : {3}", item.Table.Index, i + 1, plan.Count, item.SheetName);
                double baseFraction = (double)i / (plan.Count + 1);
                double step = 1.0 / (plan.Count + 1);
                progress.Report(label, baseFraction);

                TableModel model;
                try
                {
                    object table = document.Tables.Item(item.Table.Index);
                    model = reader.Read(table, item.Table.Index, (done, total) =>
                    {
                        if (progress.IsCancellationRequested) throw new OperationCanceledException();
                        progress.Report(label, baseFraction + step * done / Math.Max(1, total));
                    });
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Log.Error("Lecture du tableau " + item.Table.Index, ex);
                    report.Failures.Add(string.Format(CultureInfo.CurrentCulture, "Tableau {0} ({1}) non exporté : {2}", item.Table.Index, item.SheetName, ex.Message));
                    continue;
                }

                model.Caption = item.Table.Caption;
                model.CaptionPosition = item.Table.CaptionPosition;
                model.SheetName = item.SheetName;
                writer.AddTable(model);
                report.Exported++;
                foreach (var warning in model.Warnings)
                {
                    report.Warnings.Add(string.Format(CultureInfo.CurrentCulture, "Tableau {0} : {1}", item.Table.Index, warning));
                }
            }

            if (writer.SheetCount == 0)
            {
                throw new InvalidOperationException("Aucun tableau n'a pu être lu. " + string.Join(" ", report.Failures.ToArray()));
            }

            if (progress.IsCancellationRequested) throw new OperationCanceledException();
            progress.Report("Enregistrement du classeur Excel…", (double)plan.Count / (plan.Count + 1));
            SaveSafely(writer, path);
            progress.Report("Terminé.", 1);
            Log.Info(string.Format("Export de {0} tableau(x) vers {1} en {2} ms.", report.Exported, path, stopwatch.ElapsedMilliseconds));
        }

        /// <summary>
        /// Écrit d'abord un fichier temporaire, puis remplace la cible : un incident pendant
        /// l'écriture ne laisse jamais de classeur à moitié écrit.
        /// </summary>
        private static void SaveSafely(XlsxWorkbookWriter writer, string path)
        {
            string folder = Path.GetDirectoryName(Path.GetFullPath(path));
            string temp = Path.Combine(folder, "~" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp");
            try
            {
                writer.Save(temp);
                if (File.Exists(path)) File.Delete(path);
                File.Move(temp, path);
            }
            finally
            {
                try
                {
                    if (File.Exists(temp)) File.Delete(temp);
                }
                catch (Exception)
                {
                    // Fichier temporaire verrouillé : sans conséquence.
                }
            }
        }

        private static void ShowReport(IWin32Window owner, string path, ExportReport report)
        {
            var text = new StringBuilder();
            text.AppendFormat(CultureInfo.CurrentCulture, "Export terminé : {0} tableau{1} exporté{1} vers\n{2}\n", report.Exported, report.Exported > 1 ? "x" : string.Empty, path);

            var notes = report.Failures.Concat(report.Warnings).ToList();
            if (notes.Count > 0)
            {
                text.Append("\nRemarques :\n");
                foreach (var note in notes.Take(8)) text.Append("• ").Append(note).Append('\n');
                if (notes.Count > 8) text.AppendFormat(CultureInfo.CurrentCulture, "… et {0} autre(s) (voir le journal).\n", notes.Count - 8);
            }
            text.Append("\nVoulez-vous ouvrir le classeur maintenant ?");

            if (!Messages.Ask(owner, text.ToString())) return;
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Messages.Warning(owner, "Le classeur a bien été créé mais n'a pas pu être ouvert automatiquement :\n" + ex.Message);
            }
        }

        private sealed class ExportReport
        {
            public int Exported;
            public readonly List<string> Failures = new List<string>();
            public readonly List<string> Warnings = new List<string>();
        }
    }
}
