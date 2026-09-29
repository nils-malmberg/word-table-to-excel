using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Export;
using WordTableToExcel.Import;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Autotest de l'exécutable, sans Word (intégration continue) : démarrage sur le .NET Framework, filtre de
    /// messages OLE, recherche de Word, construction des fenêtres (captures PNG), lecture d'un classeur.
    /// Usage : TableauxWordExcel.exe --selftest &lt;dossier de sortie&gt; [classeur.xlsx]
    /// Code de sortie 0 si tout est correct ; le détail est écrit dans selftest.txt.
    /// </summary>
    internal static class SelfTest
    {
        public static int Run(string[] args)
        {
            string output = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "TableauxWordExcel-selftest");
            string fixture = args.Length > 1 ? args[1] : null;
            Directory.CreateDirectory(output);

            var report = new StringBuilder();
            int failures = 0;
            Action<string, Action> check = (name, action) =>
            {
                try
                {
                    action();
                    report.AppendLine("OK     " + name);
                }
                catch (Exception ex)
                {
                    failures++;
                    report.AppendLine("ÉCHEC  " + name + " : " + ex);
                }
            };

            report.AppendLine("Tableaux Word ↔ Excel " + MainForm.Version + " — .NET " + Environment.Version + ", " + (IntPtr.Size * 8) + " bits");
            Messages.Title = MainForm.AppTitle;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (var filter = OleMessageFilter.Register())
            {
                check("Filtre de messages OLE", () =>
                {
                    if (!filter.IsRegistered) throw new InvalidOperationException("CoRegisterMessageFilter a échoué.");
                });

                check("Recherche de Word", () =>
                {
                    var instances = WordInstances.Find();
                    report.AppendLine("       instances de Word trouvées : " + instances.Count);
                });

                check("Fenêtre principale", () =>
                {
                    using (var form = new MainForm(new string[0]))
                    {
                        form.StartPosition = FormStartPosition.Manual;
                        form.Location = new Point(-20000, -20000);
                        form.ShowInTaskbar = false;
                        form.TopMost = false;
                        form.Show();
                        Application.DoEvents();
                        Capture(form, Path.Combine(output, "fenetre-sans-word.png"));

                        form.ShowPreview(SampleDocuments(), new InsertionInfo
                        {
                            DocumentName = "Rapport annuel 2024.docx",
                            Page = 3,
                            ParagraphText = "Les résultats du troisième trimestre confirment la progression des ventes dans toutes les régions.\r"
                        }, "Word 16.0 · 2 documents ouverts");
                        Application.DoEvents();
                        Capture(form, Path.Combine(output, "fenetre.png"));
                        form.Close();
                    }
                });

                check("Boîte de dialogue d'export", () =>
                {
                    var tables = new List<TableEntry>
                    {
                        new TableEntry { Index = 1, Caption = "Tableau 1 : Ventes trimestrielles par région", CaptionPosition = CaptionPosition.Above, StartPage = 3, EndPage = 4 },
                        new TableEntry { Index = 2, StartPage = 4, EndPage = 4 },
                        new TableEntry { Index = 3, Caption = "Tableau 2 : Effectifs", CaptionPosition = CaptionPosition.Above, StartPage = 7, EndPage = 7 }
                    };
                    using (var dialog = new ExportDialog("Rapport annuel 2024.docx", tables, CaptionPosition.Above, true, true, false))
                    {
                        ShowOffscreen(dialog);
                        Capture(dialog, Path.Combine(output, "export.png"));

                        // « Tout » : tous décochés puis tous cochés ; un tableau décoché n'est pas exporté.
                        dialog.ToggleAll();
                        Application.DoEvents();
                        if (dialog.SelectAllState() != CheckState.Unchecked || dialog.ExcludedTables.Count != 3) throw new InvalidOperationException("« Tout » n'a pas tout décoché.");
                        dialog.ToggleAll();
                        Application.DoEvents();
                        if (dialog.SelectAllState() != CheckState.Checked || dialog.ExcludedTables.Count != 0) throw new InvalidOperationException("« Tout » n'a pas tout coché.");
                        dialog.SetTableChecked(2, false);
                        Application.DoEvents();
                        var plan = ExportPlan.Build(tables, dialog.AllTables, dialog.ExcludedTables);
                        if (plan.Count != 2 || plan.Any(p => p.Table.Index == 2)) throw new InvalidOperationException("Le tableau décoché serait exporté.");
                        Capture(dialog, Path.Combine(output, "export-decoche.png"));
                        report.AppendLine("       export : « Tout », puis tableau 2 décoché → " + plan.Count + " feuilles");
                        dialog.Close();
                    }
                });

                if (fixture != null)
                {
                    check("Lecture du classeur et boîte de dialogue d'import", () =>
                    {
                        var workbook = XlsxWorkbook.Load(fixture);
                        var culture = CultureInfo.GetCultureInfo("fr-FR");
                        var format = new ExcelFormatSettings(culture, culture, workbook.Date1904, workbook.Styles.Colors);
                        var matcher = new CaptionMatcher(new List<string> { "Tableau" });
                        var candidates = ImportService.Analyze(workbook, format, matcher);
                        int importable = 0;
                        foreach (var c in candidates)
                        {
                            report.AppendLine("       feuille « " + c.Info.Name + " » : " + (c.CanImport ? c.Range.Value.ToString() : c.Problem));
                            if (c.CanImport) importable++;
                        }
                        if (importable == 0) throw new InvalidOperationException("aucune feuille importable");
                        using (var dialog = new ImportDialog(Path.GetFileName(fixture), candidates, new Settings(), false))
                        {
                            ShowOffscreen(dialog);
                            Capture(dialog, Path.Combine(output, "import.png"));
                            report.AppendLine("       case « Tout » à l'ouverture : " + dialog.SelectAllState());

                            // « Tout » : toutes les feuilles importables telles quelles cochées, puis aucune.
                            dialog.ToggleAll();
                            Application.DoEvents();
                            if (dialog.SelectAllState() != CheckState.Checked) throw new InvalidOperationException("« Tout » n'a pas tout coché.");
                            Capture(dialog, Path.Combine(output, "import-tout.png"));
                            dialog.ToggleAll();
                            Application.DoEvents();
                            if (dialog.SelectAllState() != CheckState.Unchecked) throw new InvalidOperationException("« Tout » n'a pas tout décoché.");
                            report.AppendLine("       case « Tout » : tout cocher puis tout décocher");
                            dialog.Close();
                        }
                    });
                }
            }

            report.AppendLine(failures == 0 ? "Autotest réussi." : failures + " vérification(s) en échec.");
            File.WriteAllText(Path.Combine(output, "selftest.txt"), report.ToString(), Encoding.UTF8);
            Log.Info("Autotest : " + (failures == 0 ? "réussi" : failures + " échec(s)"));
            return failures == 0 ? 0 : 1;
        }

        private static List<DocumentEntry> SampleDocuments()
        {
            return new List<DocumentEntry>
            {
                new DocumentEntry { State = DocumentState.Open, ProcessId = 1, Name = "Rapport annuel 2024.docx", FullName = @"C:\Users\moi\Documents\Rapport annuel 2024.docx", Folder = @"C:\Users\moi\Documents", TableCount = 12, IsActive = true },
                new DocumentEntry { State = DocumentState.Open, ProcessId = 1, Name = "Document2", FullName = "Document2", TableCount = 0 },
                DocumentEntry.ForFile(@"C:\Users\moi\Documents\Archives\Budget 2023.docx")
            };
        }

        private static void ShowOffscreen(Form form)
        {
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-20000, -20000);
            form.ShowInTaskbar = false;
            form.Show();
            Application.DoEvents();
        }

        private static void Capture(Control control, string path)
        {
            using (var bitmap = new Bitmap(control.Width, control.Height))
            {
                control.DrawToBitmap(bitmap, new Rectangle(Point.Empty, control.Size));
                bitmap.Save(path, ImageFormat.Png);
            }
        }
    }
}
