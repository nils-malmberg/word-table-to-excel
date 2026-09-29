using System;
using System.IO;
using System.Runtime.InteropServices;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.Import
{
    /// <summary>
    /// Conversion d'un classeur d'un autre format (.xls, .xlsb, .ods, .csv…) en .xlsx temporaire, par Excel
    /// lui-même, s'il est installé : Excel est lancé en arrière-plan, sans macros ni mises à jour de liaisons,
    /// le fichier d'origine est ouvert en lecture seule et n'est jamais modifié.
    /// </summary>
    internal static class ExcelFileConverter
    {
        private const int XlOpenXmlWorkbook = 51;
        private const int MsoAutomationSecurityForceDisable = 3;

        public static bool IsExcelInstalled()
        {
            try
            {
                return Type.GetTypeFromProgID("Excel.Application") != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>Extensions ouvertes directement par l'extension (sans Excel).</summary>
        public static bool IsOpenXml(string path)
        {
            string ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
            return ext == ".xlsx" || ext == ".xlsm" || ext == ".xltx" || ext == ".xltm";
        }

        /// <summary>Convertit le classeur en .xlsx temporaire (à supprimer par l'appelant).</summary>
        public static string ConvertToXlsx(string path)
        {
            Type type = null;
            try
            {
                type = Type.GetTypeFromProgID("Excel.Application");
            }
            catch (Exception ex)
            {
                Log.Info("Excel introuvable : " + ex.Message);
            }
            if (type == null)
            {
                throw new ExcelImportException("Ce classeur n'est pas au format .xlsx et Excel n'est pas installé pour le convertir.\n\n"
                    + "Enregistrez-le au format « Classeur Excel (.xlsx) », puis recommencez.");
            }

            string temp = Path.Combine(Path.GetTempPath(), "WordTableToExcel-" + Guid.NewGuid().ToString("N") + ".xlsx");
            object excelObject = null;
            object workbookObject = null;
            try
            {
                excelObject = Activator.CreateInstance(type);
                dynamic excel = excelObject;
                TrySet(() => excel.Visible = false);
                TrySet(() => excel.DisplayAlerts = false);
                TrySet(() => excel.ScreenUpdating = false);
                TrySet(() => excel.AskToUpdateLinks = false);
                TrySet(() => excel.EnableEvents = false);
                TrySet(() => excel.AutomationSecurity = MsoAutomationSecurityForceDisable);

                // Un mot de passe factice évite qu'Excel n'affiche une demande de mot de passe (erreur à la place).
                workbookObject = excel.Workbooks.Open(Filename: path, UpdateLinks: 0, ReadOnly: true, IgnoreReadOnlyRecommended: true,
                    Password: "\u0001", Notify: false, AddToMru: false);
                dynamic workbook = workbookObject;
                workbook.SaveAs(Filename: temp, FileFormat: XlOpenXmlWorkbook);
                workbook.Close(SaveChanges: false);
                workbookObject = null;
                return temp;
            }
            catch (Exception ex)
            {
                Log.Error("Conversion par Excel", ex);
                try
                {
                    if (File.Exists(temp)) File.Delete(temp);
                }
                catch (Exception)
                {
                    // Fichier temporaire : sans conséquence.
                }
                throw new ExcelImportException("Excel n'a pas pu ouvrir ou convertir ce classeur (fichier protégé par mot de passe, endommagé ou format inconnu).\n\n"
                    + "Détail : " + ex.Message, ex);
            }
            finally
            {
                if (workbookObject != null)
                {
                    try
                    {
                        ((dynamic)workbookObject).Close(SaveChanges: false);
                    }
                    catch (Exception)
                    {
                        // Déjà fermé.
                    }
                    Release(workbookObject);
                }
                if (excelObject != null)
                {
                    try
                    {
                        ((dynamic)excelObject).Quit();
                    }
                    catch (Exception)
                    {
                        // Excel déjà fermé.
                    }
                    Release(excelObject);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void TrySet(Action action)
        {
            try
            {
                action();
            }
            catch (Exception)
            {
                // Propriété absente de cette version d'Excel.
            }
        }

        private static void Release(object com)
        {
            try
            {
                if (com != null && Marshal.IsComObject(com)) Marshal.FinalReleaseComObject(com);
            }
            catch (Exception)
            {
                // Objet déjà libéré.
            }
        }
    }
}
