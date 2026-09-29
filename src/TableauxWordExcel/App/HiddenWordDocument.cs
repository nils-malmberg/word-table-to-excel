using System;
using System.Runtime.InteropServices;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Document Word ouvert en arrière-plan pour exporter un fichier qui n'est pas ouvert : une instance de Word
    /// distincte et invisible est lancée (les documents de l'utilisateur ne sont pas touchés), sans macros,
    /// le fichier est ouvert en lecture seule, puis Word est refermé sans rien enregistrer.
    /// </summary>
    internal sealed class HiddenWordDocument : IDisposable
    {
        private const int WdAlertsNone = 0;
        private const int WdDoNotSaveChanges = 0;
        private const int MsoAutomationSecurityForceDisable = 3;

        private HiddenWordDocument()
        {
        }

        /// <summary>Objet Word.Application de l'instance invisible.</summary>
        public object Application { get; private set; }

        /// <summary>Objet Word.Document ouvert en lecture seule.</summary>
        public object Document { get; private set; }

        public static HiddenWordDocument Open(string path)
        {
            Type type = null;
            try
            {
                type = Type.GetTypeFromProgID("Word.Application");
            }
            catch (Exception ex)
            {
                Log.Info("Word introuvable : " + ex.Message);
            }
            if (type == null) throw new InvalidOperationException("Microsoft Word n'est pas installé sur cet ordinateur.");

            var result = new HiddenWordDocument();
            try
            {
                result.Application = Activator.CreateInstance(type);
                dynamic app = result.Application;
                TrySet(() => app.Visible = false);
                TrySet(() => app.DisplayAlerts = WdAlertsNone);
                TrySet(() => app.AutomationSecurity = MsoAutomationSecurityForceDisable);
                TrySet(() => app.ScreenUpdating = false);

                // Un mot de passe factice : pour un document protégé, Word renvoie une erreur au lieu d'afficher
                // une demande de mot de passe dans une fenêtre invisible.
                result.Document = app.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false,
                    PasswordDocument: "\u0001", Revert: false, Visible: false, NoEncodingDialog: true);
                Log.Info("Document ouvert en arrière-plan : " + path);
                return result;
            }
            catch (Exception)
            {
                result.Dispose();
                throw;
            }
        }

        public void Dispose()
        {
            object document = Document;
            object application = Application;
            Document = null;
            Application = null;

            if (document != null)
            {
                try
                {
                    ((dynamic)document).Close(SaveChanges: WdDoNotSaveChanges);
                }
                catch (Exception ex)
                {
                    Log.Info("Fermeture du document en arrière-plan : " + ex.Message);
                }
                Release(document);
            }
            if (application != null)
            {
                dynamic app = application;
                TrySet(() => app.NormalTemplate.Saved = true); // jamais de question sur le modèle Normal
                try
                {
                    app.Quit(SaveChanges: WdDoNotSaveChanges);
                }
                catch (Exception ex)
                {
                    Log.Info("Fermeture de Word en arrière-plan : " + ex.Message);
                }
                Release(application);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void TrySet(Action action)
        {
            try
            {
                action();
            }
            catch (Exception)
            {
                // Propriété absente de cette version de Word.
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
