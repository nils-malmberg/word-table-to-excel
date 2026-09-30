using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.Word;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Word invisible pour exporter les fichiers qui ne sont pas ouverts : une instance distincte (les documents de
    /// l'utilisateur ne sont pas touchés), sans macros ; chaque fichier y est ouvert en lecture seule puis refermé sans
    /// rien enregistrer. L'instance est gardée quelques minutes après un export, pour que le suivant démarre aussitôt,
    /// puis refermée ; elle l'est toujours à la fermeture de l'application.
    /// <para>Précautions :</para>
    /// <list type="bullet">
    /// <item>un document que l'utilisateur ouvre entre-temps (double-clic dans l'Explorateur) peut arriver dans cette
    /// instance : elle est alors rendue visible et laissée à l'utilisateur, jamais refermée par l'application ;</item>
    /// <item>si l'application s'arrête brutalement, l'instance restée ouverte est refermée au démarrage suivant
    /// (voir <see cref="CleanUpOrphan"/>).</item>
    /// </list>
    /// </summary>
    internal sealed class HiddenWord : IDisposable
    {
        /// <summary>Durée pendant laquelle le Word invisible est gardé après un export.</summary>
        public static readonly TimeSpan IdleLifetime = TimeSpan.FromMinutes(3);

        private const int WdAlertsNone = 0;
        private const int WdAlertsAll = -1;
        private const int WdDoNotSaveChanges = 0;
        private const int MsoAutomationSecurityByUI = 2;
        private const int MsoAutomationSecurityForceDisable = 3;

        private object _application;
        private object _document;
        private int _processId;
        private DateTime _idleSinceUtc;

        /// <summary>Objet Word.Application de l'instance invisible (null si elle n'est pas lancée).</summary>
        public object Application
        {
            get { return _application; }
        }

        /// <summary>Instance lancée (et encore utilisable, à vérifier par <see cref="Open"/>).</summary>
        public bool IsRunning
        {
            get { return _application != null; }
        }

        /// <summary>Trace du Word invisible en cours (voir <see cref="HiddenWordRecord"/>).</summary>
        internal static string RecordPath
        {
            get { return Path.Combine(Path.GetDirectoryName(Log.FilePath), "word-invisible.txt"); }
        }

        /// <summary>Ouvre le fichier en lecture seule dans le Word invisible (lancé s'il ne l'est pas déjà).</summary>
        /// <returns>Objet Word.Document, à refermer par <see cref="CloseDocument"/>.</returns>
        public object Open(string path)
        {
            CloseDocument();
            EnsureApplication();
            dynamic app = _application;
            // Un mot de passe factice : pour un document protégé, Word renvoie une erreur au lieu d'afficher
            // une demande de mot de passe dans une fenêtre invisible.
            _document = app.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false,
                PasswordDocument: "\u0001", Revert: false, Visible: false, NoEncodingDialog: true);
            if (_processId == 0) Remember(ProcessOfDocument(_document));
            Log.Info("Document ouvert en arrière-plan : " + path);
            return _document;
        }

        /// <summary>Referme le document ouvert par <see cref="Open"/> (sans rien enregistrer) ; Word reste prêt pour le suivant.</summary>
        public void CloseDocument()
        {
            object document = _document;
            _document = null;
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
            _idleSinceUtc = DateTime.UtcNow;
            WordCom.ReleaseUnusedReferences();
            // Un document de l'utilisateur est peut-être arrivé dans cette instance pendant l'export.
            if (_application != null && DocumentCount() > 0) HandOver();
        }

        /// <summary>
        /// À appeler régulièrement quand aucune opération n'est en cours : rend Word à l'utilisateur si un de ses
        /// documents y est arrivé, le referme après <see cref="IdleLifetime"/> d'inactivité.
        /// </summary>
        public void CheckIdle()
        {
            if (_application == null || _document != null) return;
            int count = DocumentCount();
            if (count < 0) Forget("Word invisible fermé par ailleurs.");
            else if (count > 0) HandOver();
            else if (DateTime.UtcNow - _idleSinceUtc > IdleLifetime) Quit();
        }

        /// <summary>Fermeture de l'application : Word invisible refermé (ou rendu à l'utilisateur s'il contient un de ses documents).</summary>
        public void Dispose()
        {
            CloseDocument();
            if (_application != null) Quit();
        }

        // ------------------------------------------------------------------ instance

        private void EnsureApplication()
        {
            if (_application != null)
            {
                int count = DocumentCount();
                if (count == 0) return; // prête : réutilisée
                if (count > 0) HandOver();
                else Forget("Word invisible fermé par ailleurs, relancé.");
            }

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

            var before = new HashSet<int>(WordProcessIds());
            _application = Activator.CreateInstance(type);
            dynamic app = _application;
            TrySet(() => app.Visible = false);
            TrySet(() => app.DisplayAlerts = WdAlertsNone);
            TrySet(() => app.AutomationSecurity = MsoAutomationSecurityForceDisable);
            TrySet(() => app.ScreenUpdating = false);
            _processId = 0;
            // Processus du Word lancé : le seul nouveau processus Word (sinon, retrouvé par la fenêtre du document ouvert).
            var started = WordProcessIds().Where(id => !before.Contains(id)).ToList();
            if (started.Count == 1) Remember(started[0]);
            Log.Info("Word lancé en arrière-plan" + (_processId != 0 ? " (processus " + _processId + ")." : "."));
        }

        /// <summary>Nombre de documents ouverts dans le Word invisible ; -1 s'il ne répond plus (fermé, planté).</summary>
        private int DocumentCount()
        {
            try
            {
                return WordCom.AsInt(((dynamic)_application).Documents.Count);
            }
            catch (Exception ex)
            {
                Log.Info("Word invisible injoignable : " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// Un document de l'utilisateur est arrivé dans le Word invisible : Word est rendu visible, avec ses réglages
        /// habituels, et laissé à l'utilisateur (l'application ne le refermera pas).
        /// </summary>
        private void HandOver()
        {
            dynamic app = _application;
            TrySet(() => app.ScreenUpdating = true);
            TrySet(() => app.DisplayAlerts = WdAlertsAll);
            TrySet(() => app.AutomationSecurity = MsoAutomationSecurityByUI);
            TrySet(() => app.Visible = true);
            Forget("Un document ouvert par l'utilisateur est arrivé dans le Word invisible : Word est rendu visible et laissé à l'utilisateur.");
        }

        private void Quit()
        {
            // Jamais de document de l'utilisateur refermé sans enregistrer : Word lui est rendu.
            if (DocumentCount() > 0)
            {
                HandOver();
                return;
            }
            dynamic app = _application;
            TrySet(() => app.NormalTemplate.Saved = true); // jamais de question sur le modèle Normal
            try
            {
                app.Quit(SaveChanges: WdDoNotSaveChanges);
            }
            catch (Exception ex)
            {
                Log.Info("Fermeture de Word en arrière-plan : " + ex.Message);
            }
            Forget("Word invisible refermé.");
        }

        /// <summary>Oublie l'instance (sans la fermer) : références libérées, trace effacée.</summary>
        private void Forget(string reason)
        {
            object application = _application;
            _application = null;
            _processId = 0;
            Release(application);
            DeleteRecord();
            WordCom.ReleaseUnusedReferences();
            Log.Info(reason);
        }

        // ------------------------------------------------------------------ Word invisible orphelin

        /// <summary>
        /// Au démarrage : si l'application s'est arrêtée brutalement en laissant son Word invisible ouvert, ce Word
        /// est refermé (seulement s'il s'agit bien de lui : même processus, même heure de démarrage, aucune fenêtre visible).
        /// </summary>
        public static void CleanUpOrphan()
        {
            try
            {
                string path = RecordPath;
                if (!File.Exists(path)) return;
                string text = File.ReadAllText(path);
                DeleteRecord();

                int processId;
                DateTime startUtc;
                if (!HiddenWordRecord.TryParse(text, out processId, out startUtc)) return;
                Process process;
                try
                {
                    process = Process.GetProcessById(processId);
                }
                catch (ArgumentException)
                {
                    return; // déjà terminé
                }
                using (process)
                {
                    if (!HiddenWordRecord.IsOrphan(process.ProcessName, process.StartTime.ToUniversalTime(), startUtc, HasVisibleWindow(processId))) return;
                    process.Kill();
                    process.WaitForExit(5000);
                    Log.Info("Word invisible resté ouvert après un arrêt brutal de l'application (processus " + processId + ") : refermé.");
                }
            }
            catch (Exception ex)
            {
                Log.Info("Recherche d'un Word invisible resté ouvert : " + ex.Message);
            }
        }

        private void Remember(int processId)
        {
            if (processId <= 0) return;
            try
            {
                using (var process = Process.GetProcessById(processId))
                {
                    string path = RecordPath;
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    File.WriteAllText(path, HiddenWordRecord.Format(processId, process.StartTime.ToUniversalTime()));
                }
                _processId = processId;
            }
            catch (Exception ex)
            {
                Log.Info("Processus du Word invisible non noté : " + ex.Message);
            }
        }

        private static void DeleteRecord()
        {
            try
            {
                if (File.Exists(RecordPath)) File.Delete(RecordPath);
            }
            catch (Exception)
            {
                // Sans conséquence : la trace est vérifiée avant toute fermeture.
            }
        }

        private static List<int> WordProcessIds()
        {
            var ids = new List<int>();
            try
            {
                foreach (var process in Process.GetProcessesByName("WINWORD"))
                {
                    using (process) ids.Add(process.Id);
                }
            }
            catch (Exception)
            {
                // Liste des processus inaccessible : le processus sera retrouvé par la fenêtre du document.
            }
            return ids;
        }

        /// <summary>Processus propriétaire de la fenêtre (invisible) du document, Word 2013 et suivants ; 0 si inconnu.</summary>
        private static int ProcessOfDocument(object document)
        {
            try
            {
                var hwnd = new IntPtr(Convert.ToInt64(((dynamic)document).ActiveWindow.Hwnd));
                uint pid;
                NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                return (int)pid;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool HasVisibleWindow(int processId)
        {
            bool visible = false;
            NativeMethods.EnumWindowsProc callback = (hwnd, l) =>
            {
                uint pid;
                NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                if ((int)pid == processId && NativeMethods.IsWindowVisible(hwnd))
                {
                    visible = true;
                    return false;
                }
                return true;
            };
            NativeMethods.EnumWindows(callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            return visible;
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
