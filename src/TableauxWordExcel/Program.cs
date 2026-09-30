using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Application « Tableaux Word ↔ Excel » : un seul fichier .exe, sans installation ni droits d'administrateur.
    /// Elle pilote Word de l'extérieur (comme le ferait un script) : rien n'est chargé dans Word, les règles de
    /// Word sur les compléments ne s'appliquent donc pas.
    /// </summary>
    internal static class Program
    {
        private const string MutexName = @"Local\TableauxWordExcel-8C1F5E2A";

        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length > 0 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
            {
                return SelfTest.Run(args.Skip(1).ToArray());
            }

            bool firstInstance;
            using (var mutex = new Mutex(true, MutexName, out firstInstance))
            {
                if (!firstInstance)
                {
                    // Déjà ouverte : on remet simplement sa fenêtre au premier plan.
                    ActivateRunningInstance();
                    return 0;
                }

                Messages.Title = MainForm.AppTitle;
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) => Messages.Error(null, "Erreur inattendue de l'application.", e.Exception, false);
                AppDomain.CurrentDomain.UnhandledException += (s, e) => Log.Error("Erreur fatale", e.ExceptionObject as Exception);
                Log.Info("Démarrage de l'application " + MainForm.Version + ", .NET " + Environment.Version + ", " + (IntPtr.Size * 8) + " bits.");
                HiddenWord.CleanUpOrphan(); // Word invisible laissé ouvert par un arrêt brutal précédent

                using (var filter = OleMessageFilter.Register())
                {
                    if (!filter.IsRegistered) Log.Info("Filtre de messages OLE non installé.");
                    Application.Run(new MainForm(args));
                }
                GC.KeepAlive(mutex);
            }
            return 0;
        }

        private static void ActivateRunningInstance()
        {
            try
            {
                var current = Process.GetCurrentProcess();
                foreach (var process in Process.GetProcessesByName(current.ProcessName))
                {
                    if (process.Id == current.Id || process.MainWindowHandle == IntPtr.Zero) continue;
                    NativeMethods.BringToFront(process.MainWindowHandle);
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Info("Activation de l'application déjà ouverte : " + ex.Message);
            }
        }
    }
}
