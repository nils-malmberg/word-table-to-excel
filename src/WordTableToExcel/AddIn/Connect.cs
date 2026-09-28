using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using WordTableToExcel.AddIn.Interop;
using WordTableToExcel.Export;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;

namespace WordTableToExcel.AddIn
{
    /// <summary>
    /// Point d'entrée du complément COM chargé par Word.
    /// <list type="bullet">
    /// <item>Word 2007 et suivants : bouton « Tableaux vers Excel » dans le ruban (onglets Accueil et Références) ;</item>
    /// <item>Word 2000 à 2003 : bouton dans la barre d'outils Standard.</item>
    /// </list>
    /// Toute exception est interceptée : une erreur de l'extension ne doit jamais faire planter Word
    /// ni conduire Word à désactiver le complément.
    /// </summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdValue)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public const string ClassId = "03F63233-F2FE-4A75-AF7A-F99CBEDC8030";
        public const string ProgIdValue = "WordTableToExcel.Connect";
        public const string FriendlyName = "Tableaux Word vers Excel";
        public const string Description = "Exporte les tableaux du document (avec ou sans légende) vers un classeur Excel, en conservant la mise en forme.";

        private object _application;
        private LegacyToolbar _legacyToolbar;
        private bool _busy;
        private static bool _visualStylesEnabled;

        // ------------------------------------------------------------------ IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = Application;
                Log.Info("Connexion à Word " + WordVersion(Application) + " (" + ConnectMode + "), .NET " + Environment.Version + ", " + (IntPtr.Size * 8) + " bits.");
                if (ConnectMode != ext_ConnectMode.ext_cm_Startup) SetupLegacyToolbar();
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_legacyToolbar != null) _legacyToolbar.Dispose();
            }
            catch (Exception ex)
            {
                Log.Error("OnDisconnection", ex);
            }
            finally
            {
                _legacyToolbar = null;
                _application = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                SetupLegacyToolbar();
            }
            catch (Exception ex)
            {
                Log.Error("OnStartupComplete", ex);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        // ------------------------------------------------------------------ Ruban

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("WordTableToExcel.Ribbon.xml"))
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Log.Error("GetCustomUI", ex);
                return string.Empty;
            }
        }

        /// <summary>Rappel du bouton du ruban (appelé par Office via IDispatch).</summary>
        public void OnExportClick(IRibbonControl control)
        {
            RunExport();
        }

        // ------------------------------------------------------------------ Export

        internal void RunExport()
        {
            if (_busy) return; // pas de double export
            _busy = true;
            try
            {
                EnableVisualStyles();
                new ExportService(_application).Run();
            }
            catch (Exception ex)
            {
                var inner = ex is ExportFailedException && ex.InnerException != null ? ex.InnerException : ex;
                Messages.Error(WindowOwner.FromWord(_application), "L'export des tableaux a échoué.", inner);
            }
            finally
            {
                _busy = false;
            }
        }

        private static void EnableVisualStyles()
        {
            if (_visualStylesEnabled) return;
            _visualStylesEnabled = true;
            try
            {
                Application.EnableVisualStyles();
            }
            catch (Exception)
            {
                // Apparence classique : sans importance.
            }
        }

        private void SetupLegacyToolbar()
        {
            if (_legacyToolbar != null || _application == null) return;
            if (WordMajorVersion(_application) >= 12) return; // le ruban est utilisé
            _legacyToolbar = LegacyToolbar.Create(_application, RunExport);
        }

        internal static string WordVersion(object application)
        {
            try
            {
                return Convert.ToString(((dynamic)application).Version);
            }
            catch (Exception)
            {
                return "?";
            }
        }

        internal static int WordMajorVersion(object application)
        {
            string version = WordVersion(application);
            int dot = version.IndexOf('.');
            int major;
            return int.TryParse(dot > 0 ? version.Substring(0, dot) : version, out major) ? major : 0;
        }

        // ------------------------------------------------------------------ Enregistrement (regasm)

        private const string AddInKeyPath = @"Software\Microsoft\Office\Word\Addins\" + ProgIdValue;

        /// <summary>
        /// Appelé par « regasm /codebase » (droits administrateur) : déclare le complément auprès de Word pour
        /// tous les utilisateurs. RegAsm 64 bits écrit dans la vue 64 bits du registre (Office 64 bits),
        /// RegAsm 32 bits dans la vue WOW6432Node (Office 32 bits).
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            if (type != typeof(Connect)) return;
            using (var key = Registry.LocalMachine.CreateSubKey(AddInKeyPath))
            {
                if (key == null) return;
                key.SetValue("FriendlyName", FriendlyName);
                key.SetValue("Description", Description);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            if (type != typeof(Connect)) return;
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(AddInKeyPath);
            }
            catch (ArgumentException)
            {
                // Clé déjà absente.
            }
        }
    }
}
