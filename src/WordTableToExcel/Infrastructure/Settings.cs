using System;
using Microsoft.Win32;

namespace WordTableToExcel.Infrastructure
{
    /// <summary>Préférences de l'utilisateur, mémorisées dans HKCU\Software\WordTableToExcel.</summary>
    internal sealed class Settings
    {
        private const string KeyPath = @"Software\WordTableToExcel";

        public bool AllTables;
        public bool IncludeCaptionRow = true;
        public bool ConvertNumbers;
        public string LastFolder;
        /// <summary>Libellés de légende supplémentaires, séparés par des points-virgules (réglage avancé).</summary>
        public string ExtraCaptionLabels;

        // Import Excel → Word
        public bool ImportAddCaption = true;
        /// <summary>null = selon les légendes déjà présentes dans le document.</summary>
        public bool? ImportCaptionBelow;
        public bool ImportFitToPage = true;
        public bool ImportSkipHidden = true;
        public bool ImportGridlines;
        public string ImportLastFolder;

        // Application autonome
        /// <summary>Fenêtre toujours au premier plan (pratique pour cliquer dans Word puis dans l'application).</summary>
        public bool AppTopMost = true;
        /// <summary>Position et taille de la fenêtre : « x,y,largeur,hauteur » (pixels).</summary>
        public string AppBounds;
        /// <summary>Dossier du dernier fichier Word ajouté à la liste.</summary>
        public string AppLastWordFolder;

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (key == null) return s;
                    s.AllTables = ReadBool(key, "AllTables", s.AllTables);
                    s.IncludeCaptionRow = ReadBool(key, "IncludeCaptionRow", s.IncludeCaptionRow);
                    s.ConvertNumbers = ReadBool(key, "ConvertNumbers", s.ConvertNumbers);
                    s.LastFolder = key.GetValue("LastFolder") as string;
                    s.ExtraCaptionLabels = key.GetValue("ExtraCaptionLabels") as string;
                    s.ImportAddCaption = ReadBool(key, "ImportAddCaption", s.ImportAddCaption);
                    object below = key.GetValue("ImportCaptionBelow");
                    if (below is int) s.ImportCaptionBelow = (int)below != 0;
                    s.ImportFitToPage = ReadBool(key, "ImportFitToPage", s.ImportFitToPage);
                    s.ImportSkipHidden = ReadBool(key, "ImportSkipHidden", s.ImportSkipHidden);
                    s.ImportGridlines = ReadBool(key, "ImportGridlines", s.ImportGridlines);
                    s.ImportLastFolder = key.GetValue("ImportLastFolder") as string;
                    s.AppTopMost = ReadBool(key, "AppTopMost", s.AppTopMost);
                    s.AppBounds = key.GetValue("AppBounds") as string;
                    s.AppLastWordFolder = key.GetValue("AppLastWordFolder") as string;
                }
            }
            catch (Exception ex)
            {
                Log.Error("Lecture des préférences impossible", ex);
            }
            return s;
        }

        public void Save()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    if (key == null) return;
                    key.SetValue("AllTables", AllTables ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("IncludeCaptionRow", IncludeCaptionRow ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("ConvertNumbers", ConvertNumbers ? 1 : 0, RegistryValueKind.DWord);
                    if (!string.IsNullOrEmpty(LastFolder)) key.SetValue("LastFolder", LastFolder);
                    key.SetValue("ImportAddCaption", ImportAddCaption ? 1 : 0, RegistryValueKind.DWord);
                    if (ImportCaptionBelow.HasValue) key.SetValue("ImportCaptionBelow", ImportCaptionBelow.Value ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("ImportFitToPage", ImportFitToPage ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("ImportSkipHidden", ImportSkipHidden ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("ImportGridlines", ImportGridlines ? 1 : 0, RegistryValueKind.DWord);
                    if (!string.IsNullOrEmpty(ImportLastFolder)) key.SetValue("ImportLastFolder", ImportLastFolder);
                    key.SetValue("AppTopMost", AppTopMost ? 1 : 0, RegistryValueKind.DWord);
                    if (!string.IsNullOrEmpty(AppBounds)) key.SetValue("AppBounds", AppBounds);
                    if (!string.IsNullOrEmpty(AppLastWordFolder)) key.SetValue("AppLastWordFolder", AppLastWordFolder);
                }
            }
            catch (Exception ex)
            {
                Log.Error("Enregistrement des préférences impossible", ex);
            }
        }

        public string[] ExtraLabels()
        {
            if (string.IsNullOrEmpty(ExtraCaptionLabels)) return new string[0];
            return ExtraCaptionLabels.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static bool ReadBool(RegistryKey key, string name, bool defaultValue)
        {
            object v = key.GetValue(name);
            if (v is int) return (int)v != 0;
            return defaultValue;
        }
    }
}
