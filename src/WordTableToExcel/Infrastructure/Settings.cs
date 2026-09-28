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
