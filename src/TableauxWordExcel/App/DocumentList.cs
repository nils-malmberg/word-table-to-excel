using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordTableToExcel.App
{
    /// <summary>Origine d'un document proposé à l'export.</summary>
    internal enum DocumentState
    {
        /// <summary>Ouvert dans Word.</summary>
        Open,
        /// <summary>Ouvert dans Word en mode protégé (lecture seule).</summary>
        ProtectedView,
        /// <summary>Fichier ajouté par l'utilisateur, ouvert en arrière-plan le temps de l'export.</summary>
        File
    }

    /// <summary>Ligne de la liste des documents : document ouvert dans Word ou fichier ajouté.</summary>
    internal sealed class DocumentEntry
    {
        public DocumentState State;
        /// <summary>Processus Word qui a ouvert le document (0 : inconnu ou fichier).</summary>
        public int ProcessId;
        public string Name = string.Empty;
        /// <summary>Chemin complet (ou adresse OneDrive/SharePoint) ; nom seul pour un document jamais enregistré.</summary>
        public string FullName = string.Empty;
        /// <summary>Dossier du document ; vide s'il n'a jamais été enregistré.</summary>
        public string Folder = string.Empty;
        /// <summary>Nombre de tableaux ; -1 si inconnu (fichier non ouvert).</summary>
        public int TableCount = -1;
        /// <summary>Document actif de Word (celui dans lequel l'utilisateur a cliqué en dernier).</summary>
        public bool IsActive;

        public bool IsSaved
        {
            get { return !string.IsNullOrEmpty(Folder); }
        }

        /// <summary>Identifiant stable d'une actualisation à l'autre (pour conserver la sélection).</summary>
        public string Key
        {
            get
            {
                string path = (FullName ?? string.Empty).ToUpperInvariant();
                switch (State)
                {
                    case DocumentState.File: return "F|" + path;
                    case DocumentState.ProtectedView: return ProcessId + "|P|" + path;
                    default: return ProcessId + "|D|" + path;
                }
            }
        }

        public static DocumentEntry ForFile(string path)
        {
            return new DocumentEntry
            {
                State = DocumentState.File,
                Name = Path.GetFileName(path) ?? path,
                FullName = path,
                Folder = Path.GetDirectoryName(path) ?? string.Empty
            };
        }
    }

    /// <summary>Règles de la liste des documents (sans dépendance à Word ni à l'interface).</summary>
    internal static class DocumentList
    {
        private static readonly string[] WordExtensions = { ".docx", ".docm", ".doc", ".dotx", ".dotm", ".dot", ".rtf", ".odt" };
        private static readonly string[] ExcelExtensions = { ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods", ".csv" };

        public static bool IsWordFile(string path)
        {
            return HasExtension(path, WordExtensions);
        }

        public static bool IsExcelFile(string path)
        {
            return HasExtension(path, ExcelExtensions);
        }

        /// <summary>Filtre des boîtes de dialogue d'ouverture de fichiers Word.</summary>
        public static string WordFileFilter
        {
            get
            {
                return "Documents Word (*.docx;*.docm;*.doc;*.rtf;*.odt)|*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.rtf;*.odt|Tous les fichiers (*.*)|*.*";
            }
        }

        /// <summary>
        /// Liste affichée : documents ouverts dans Word, puis fichiers ajoutés. Un fichier déjà ouvert dans Word
        /// n'apparaît qu'une fois, comme document ouvert (c'est la version affichée à l'écran, avec les modifications
        /// non enregistrées, qui est exportée).
        /// </summary>
        public static List<DocumentEntry> Merge(IEnumerable<DocumentEntry> open, IEnumerable<string> files)
        {
            var result = new List<DocumentEntry>();
            foreach (var entry in open)
            {
                if (result.Any(r => r.Key == entry.Key)) continue;
                result.Add(entry);
            }
            foreach (var file in files)
            {
                if (string.IsNullOrEmpty(file)) continue;
                if (result.Any(r => (r.State == DocumentState.File || r.IsSaved) && SamePath(r.FullName, file))) continue;
                result.Add(DocumentEntry.ForFile(file));
            }
            return result;
        }

        /// <summary>Même fichier (chemins comparés sans tenir compte de la casse, après normalisation).</summary>
        public static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            return string.Equals(Normalize(a), Normalize(b), StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string path)
        {
            string p = path.Trim();
            if (p.IndexOf("://", StringComparison.Ordinal) > 0) return p.Replace('\\', '/'); // OneDrive, SharePoint
            try
            {
                return Path.GetFullPath(p);
            }
            catch (Exception)
            {
                return p;
            }
        }

        private static bool HasExtension(string path, string[] extensions)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string ext;
            try
            {
                ext = Path.GetExtension(path);
            }
            catch (ArgumentException)
            {
                return false;
            }
            return !string.IsNullOrEmpty(ext) && extensions.Contains(ext.ToLowerInvariant());
        }
    }
}
