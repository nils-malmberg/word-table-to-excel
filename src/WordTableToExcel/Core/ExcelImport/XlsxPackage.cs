using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordTableToExcel.Core.Layout;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Erreur de lecture d'un classeur, avec un message destiné à l'utilisateur.</summary>
    public sealed class ExcelImportException : Exception
    {
        public ExcelImportException(string message)
            : base(message)
        {
        }

        public ExcelImportException(string message, Exception inner)
            : base(message, inner)
        {
        }

        public ExcelImportException(string message, bool convertibleByExcel)
            : base(message)
        {
            ConvertibleByExcel = convertibleByExcel;
        }

        /// <summary>Le fichier est dans un autre format de classeur (.xls, .xlsb, .ods…) qu'Excel sait convertir.</summary>
        public bool ConvertibleByExcel { get; private set; }
    }

    /// <summary>Relation OPC (fichier .rels).</summary>
    public sealed class OpcRelationship
    {
        public string Id;
        public string Type;
        /// <summary>Cible résolue (chemin de partie sans « / » initial) ou URL si <see cref="External"/>.</summary>
        public string Target;
        public bool External;

        /// <summary>true si le type de relation se termine par <paramref name="suffix"/> (« /worksheet »…), quel que soit l'espace de noms (transitionnel ou strict).</summary>
        public bool IsOfType(string suffix)
        {
            return Type != null && Type.EndsWith("/" + suffix, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>Accès aux parties XML d'un paquet OPC (xlsx) : lecture sûre, relations.</summary>
    internal static class XlsxPackage
    {
        public static XmlReaderSettings ReaderSettings()
        {
            return new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                CheckCharacters = false,
                CloseInput = true
            };
        }

        public static XDocument LoadXml(ZipReader zip, string partName)
        {
            byte[] data = zip.Read(partName);
            if (data == null) return null;
            try
            {
                using (var reader = XmlReader.Create(new MemoryStream(data, false), ReaderSettings()))
                {
                    return XDocument.Load(reader);
                }
            }
            catch (XmlException ex)
            {
                throw new ExcelImportException("Le classeur est endommagé (partie « " + partName + " » illisible).", ex);
            }
        }

        public static XmlReader OpenReader(ZipReader zip, string partName)
        {
            byte[] data = zip.Read(partName);
            if (data == null) return null;
            return XmlReader.Create(new MemoryStream(data, false), ReaderSettings());
        }

        /// <summary>Relations de la partie <paramref name="partName"/> (fichier _rels/xxx.rels voisin).</summary>
        public static List<OpcRelationship> Relationships(ZipReader zip, string partName)
        {
            var result = new List<OpcRelationship>();
            string directory = Directory(partName);
            string fileName = partName.Substring(directory.Length);
            string relsName = directory + "_rels/" + fileName + ".rels";
            var doc = LoadXml(zip, relsName);
            if (doc == null || doc.Root == null) return result;
            foreach (var rel in doc.Root.Elements())
            {
                if (!OoxmlXml.Is(rel, "Relationship")) continue;
                string target = OoxmlXml.Attr(rel, "Target");
                if (target == null) continue;
                bool external = string.Equals(OoxmlXml.Attr(rel, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                result.Add(new OpcRelationship
                {
                    Id = OoxmlXml.Attr(rel, "Id"),
                    Type = OoxmlXml.Attr(rel, "Type"),
                    External = external,
                    Target = external ? target : ResolveTarget(directory, target)
                });
            }
            return result;
        }

        /// <summary>Résout une cible relative (« ../theme/theme1.xml ») par rapport au dossier de la partie source.</summary>
        public static string ResolveTarget(string sourceDirectory, string target)
        {
            string t = Uri.UnescapeDataString(target.Replace('\\', '/'));
            int hash = t.IndexOf('#');
            if (hash >= 0) t = t.Substring(0, hash);
            var segments = new List<string>();
            if (!t.StartsWith("/", StringComparison.Ordinal))
            {
                foreach (var s in sourceDirectory.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)) segments.Add(s);
            }
            foreach (var s in t.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (s == ".") continue;
                if (s == "..")
                {
                    if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(s);
            }
            return string.Join("/", segments.ToArray());
        }

        /// <summary>Dossier d'une partie, avec « / » final (« xl/worksheets/ »), ou chaîne vide.</summary>
        public static string Directory(string partName)
        {
            int slash = partName.LastIndexOf('/');
            return slash < 0 ? string.Empty : partName.Substring(0, slash + 1);
        }

        /// <summary>
        /// Décode les caractères échappés par Excel sous la forme « _xHHHH_ » (caractères de contrôle,
        /// « _x005F_ » pour un tiret bas littéral).
        /// </summary>
        public static string DecodeEscapes(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf("_x", StringComparison.Ordinal) < 0) return text ?? string.Empty;
            var sb = new StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                if (text[i] == '_' && i + 6 < text.Length && text[i + 1] == 'x' && text[i + 6] == '_' && IsHex(text, i + 2, 4))
                {
                    int code = int.Parse(text.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    sb.Append((char)code);
                    i += 7;
                    continue;
                }
                sb.Append(text[i]);
                i++;
            }
            return sb.ToString();
        }

        private static bool IsHex(string s, int start, int length)
        {
            for (int i = start; i < start + length; i++)
            {
                char c = s[i];
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!hex) return false;
            }
            return true;
        }

        /// <summary>
        /// Identifie un fichier qui n'est pas un classeur .xlsx lisible et renvoie un message clair,
        /// ou null si le contenu ressemble à une archive ZIP.
        /// </summary>
        public static string DiagnoseNonZip(byte[] data)
        {
            bool convertible;
            return DiagnoseNonZip(data, out convertible);
        }

        public static string DiagnoseNonZip(byte[] data, out bool convertibleByExcel)
        {
            convertibleByExcel = false;
            if (data == null || data.Length == 0) return "Le fichier est vide.";
            if (data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B) return null; // « PK »
            if (data.Length >= 8 && data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0)
            {
                if (ContainsUtf16(data, "EncryptionInfo") || ContainsUtf16(data, "EncryptedPackage"))
                {
                    return "Ce classeur est protégé par un mot de passe (chiffré).\n\nOuvrez-le dans Excel, retirez le mot de passe "
                         + "(Fichier › Informations › Protéger le classeur), enregistrez-le, puis recommencez.";
                }
                convertibleByExcel = true;
                return "Ce fichier est au format Excel 97-2003 (.xls).\n\nOuvrez-le dans Excel et enregistrez-le au format "
                     + "« Classeur Excel (.xlsx) », puis recommencez.";
            }
            string head = Encoding.UTF8.GetString(data, 0, Math.Min(data.Length, 512)).TrimStart('\uFEFF', ' ', '\r', '\n', '\t');
            if (head.StartsWith("<", StringComparison.Ordinal))
            {
                convertibleByExcel = true;
                return "Ce fichier est au format XML ou HTML, pas au format Excel .xlsx.\n\nOuvrez-le dans Excel et enregistrez-le au format "
                     + "« Classeur Excel (.xlsx) », puis recommencez.";
            }
            return "Ce fichier n'est pas un classeur Excel (.xlsx) valide.";
        }

        private static bool ContainsUtf16(byte[] data, string text)
        {
            byte[] pattern = Encoding.Unicode.GetBytes(text);
            int limit = data.Length - pattern.Length;
            for (int i = 0; i <= limit; i++)
            {
                int j = 0;
                while (j < pattern.Length && data[i + j] == pattern[j]) j++;
                if (j == pattern.Length) return true;
            }
            return false;
        }
    }
}
