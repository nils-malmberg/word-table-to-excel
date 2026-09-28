using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace WordTableToExcel.Infrastructure
{
    /// <summary>Journal de diagnostic : %LOCALAPPDATA%\WordTableToExcel\WordTableToExcel.log.</summary>
    internal static class Log
    {
        private const long MaxSize = 1024 * 1024;
        private static readonly object Sync = new object();

        public static string FilePath
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordTableToExcel");
                return Path.Combine(folder, "WordTableToExcel.log");
            }
        }

        public static void Info(string message)
        {
            Write("INFO ", message);
        }

        public static void Error(string message, Exception ex)
        {
            Write("ERROR", message + (ex == null ? string.Empty : Environment.NewLine + ex));
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (Sync)
                {
                    string path = FilePath;
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    var info = new FileInfo(path);
                    if (info.Exists && info.Length > MaxSize)
                    {
                        string old = path + ".1";
                        if (File.Exists(old)) File.Delete(old);
                        File.Move(path, old);
                    }
                    string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + level + " " + message + Environment.NewLine;
                    File.AppendAllText(path, line, Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // Le journal ne doit jamais perturber Word.
            }
        }
    }
}
