using System;
using System.Windows.Forms;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.UI
{
    internal static class Messages
    {
        public const string Title = "Tableaux Word vers Excel";

        public static void Info(IWin32Window owner, string text)
        {
            MessageBox.Show(owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void Warning(IWin32Window owner, string text)
        {
            MessageBox.Show(owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static bool Ask(IWin32Window owner, string text)
        {
            return MessageBox.Show(owner, text, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes;
        }

        /// <summary>Affiche une erreur inattendue ; l'extension ne laisse jamais une exception remonter jusqu'à Word.</summary>
        public static void Error(IWin32Window owner, string context, Exception ex)
        {
            Error(owner, context, ex, true);
        }

        /// <param name="documentUnchanged">true : rappeler que le document Word n'a pas été modifié (export).</param>
        public static void Error(IWin32Window owner, string context, Exception ex, bool documentUnchanged)
        {
            Log.Error(context, ex);
            try
            {
                string text = context + Environment.NewLine + Environment.NewLine
                    + (documentUnchanged ? "Le document Word n'a pas été modifié." + Environment.NewLine + Environment.NewLine : string.Empty)
                    + "Détail : " + (ex == null ? "(aucun)" : ex.Message) + Environment.NewLine + Environment.NewLine
                    + "Journal : " + Log.FilePath;
                MessageBox.Show(owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception)
            {
                // Rien de plus à faire.
            }
        }
    }
}
