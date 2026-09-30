using System;
using System.Globalization;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Trace du Word invisible lancé par l'application (numéro de processus et heure de démarrage), écrite tant qu'il
    /// tourne : si l'application s'arrête brutalement, le démarrage suivant le retrouve et le ferme. L'heure de
    /// démarrage garantit qu'on ne confond pas avec un autre programme qui aurait repris le même numéro.
    /// </summary>
    internal static class HiddenWordRecord
    {
        public static string Format(int processId, DateTime startUtc)
        {
            return processId.ToString(CultureInfo.InvariantCulture) + ";" + startUtc.Ticks.ToString(CultureInfo.InvariantCulture);
        }

        public static bool TryParse(string text, out int processId, out DateTime startUtc)
        {
            processId = 0;
            startUtc = DateTime.MinValue;
            if (string.IsNullOrEmpty(text)) return false;
            string[] parts = text.Trim().Split(';');
            long ticks;
            if (parts.Length != 2
                || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out processId) || processId <= 0
                || !long.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out ticks)
                || ticks <= 0 || ticks > DateTime.MaxValue.Ticks)
            {
                processId = 0;
                return false;
            }
            startUtc = new DateTime(ticks, DateTimeKind.Utc);
            return true;
        }

        /// <summary>
        /// Le processus trouvé sous ce numéro est-il bien le Word invisible noté (programme Word, démarré à la même heure,
        /// sans fenêtre visible) ? Sinon, on n'y touche pas.
        /// </summary>
        public static bool IsOrphan(string processName, DateTime processStartUtc, DateTime recordedStartUtc, bool hasVisibleWindow)
        {
            return string.Equals(processName, "WINWORD", StringComparison.OrdinalIgnoreCase)
                && Math.Abs((processStartUtc - recordedStartUtc).TotalSeconds) < 2
                && !hasVisibleWindow;
        }
    }
}
