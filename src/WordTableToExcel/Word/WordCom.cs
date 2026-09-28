using System;
using System.Globalization;

namespace WordTableToExcel.Word
{
    /// <summary>Utilitaires de conversion des valeurs renvoyées par le modèle objet de Word (liaison tardive).</summary>
    internal static class WordCom
    {
        /// <summary>wdUndefined : valeur renvoyée quand une propriété est hétérogène sur la plage.</summary>
        public const int Undefined = 9999999;

        public static int AsInt(object value)
        {
            if (value == null) return 0;
            if (value is bool) return (bool)value ? -1 : 0;
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (OverflowException)
            {
                return Undefined;
            }
            catch (FormatException)
            {
                return Undefined;
            }
            catch (InvalidCastException)
            {
                return Undefined;
            }
        }

        public static double AsDouble(object value)
        {
            if (value == null) return 0;
            try
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                return Undefined;
            }
            catch (InvalidCastException)
            {
                return Undefined;
            }
        }

        public static string AsString(object value)
        {
            if (value == null) return string.Empty;
            return value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>Valeur booléenne Word (True = -1) ; false si indéfinie.</summary>
        public static bool IsTrue(object value)
        {
            int v = AsInt(value);
            return v != 0 && v != Undefined;
        }

        public static bool IsUndefined(int value)
        {
            return value == Undefined;
        }

        public static bool IsUndefined(double value)
        {
            return Math.Abs(value - Undefined) < 0.5;
        }
    }

    /// <summary>Suivi de progression et d'annulation pendant la lecture du document.</summary>
    public interface IExportProgress
    {
        void Report(string message, double fraction);
        bool IsCancellationRequested { get; }
    }
}
