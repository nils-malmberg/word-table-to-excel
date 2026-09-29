using System;
using System.Globalization;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>Date et heure décomposées d'un numéro de série Excel.</summary>
    public struct ExcelDateParts
    {
        public int Year;
        public int Month;
        public int Day;
        /// <summary>0 = dimanche … 6 = samedi.</summary>
        public int DayOfWeek;
        public int Hour;
        public int Minute;
        public int Second;
        /// <summary>Millisecondes (0 à 999).</summary>
        public int Millisecond;
        /// <summary>Nombre entier de jours (partie entière du numéro de série, après arrondi de l'heure).</summary>
        public long Days;
        /// <summary>Durée totale en millisecondes (pour les formats de durée [h], [m], [s]).</summary>
        public long TotalMilliseconds;
    }

    /// <summary>
    /// Numéros de série Excel ↔ dates, calendriers 1900 (avec le 29 février 1900 fictif hérité de Lotus 1-2-3)
    /// et 1904.
    /// </summary>
    public static class ExcelDates
    {
        /// <summary>Dernier numéro de série valide + 1 (1er janvier 10000) dans le calendrier 1900.</summary>
        public const double MaxSerial1900 = 2958466;
        public const double MaxSerial1904 = 2957004;

        private static readonly DateTime Base1900 = new DateTime(1899, 12, 30);
        private static readonly DateTime Base1904 = new DateTime(1904, 1, 1);

        public static bool IsValidSerial(double serial, bool date1904)
        {
            if (double.IsNaN(serial) || double.IsInfinity(serial) || serial < 0) return false;
            return serial < (date1904 ? MaxSerial1904 : MaxSerial1900);
        }

        /// <summary>
        /// Décompose un numéro de série. <paramref name="secondDecimals"/> = nombre de décimales de secondes
        /// affichées (0 à 3) : l'heure est arrondie à cette précision, comme le fait Excel à l'affichage ;
        /// -1 = heure non affichée (arrondi à la milliseconde seulement, pour absorber les imprécisions de calcul).
        /// </summary>
        public static bool TryGetParts(double serial, bool date1904, int secondDecimals, out ExcelDateParts parts)
        {
            parts = default(ExcelDateParts);
            if (!IsValidSerial(serial, date1904)) return false;

            // Millisecondes depuis l'origine, arrondies au millième puis à la précision affichée.
            double totalMs = Math.Round(serial * 86400000.0, MidpointRounding.AwayFromZero);
            long unit = secondDecimals < 0 || secondDecimals >= 3 ? 1 : secondDecimals == 2 ? 10 : secondDecimals == 1 ? 100 : 1000;
            long ms = (long)totalMs;
            ms = (ms + unit / 2) / unit * unit;

            long days = ms / 86400000L;
            long timeOfDay = ms - days * 86400000L;
            parts.Days = days;
            parts.TotalMilliseconds = ms;
            parts.Hour = (int)(timeOfDay / 3600000L);
            parts.Minute = (int)(timeOfDay / 60000L % 60);
            parts.Second = (int)(timeOfDay / 1000L % 60);
            parts.Millisecond = (int)(timeOfDay % 1000);

            if (date1904)
            {
                parts.DayOfWeek = (int)((days + 5) % 7); // 1er janvier 1904 : vendredi
                if (days >= (long)MaxSerial1904) return false;
                var date = Base1904.AddDays(days);
                parts.Year = date.Year;
                parts.Month = date.Month;
                parts.Day = date.Day;
                return true;
            }

            parts.DayOfWeek = (int)((days + 6) % 7); // série 0 (« 0 janvier 1900 ») : samedi
            if (days >= (long)MaxSerial1900) return false;
            if (days == 0)
            {
                parts.Year = 1900;
                parts.Month = 1;
                parts.Day = 0;
            }
            else if (days == 60)
            {
                parts.Year = 1900;
                parts.Month = 2;
                parts.Day = 29; // jour fictif conservé par Excel
            }
            else
            {
                var date = Base1900.AddDays(days < 60 ? days + 1 : days);
                parts.Year = date.Year;
                parts.Month = date.Month;
                parts.Day = date.Day;
            }
            return true;
        }

        /// <summary>Date → numéro de série Excel.</summary>
        public static double ToSerial(DateTime date, bool date1904)
        {
            if (date1904) return (date - Base1904).TotalDays;
            double serial = (date - Base1900).TotalDays;
            if (serial < 61) serial -= 1; // avant le 1er mars 1900 : décalage du 29 février fictif
            return serial;
        }

        /// <summary>Lit une valeur de cellule de type « d » (ISO 8601 : « 2024-03-15T10:30:00 », « 10:30:00 »).</summary>
        public static bool TryParseIso(string text, bool date1904, out double serial)
        {
            serial = 0;
            if (string.IsNullOrEmpty(text)) return false;
            string s = text.Trim();
            DateTime date;
            string[] timeOnly = { "HH:mm:ss", "HH:mm:ss.FFFFFFF", "HH:mm", "THH:mm:ss", "THH:mm:ss.FFFFFFF" };
            if (DateTime.TryParseExact(s, timeOnly, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out date))
            {
                serial = date.TimeOfDay.TotalDays;
                return true;
            }
            if (!DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces, out date)) return false;
            if (date.Kind == DateTimeKind.Local) date = date.ToUniversalTime();
            serial = ToSerial(date, date1904);
            return serial >= 0;
        }
    }
}
