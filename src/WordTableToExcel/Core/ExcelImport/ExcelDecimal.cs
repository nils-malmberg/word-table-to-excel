using System;
using System.Globalization;
using System.Text;

namespace WordTableToExcel.Core.ExcelImport
{
    /// <summary>
    /// Nombre positif en représentation décimale exacte, limité comme dans Excel à 15 chiffres significatifs :
    /// valeur = 0,D1D2…Dn × 10^Exposant (D1 ≠ 0 ; zéro = aucun chiffre).
    /// Les arrondis se font sur les chiffres décimaux (au plus proche, 5 vers le haut), ce qui reproduit
    /// l'affichage d'Excel (1,005 au format 0,00 → 1,01) quel que soit le moteur .NET utilisé.
    /// </summary>
    internal sealed class ExcelDecimal
    {
        public const int ExcelPrecision = 15;

        public static readonly ExcelDecimal Zero = new ExcelDecimal(string.Empty, 0);

        private ExcelDecimal(string digits, int exponent)
        {
            Digits = digits;
            Exponent = digits.Length == 0 ? 0 : exponent;
        }

        /// <summary>Chiffres significatifs, sans zéro de tête ni de fin.</summary>
        public readonly string Digits;
        /// <summary>Position de la virgule : nombre de chiffres de la partie entière si positif.</summary>
        public readonly int Exponent;

        public bool IsZero
        {
            get { return Digits.Length == 0; }
        }

        /// <summary>Partie entière de log10 (123 → 2, 0,05 → -2). Non défini pour zéro.</summary>
        public int Log10Floor
        {
            get { return Exponent - 1; }
        }

        /// <summary>Valeur absolue de <paramref name="value"/>, arrondie à 15 chiffres significatifs.</summary>
        public static ExcelDecimal FromDouble(double value)
        {
            value = Math.Abs(value);
            if (value == 0 || double.IsNaN(value) || double.IsInfinity(value)) return Zero;
            // « E16 » : 17 chiffres significatifs, exacts sur .NET Framework comme sur .NET moderne.
            string s = value.ToString("E16", CultureInfo.InvariantCulture);
            int e = s.IndexOf('E');
            string digits = s.Substring(0, 1) + s.Substring(2, e - 2);
            int exponent = int.Parse(s.Substring(e + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture) + 1;
            return Normalize(digits, exponent).RoundSignificant(ExcelPrecision);
        }

        /// <summary>Lit une écriture décimale invariante (« 123.45 », « 1E-5 »).</summary>
        public static ExcelDecimal Parse(string text)
        {
            double d;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return Zero;
            return FromDouble(d);
        }

        /// <summary>Multiplie par 10^<paramref name="power"/> (exact).</summary>
        public ExcelDecimal Shift(int power)
        {
            return IsZero || power == 0 ? this : new ExcelDecimal(Digits, Exponent + power);
        }

        public ExcelDecimal RoundSignificant(int count)
        {
            return RoundAt(count);
        }

        /// <summary>Arrondit à <paramref name="decimals"/> chiffres après la virgule.</summary>
        public ExcelDecimal RoundDecimals(int decimals)
        {
            return RoundAt(Exponent + decimals);
        }

        /// <summary>Tronque (sans arrondir) à <paramref name="decimals"/> chiffres après la virgule.</summary>
        public ExcelDecimal TruncateDecimals(int decimals)
        {
            int keep = Exponent + decimals;
            if (IsZero || keep >= Digits.Length) return this;
            if (keep <= 0) return Zero;
            return Normalize(Digits.Substring(0, keep), Exponent);
        }

        /// <summary>Chiffres de la partie entière (chaîne vide si elle vaut zéro).</summary>
        public string IntegerDigits()
        {
            if (IsZero || Exponent <= 0) return string.Empty;
            if (Exponent >= Digits.Length) return Digits + new string('0', Exponent - Digits.Length);
            return Digits.Substring(0, Exponent);
        }

        /// <summary>Les <paramref name="count"/> premiers chiffres après la virgule (complétés par des zéros).</summary>
        public string FractionDigits(int count)
        {
            if (count <= 0) return string.Empty;
            var sb = new StringBuilder(count);
            for (int i = 0; i < count; i++)
            {
                int index = Exponent + i; // position dans Digits du i-ème chiffre après la virgule
                sb.Append(index >= 0 && index < Digits.Length ? Digits[index] : '0');
            }
            return sb.ToString();
        }

        /// <summary>Nombre de chiffres après la virgule nécessaires pour écrire la valeur exactement.</summary>
        public int DecimalCount
        {
            get { return IsZero ? 0 : Math.Max(0, Digits.Length - Exponent); }
        }

        public double ToDouble()
        {
            if (IsZero) return 0;
            return double.Parse("0." + Digits + "E" + Exponent.ToString(CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        /// <summary>Écriture décimale simple, pour les tests et le journal (« 1234.5 », « 0.001 »).</summary>
        public override string ToString()
        {
            if (IsZero) return "0";
            string integer = IntegerDigits();
            string fraction = FractionDigits(DecimalCount);
            return (integer.Length == 0 ? "0" : integer) + (fraction.Length > 0 ? "." + fraction : string.Empty);
        }

        private ExcelDecimal RoundAt(int keep)
        {
            if (IsZero || keep >= Digits.Length) return this;
            if (keep < 0) return Zero;
            if (keep == 0) return Digits[0] >= '5' ? new ExcelDecimal("1", Exponent + 1) : Zero;

            bool up = Digits[keep] >= '5';
            char[] kept = Digits.Substring(0, keep).ToCharArray();
            int exponent = Exponent;
            if (up)
            {
                int i = kept.Length - 1;
                while (i >= 0)
                {
                    if (kept[i] == '9')
                    {
                        kept[i] = '0';
                        i--;
                        continue;
                    }
                    kept[i]++;
                    break;
                }
                if (i < 0)
                {
                    // 999 → 1000 : un chiffre de plus devant.
                    return new ExcelDecimal("1", exponent + 1);
                }
            }
            return Normalize(new string(kept), exponent);
        }

        private static ExcelDecimal Normalize(string digits, int exponent)
        {
            int start = 0;
            while (start < digits.Length && digits[start] == '0')
            {
                start++;
                exponent--;
            }
            int end = digits.Length;
            while (end > start && digits[end - 1] == '0') end--;
            if (end <= start) return Zero;
            return new ExcelDecimal(digits.Substring(start, end - start), exponent);
        }
    }
}
