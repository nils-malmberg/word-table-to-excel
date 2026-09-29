using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Import
{
    /// <summary>Mesure la largeur réelle d'un texte (GDI), pour dimensionner les colonnes sans couper les valeurs.</summary>
    internal sealed class TextMeasurer : IDisposable
    {
        private readonly Dictionary<string, Font> _fonts = new Dictionary<string, Font>(StringComparer.Ordinal);
        private readonly float _dpi;
        private bool _failed;

        public TextMeasurer()
        {
            try
            {
                using (var g = Graphics.FromHwnd(IntPtr.Zero)) _dpi = g.DpiX;
            }
            catch (Exception)
            {
                _dpi = 96;
            }
            if (_dpi <= 0) _dpi = 96;
        }

        /// <summary>Largeur du texte en points (estimation si la police est indisponible).</summary>
        public double Measure(string text, RunFormat format)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            if (_failed) return SheetConverter.EstimateTextWidth(text, format);
            try
            {
                var font = GetFont(format);
                double widest = 0;
                foreach (var line in text.Split('\n'))
                {
                    if (line.Length == 0) continue;
                    var size = TextRenderer.MeasureText(line, font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
                    widest = Math.Max(widest, size.Width * 72.0 / _dpi);
                }
                return widest;
            }
            catch (Exception)
            {
                _failed = true;
                return SheetConverter.EstimateTextWidth(text, format);
            }
        }

        private Font GetFont(RunFormat format)
        {
            string name = string.IsNullOrEmpty(format.FontName) ? "Calibri" : format.FontName;
            float size = (float)(format.Size ?? 11);
            var style = FontStyle.Regular;
            if (format.Bold) style |= FontStyle.Bold;
            if (format.Italic) style |= FontStyle.Italic;
            string key = name + "|" + size + "|" + (int)style;
            Font font;
            if (!_fonts.TryGetValue(key, out font))
            {
                font = new Font(name, size, style, GraphicsUnit.Point);
                _fonts[key] = font;
            }
            return font;
        }

        public void Dispose()
        {
            foreach (var f in _fonts.Values) f.Dispose();
            _fonts.Clear();
        }
    }
}
