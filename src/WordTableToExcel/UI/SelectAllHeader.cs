using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace WordTableToExcel.UI
{
    /// <summary>
    /// Case « Tout » dessinée dans l'en-tête d'une colonne de cases à cocher d'une grille : un clic coche toutes les
    /// lignes concernées, ou les décoche toutes si elles le sont déjà ; état partiel si certaines seulement.
    /// Les cases de la colonne prennent effet dès le clic. Utilisée par les fenêtres d'export et d'import.
    /// </summary>
    internal sealed class SelectAllHeader
    {
        private readonly DataGridView _grid;
        private readonly int _column;
        private readonly int _parkColumn;
        private readonly Func<int, bool> _selectable;
        private readonly Func<int, bool> _eligible;

        /// <param name="column">Colonne des cases à cocher.</param>
        /// <param name="parkColumn">Colonne en lecture seule où placer la cellule courante pendant « Tout ».</param>
        /// <param name="selectable">La ligne peut-elle être cochée ?</param>
        /// <param name="eligible">La ligne est-elle cochée par « Tout » ?</param>
        public SelectAllHeader(DataGridView grid, int column, int parkColumn, Func<int, bool> selectable, Func<int, bool> eligible)
        {
            _grid = grid;
            _column = column;
            _parkColumn = parkColumn;
            _selectable = selectable;
            _eligible = eligible;

            grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                // La case à cocher prend effet immédiatement (sans attendre de quitter la cellule).
                if (grid.IsCurrentCellDirty && grid.CurrentCell != null && grid.CurrentCell.ColumnIndex == column)
                {
                    grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
            grid.CellValueChanged += (s, e) =>
            {
                if (e.ColumnIndex == column) Invalidate();
            };
            grid.CellPainting += Paint;
            grid.ColumnHeaderMouseClick += (s, e) =>
            {
                if (e.ColumnIndex == column && e.Button == MouseButtons.Left) Toggle();
            };
        }

        public bool IsChecked(int row)
        {
            return _selectable(row) && Convert.ToBoolean(_grid.Rows[row].Cells[_column].Value ?? false, CultureInfo.InvariantCulture);
        }

        /// <summary>Cochée si toutes les lignes concernées le sont, partielle si certaines seulement.</summary>
        public CheckState State
        {
            get
            {
                int eligible = 0, eligibleChecked = 0, anyChecked = 0;
                for (int i = 0; i < _grid.Rows.Count; i++)
                {
                    bool isChecked = IsChecked(i);
                    if (isChecked) anyChecked++;
                    if (!_selectable(i) || !_eligible(i)) continue;
                    eligible++;
                    if (isChecked) eligibleChecked++;
                }
                if (anyChecked == 0) return CheckState.Unchecked;
                return eligible > 0 && eligibleChecked == eligible ? CheckState.Checked : CheckState.Indeterminate;
            }
        }

        /// <summary>Clic sur « Tout » : tout décocher si tout est coché, sinon cocher toutes les lignes concernées.</summary>
        public void Toggle()
        {
            _grid.EndEdit();
            // La cellule en cours d'édition garderait son ancienne valeur affichée.
            if (_grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == _column)
            {
                _grid.CurrentCell = _grid.Rows[_grid.CurrentCell.RowIndex].Cells[_parkColumn];
            }
            bool uncheck = State == CheckState.Checked;
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                if (!_selectable(i)) continue;
                if (uncheck) _grid.Rows[i].Cells[_column].Value = false;
                else if (_eligible(i)) _grid.Rows[i].Cells[_column].Value = true;
            }
            Invalidate();
        }

        public void Invalidate()
        {
            _grid.InvalidateCell(_column, -1);
        }

        private void Paint(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex != _column) return;
            e.PaintBackground(e.CellBounds, false);
            var state = State;
            var glyphState = state == CheckState.Checked ? CheckBoxState.CheckedNormal
                : state == CheckState.Indeterminate ? CheckBoxState.MixedNormal : CheckBoxState.UncheckedNormal;
            Size glyph = CheckBoxRenderer.GetGlyphSize(e.Graphics, glyphState);
            var location = new Point(e.CellBounds.Left + 6, e.CellBounds.Top + (e.CellBounds.Height - glyph.Height) / 2);
            CheckBoxRenderer.DrawCheckBox(e.Graphics, location, glyphState);
            int textLeft = location.X + glyph.Width + 4;
            var textBounds = new Rectangle(textLeft, e.CellBounds.Top, Math.Max(0, e.CellBounds.Right - textLeft - 2), e.CellBounds.Height);
            TextRenderer.DrawText(e.Graphics, "Tout", e.CellStyle.Font, textBounds, e.CellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            e.Handled = true;
        }
    }
}
