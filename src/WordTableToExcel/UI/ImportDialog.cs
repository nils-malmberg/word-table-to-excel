using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using WordTableToExcel.Core.ExcelImport;
using WordTableToExcel.Import;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.UI
{
    /// <summary>Options choisies dans la boîte de dialogue d'import.</summary>
    internal sealed class ImportChoices
    {
        public bool AddCaption;
        public bool CaptionBelow;
        public bool FitToPage;
        public bool SkipHidden;
        public bool AddGridlines;
    }

    /// <summary>
    /// Boîte de dialogue d'import : une ligne par feuille du classeur (à importer ou non, plage, texte de la
    /// légende), options de légende et de mise en page.
    /// </summary>
    internal sealed class ImportDialog : Form
    {
        private const int ColumnImport = 0;
        private const int ColumnSheet = 1;
        private const int ColumnRange = 2;
        private const int ColumnSize = 3;
        private const int ColumnCaption = 4;
        private const int ColumnRemark = 5;

        private readonly IList<ImportCandidate> _candidates;
        private readonly DataGridView _grid;
        private readonly CheckBox _addCaption;
        private readonly ComboBox _captionPosition;
        private readonly CheckBox _fitToPage;
        private readonly CheckBox _skipHidden;
        private readonly CheckBox _gridlines;

        public ImportDialog(string fileName, IList<ImportCandidate> candidates, Settings settings, bool captionBelow)
        {
            _candidates = candidates;

            Text = "Importer des tableaux depuis Excel";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            const int width = 760;
            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(14, 12, 14, 12),
                Location = new Point(0, 0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, width));

            root.Controls.Add(new Label
            {
                Text = "Classeur : " + fileName,
                Font = new Font(Font, FontStyle.Bold),
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Margin = new Padding(0, 0, 0, 4)
            });
            root.Controls.Add(new Label
            {
                Text = "Cochez les feuilles à importer (case « Tout » : toutes à la fois) : chacune devient un tableau Word, inséré à "
                     + "l'emplacement du curseur, avec sa mise en forme et les valeurs telles qu'Excel les affiche.",
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Margin = new Padding(0, 0, 0, 8)
            });

            _grid = BuildGrid(width);
            root.Controls.Add(_grid);
            root.Controls.Add(new Label
            {
                Text = "Plage : modifiable si nécessaire (par exemple A3:F20).   Légende : texte placé après « Tableau N : » ; "
                     + "s'il est vide, « " + Word.WordTableInserter.Placeholder + " » est inséré, surligné en jaune, à compléter.",
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 4, 0, 10)
            });

            // --- Options
            var options = new TableLayoutPanel
            {
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 0, 0, 8)
            };
            _addCaption = new CheckBox { Text = "Ajouter une légende numérotée à chaque tableau", AutoSize = true, Checked = settings.ImportAddCaption, Margin = new Padding(0, 4, 8, 2) };
            _captionPosition = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Margin = new Padding(0, 1, 0, 2) };
            _captionPosition.Items.Add("au-dessus du tableau");
            _captionPosition.Items.Add("sous le tableau");
            _captionPosition.SelectedIndex = captionBelow ? 1 : 0;
            _addCaption.CheckedChanged += (s, e) => _captionPosition.Enabled = _addCaption.Checked;
            _captionPosition.Enabled = _addCaption.Checked;
            options.Controls.Add(_addCaption, 0, 0);
            options.Controls.Add(_captionPosition, 1, 0);

            _fitToPage = new CheckBox { Text = "Réduire les tableaux trop larges pour la page (sans couper les nombres)", AutoSize = true, Checked = settings.ImportFitToPage, Margin = new Padding(0, 4, 0, 2) };
            _skipHidden = new CheckBox { Text = "Ignorer les lignes et colonnes masquées dans Excel", AutoSize = true, Checked = settings.ImportSkipHidden, Margin = new Padding(0, 2, 0, 2) };
            _gridlines = new CheckBox { Text = "Ajouter un quadrillage gris aux cellules sans bordure", AutoSize = true, Checked = settings.ImportGridlines, Margin = new Padding(0, 2, 0, 2) };
            _skipHidden.CheckedChanged += (s, e) => RefreshSizes();
            options.Controls.Add(_fitToPage, 0, 1);
            options.SetColumnSpan(_fitToPage, 2);
            options.Controls.Add(_skipHidden, 0, 2);
            options.SetColumnSpan(_skipHidden, 2);
            options.Controls.Add(_gridlines, 0, 3);
            options.SetColumnSpan(_gridlines, 2);
            root.Controls.Add(Section("Options"));
            root.Controls.Add(options);

            // --- Boutons
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 6, 0, 0)
            };
            var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
            var import = new Button { Text = "Importer", AutoSize = true, MinimumSize = new Size(88, 0) };
            import.Click += OnImportClick;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(import);
            root.Controls.Add(buttons);

            Controls.Add(root);
            AcceptButton = import;
            CancelButton = cancel;

            FillGrid();
        }

        public ImportChoices Choices { get; private set; }

        private static Label Section(string text)
        {
            return new Label { Text = text, AutoSize = true, Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold), Margin = new Padding(0, 0, 0, 4) };
        }

        private DataGridView BuildGrid(int width)
        {
            var grid = new DataGridView
            {
                Width = width,
                Height = 230,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                BackgroundColor = SystemColors.Window,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                Margin = new Padding(0)
            };
            // En-tête de la colonne : case « Tout » (dessinée ci-dessous) qui coche ou décoche toutes les feuilles.
            grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Tout", Width = 62, ToolTipText = "Cocher ou décocher toutes les feuilles" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Feuille", Width = 130, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Plage", Width = 90 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Taille", Width = 80, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Texte de la légende", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 150 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Remarque", Width = 170, ReadOnly = true });
            foreach (DataGridViewColumn c in grid.Columns) c.SortMode = DataGridViewColumnSortMode.NotSortable;
            grid.CellEndEdit += (s, e) =>
            {
                if (e.ColumnIndex == ColumnRange)
                {
                    RefreshRow(e.RowIndex);
                    grid.InvalidateCell(ColumnImport, -1);
                }
            };
            grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                // La case à cocher prend effet immédiatement.
                if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell) grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            grid.DataError += (s, e) => e.ThrowException = false;
            grid.CellValueChanged += (s, e) =>
            {
                if (e.ColumnIndex == ColumnImport) grid.InvalidateCell(ColumnImport, -1);
            };
            grid.CellPainting += PaintSelectAllHeader;
            grid.ColumnHeaderMouseClick += (s, e) =>
            {
                if (e.ColumnIndex == ColumnImport && e.Button == MouseButtons.Left) ToggleAll();
            };
            return grid;
        }

        // ------------------------------------------------------------------ case « Tout »

        /// <summary>Feuille cochable par « Tout » : importable, avec une plage valide qui tient dans un tableau Word.</summary>
        private bool Eligible(int index)
        {
            var c = _candidates[index];
            if (!c.CanImport) return false;
            CellRange range;
            string text = Convert.ToString(_grid.Rows[index].Cells[ColumnRange].Value, CultureInfo.CurrentCulture) ?? string.Empty;
            if (!CellRange.TryParse(text, out range)) return false;
            int rows, columns;
            Count(c, range, out rows, out columns);
            return columns <= SheetConverter.MaxWordColumns && (long)rows * columns <= SheetConverter.MaxCells;
        }

        private bool IsChecked(int index)
        {
            return _candidates[index].CanImport && Convert.ToBoolean(_grid.Rows[index].Cells[ColumnImport].Value ?? false, CultureInfo.InvariantCulture);
        }

        /// <summary>État de la case « Tout » : cochée si toutes les feuilles cochables le sont, partielle si certaines seulement.</summary>
        internal CheckState SelectAllState()
        {
            if (_grid == null) return CheckState.Unchecked;
            int eligible = 0, eligibleChecked = 0, anyChecked = 0;
            for (int i = 0; i < _candidates.Count && i < _grid.Rows.Count; i++)
            {
                bool isChecked = IsChecked(i);
                if (isChecked) anyChecked++;
                if (!Eligible(i)) continue;
                eligible++;
                if (isChecked) eligibleChecked++;
            }
            if (anyChecked == 0) return CheckState.Unchecked;
            return eligible > 0 && eligibleChecked == eligible ? CheckState.Checked : CheckState.Indeterminate;
        }

        /// <summary>Clic sur « Tout » : tout décocher si tout est coché, sinon cocher toutes les feuilles cochables.</summary>
        internal void ToggleAll()
        {
            _grid.EndEdit();
            // La cellule en cours d'édition garderait son ancienne valeur affichée.
            if (_grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == ColumnImport)
            {
                _grid.CurrentCell = _grid.Rows[_grid.CurrentCell.RowIndex].Cells[ColumnSheet];
            }
            bool uncheck = SelectAllState() == CheckState.Checked;
            for (int i = 0; i < _candidates.Count; i++)
            {
                if (!_candidates[i].CanImport) continue;
                if (uncheck) _grid.Rows[i].Cells[ColumnImport].Value = false;
                else if (Eligible(i)) _grid.Rows[i].Cells[ColumnImport].Value = true;
            }
            _grid.InvalidateCell(ColumnImport, -1);
        }

        private void PaintSelectAllHeader(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex != ColumnImport) return;
            e.PaintBackground(e.CellBounds, false);
            var state = SelectAllState();
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

        private void FillGrid()
        {
            foreach (var c in _candidates)
            {
                int index = _grid.Rows.Add(c.Selected && c.CanImport, c.Info.Name, c.Range.HasValue ? c.Range.Value.ToString() : string.Empty,
                    string.Empty, c.CaptionTitle ?? string.Empty, string.Empty);
                var row = _grid.Rows[index];
                if (!c.CanImport)
                {
                    row.ReadOnly = true;
                    row.DefaultCellStyle.ForeColor = SystemColors.GrayText;
                }
                RefreshRow(index);
            }
        }

        private void RefreshSizes()
        {
            for (int i = 0; i < _grid.Rows.Count; i++) RefreshRow(i);
            _grid.InvalidateCell(ColumnImport, -1);
        }

        /// <summary>Met à jour la taille et la remarque d'une ligne après modification de la plage.</summary>
        private void RefreshRow(int index)
        {
            if (index < 0 || index >= _candidates.Count) return;
            var c = _candidates[index];
            var row = _grid.Rows[index];
            if (!c.CanImport)
            {
                row.Cells[ColumnRemark].Value = c.Problem ?? "Rien à importer";
                return;
            }
            CellRange range;
            string text = Convert.ToString(row.Cells[ColumnRange].Value, CultureInfo.CurrentCulture) ?? string.Empty;
            if (!CellRange.TryParse(text, out range))
            {
                row.Cells[ColumnSize].Value = string.Empty;
                row.Cells[ColumnRemark].Value = "Plage invalide";
                row.Cells[ColumnRange].Style.ForeColor = Color.Firebrick;
                return;
            }
            row.Cells[ColumnRange].Style.ForeColor = SystemColors.WindowText;
            int rows, columns;
            Count(c, range, out rows, out columns);
            row.Cells[ColumnSize].Value = rows.ToString(CultureInfo.CurrentCulture) + " × " + columns.ToString(CultureInfo.CurrentCulture);

            var remarks = new List<string>();
            if (!c.Info.IsVisible) remarks.Add("Feuille masquée");
            if (columns > SheetConverter.MaxWordColumns) remarks.Add("Trop de colonnes (max. " + SheetConverter.MaxWordColumns + ")");
            if ((long)rows * columns > SheetConverter.MaxCells) remarks.Add("Plage trop grande");
            if (c.DetectedCaption != null && c.DetectedRange.HasValue && range.Equals(c.DetectedRange.Value)) remarks.Add("Légende trouvée dans la feuille");
            if (c.Sheet.HasDrawings) remarks.Add("Images non importées");
            row.Cells[ColumnRemark].Value = string.Join(" ; ", remarks.ToArray());
        }

        private void Count(ImportCandidate c, CellRange range, out int rows, out int columns)
        {
            bool skip = _skipHidden.Checked;
            rows = 0;
            columns = 0;
            for (int r = range.FirstRow; r <= range.LastRow && rows <= SheetConverter.MaxCells; r++)
            {
                if (!skip || !c.Sheet.IsRowHidden(r)) rows++;
            }
            for (int col = range.FirstColumn; col <= range.LastColumn && columns <= SheetConverter.MaxWordColumns + 1; col++)
            {
                if (!skip || !c.Sheet.IsColumnHidden(col)) columns++;
            }
        }

        private void OnImportClick(object sender, EventArgs e)
        {
            _grid.EndEdit();
            int selected = 0;
            for (int i = 0; i < _candidates.Count; i++)
            {
                var c = _candidates[i];
                var row = _grid.Rows[i];
                bool check = c.CanImport && Convert.ToBoolean(row.Cells[ColumnImport].Value ?? false, CultureInfo.InvariantCulture);
                if (!check) continue;

                CellRange range;
                string text = Convert.ToString(row.Cells[ColumnRange].Value, CultureInfo.CurrentCulture) ?? string.Empty;
                if (!CellRange.TryParse(text, out range))
                {
                    Messages.Warning(this, "La plage indiquée pour la feuille « " + c.Info.Name + " » n'est pas valide : « " + text + " ».\n\n"
                        + "Indiquez une plage de la forme A1:F20.");
                    _grid.CurrentCell = row.Cells[ColumnRange];
                    return;
                }
                int rows, columns;
                Count(c, range, out rows, out columns);
                if (columns > SheetConverter.MaxWordColumns)
                {
                    Messages.Warning(this, "La plage de la feuille « " + c.Info.Name + " » compte plus de " + SheetConverter.MaxWordColumns
                        + " colonnes, le maximum d'un tableau Word.\n\nIndiquez une plage plus étroite ou décochez cette feuille.");
                    _grid.CurrentCell = row.Cells[ColumnRange];
                    return;
                }
                if ((long)rows * columns > SheetConverter.MaxCells)
                {
                    Messages.Warning(this, "La plage de la feuille « " + c.Info.Name + " » est trop grande pour un tableau Word ("
                        + rows + " lignes × " + columns + " colonnes).\n\nIndiquez une plage plus petite ou décochez cette feuille.");
                    _grid.CurrentCell = row.Cells[ColumnRange];
                    return;
                }
                selected++;
            }
            if (selected == 0)
            {
                Messages.Info(this, "Cochez au moins une feuille à importer.");
                return;
            }

            for (int i = 0; i < _candidates.Count; i++)
            {
                var c = _candidates[i];
                var row = _grid.Rows[i];
                c.Selected = c.CanImport && Convert.ToBoolean(row.Cells[ColumnImport].Value ?? false, CultureInfo.InvariantCulture);
                if (!c.Selected) continue;
                CellRange range;
                CellRange.TryParse(Convert.ToString(row.Cells[ColumnRange].Value, CultureInfo.CurrentCulture), out range);
                c.Range = range;
                c.CaptionTitle = (Convert.ToString(row.Cells[ColumnCaption].Value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
            }
            Choices = new ImportChoices
            {
                AddCaption = _addCaption.Checked,
                CaptionBelow = _captionPosition.SelectedIndex == 1,
                FitToPage = _fitToPage.Checked,
                SkipHidden = _skipHidden.Checked,
                AddGridlines = _gridlines.Checked
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
