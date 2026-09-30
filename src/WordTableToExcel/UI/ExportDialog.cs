using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Export;

namespace WordTableToExcel.UI
{
    /// <summary>
    /// Boîte de dialogue principale : position des légendes (détection automatique, dont le résultat est affiché,
    /// ou choix de l'utilisateur), choix entre l'option A (tableaux avec légende) et l'option B (tous les tableaux),
    /// liste des tableaux proposés (cases à cocher, pages de début et de fin, feuille Excel créée, légende), options.
    /// </summary>
    internal sealed class ExportDialog : Form
    {
        private const int ColumnExport = 0;
        private const int ColumnIndex = 1;
        private const int ColumnStartPage = 2;
        private const int ColumnEndPage = 3;
        private const int ColumnSheet = 4;
        private const int ColumnCaption = 5;

        private readonly IList<TableEntry> _tables;
        private readonly int _deletedTables;
        private readonly Label _summary;
        private readonly ComboBox _captionSide;
        private readonly Label _captionedHint;
        private readonly RadioButton _captionedOnly;
        private readonly RadioButton _allTables;
        private readonly CheckBox _includeCaption;
        private readonly CheckBox _includeSummary;
        private readonly CheckBox _convertNumbers;
        private readonly DataGridView _preview;
        private readonly SelectAllHeader _selectAll;
        private readonly Button _exportButton;
        private readonly Label _previewTitle;
        /// <summary>Tableaux décochés (rang dans le document), conservés quand on passe de A à B.</summary>
        private readonly HashSet<int> _excluded = new HashSet<int>();
        /// <summary>Tableau affiché sur chaque ligne de la liste.</summary>
        private readonly List<TableEntry> _rowTables = new List<TableEntry>();
        private bool _filling;

        /// <param name="detected">Position des légendes détectée dans le document (affichée pour le choix « Automatique »).</param>
        /// <param name="captionPosition">Position choisie (None : automatique), avec laquelle les légendes de <paramref name="tables"/> ont été attribuées.</param>
        /// <param name="deletedTables">Tableaux supprimés en suivi des modifications, écartés (signalés dans le résumé).</param>
        /// <param name="includeSummary">Feuille « Sommaire » en tête du classeur (à partir de deux feuilles).</param>
        public ExportDialog(string documentName, IList<TableEntry> tables, CaptionPosition detected, CaptionPosition captionPosition, bool allTables,
            bool includeCaption, bool convertNumbers, bool includeSummary, int deletedTables = 0)
        {
            _tables = tables;
            _deletedTables = deletedTables;
            int total = tables.Count;

            Text = "Exporter les tableaux vers Excel";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            const int width = 640;
            // Une seule colonne, sans conteneur imbriqué redimensionné automatiquement :
            // mise en page stable quelle que soit la résolution (DPI) ou la police système.
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
                Text = "Document : " + documentName,
                Font = new Font(Font, FontStyle.Bold),
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Margin = new Padding(0, 0, 0, 4)
            });

            _summary = new Label { AutoSize = true, MaximumSize = new Size(width, 0), Margin = new Padding(0, 0, 0, 4) };
            root.Controls.Add(_summary);

            // --- Position des légendes : le choix « Automatique » affiche ce qui a été détecté, pour que l'utilisateur
            // voie tout de suite si la détection se trompe et impose la bonne position.
            var sideRow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(0, 4, 0, 0) };
            sideRow.Controls.Add(new Label { Text = "Position des légendes :", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
            _captionSide = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 340, Margin = new Padding(0, 2, 0, 0) };
            _captionSide.Items.Add("Automatique \u2014 détectée : " + (detected == CaptionPosition.Below ? "au-dessous des tableaux" : "au-dessus des tableaux"));
            _captionSide.Items.Add("Au-dessus des tableaux");
            _captionSide.Items.Add("Au-dessous des tableaux");
            _captionSide.SelectedIndex = captionPosition == CaptionPosition.Above ? 1 : captionPosition == CaptionPosition.Below ? 2 : 0;
            _captionSide.Enabled = tables.Any(t => t.CaptionAbove != null || t.CaptionBelow != null);
            sideRow.Controls.Add(_captionSide);
            root.Controls.Add(sideRow);
            root.Controls.Add(Hint("Si la position détectée n'est pas la bonne, choisissez-la : les légendes de la liste sont réattribuées aussitôt.", width));

            // --- Choix A / B
            root.Controls.Add(Section("Tableaux à exporter"));
            _captionedOnly = new RadioButton
            {
                AutoSize = true,
                Margin = new Padding(8, 2, 0, 0)
            };
            root.Controls.Add(_captionedOnly);
            _captionedHint = Hint(string.Empty, width);
            root.Controls.Add(_captionedHint);

            _allTables = new RadioButton
            {
                Text = string.Format(CultureInfo.CurrentCulture, "B \u2014 Tous les tableaux du document ({0})", total),
                AutoSize = true,
                Margin = new Padding(8, 6, 0, 0)
            };
            root.Controls.Add(_allTables);
            root.Controls.Add(Hint("Les tableaux sans légende sont nommés Tableau_1, Tableau_2… d'après leur rang dans le document.", width));

            // --- Aperçu
            _previewTitle = Section(string.Empty);
            root.Controls.Add(_previewTitle);
            _preview = BuildGrid(width);
            _selectAll = new SelectAllHeader(_preview, ColumnExport, ColumnIndex, i => i < _rowTables.Count, i => true);
            root.Controls.Add(_preview);
            root.Controls.Add(new Label
            {
                Text = "Décochez les tableaux à ne pas exporter (case « Tout » : tous à la fois). Les pages aident à les repérer dans le document.",
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 3, 0, 0)
            });

            // --- Options
            root.Controls.Add(Section("Options"));
            _includeCaption = new CheckBox
            {
                Text = "Écrire la légende complète en haut de chaque feuille (cellule A1)",
                AutoSize = true,
                Checked = includeCaption,
                Margin = new Padding(8, 2, 0, 0)
            };
            _includeSummary = new CheckBox
            {
                Text = "Ajouter une feuille « Sommaire » en tête du classeur (liste des tableaux, avec un lien vers chaque feuille)",
                AutoSize = true,
                Checked = includeSummary,
                Margin = new Padding(8, 4, 0, 0)
            };
            _convertNumbers = new CheckBox
            {
                Text = "Convertir les nombres en valeurs numériques Excel",
                AutoSize = true,
                Checked = convertNumbers,
                Margin = new Padding(8, 4, 0, 0)
            };
            root.Controls.Add(_includeCaption);
            root.Controls.Add(_includeSummary);
            root.Controls.Add(_convertNumbers);
            root.Controls.Add(Hint("Par exemple « 1 234,50 », « 12,5 % », « 45 € ». Sinon, le contenu des cellules est copié tel quel, sous forme de texte.", width));

            root.Controls.Add(new Label
            {
                Text = "Le document Word est seulement lu : il n'est jamais modifié.",
                ForeColor = SystemColors.GrayText,
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Margin = new Padding(0, 12, 0, 8)
            });

            // --- Boutons
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0)
            };
            var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(95, 28) };
            _exportButton = new Button { Text = "Exporter…", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(95, 28) };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_exportButton);
            root.Controls.Add(buttons);

            Controls.Add(root);
            AcceptButton = _exportButton;
            CancelButton = cancel;

            RefreshCaptionTexts();
            if (!_captionedOnly.Enabled || allTables) _allTables.Checked = true;
            else _captionedOnly.Checked = true;

            _captionedOnly.CheckedChanged += (s, e) => RefreshPreview();
            _allTables.CheckedChanged += (s, e) => RefreshPreview();
            _captionSide.SelectedIndexChanged += (s, e) =>
            {
                ExportPlan.AssignCaptions(_tables, CaptionPositionChoice);
                RefreshCaptionTexts();
                RefreshPreview();
            };
            RefreshPreview();
        }

        /// <summary>Position des légendes choisie (None : automatique).</summary>
        public CaptionPosition CaptionPositionChoice
        {
            get
            {
                switch (_captionSide.SelectedIndex)
                {
                    case 1: return CaptionPosition.Above;
                    case 2: return CaptionPosition.Below;
                    default: return CaptionPosition.None;
                }
            }
        }

        /// <summary>Choix de la position des légendes (autotest).</summary>
        internal void SetCaptionPosition(CaptionPosition position)
        {
            _captionSide.SelectedIndex = position == CaptionPosition.Above ? 1 : position == CaptionPosition.Below ? 2 : 0;
        }

        /// <summary>Textes qui dépendent du nombre de tableaux légendés (après une réattribution des légendes).</summary>
        private void RefreshCaptionTexts()
        {
            int total = _tables.Count;
            int captioned = _tables.Count(t => t.HasCaption);
            string summary = string.Format(CultureInfo.CurrentCulture, "{0} tableau{1} trouvé{2}, dont {3} avec une légende.",
                total, total > 1 ? "x" : string.Empty, total > 1 ? "s" : string.Empty, captioned);
            if (_deletedTables > 0)
            {
                summary += string.Format(CultureInfo.CurrentCulture, " {0} tableau{1} supprimé{2} en suivi des modifications {3} ignoré{2}.",
                    _deletedTables, _deletedTables > 1 ? "x" : string.Empty, _deletedTables > 1 ? "s" : string.Empty, _deletedTables > 1 ? "sont" : "est");
            }
            _summary.Text = summary;
            _captionedOnly.Text = string.Format(CultureInfo.CurrentCulture, "A \u2014 Uniquement les tableaux qui ont une légende ({0})", captioned);
            _captionedOnly.Enabled = captioned > 0;
            _captionedHint.Text = captioned > 0
                ? "Légende : champ SEQ « Tableau », « Table », « Tabla »… ou texte commençant par « Tabl… », juste au-dessus ou au-dessous du tableau. Chaque feuille porte le nom de sa légende."
                : "Aucune légende de tableau n'a été détectée dans ce document.";
            if (captioned == 0 && _captionedOnly.Checked) _allTables.Checked = true;
        }

        public bool AllTables
        {
            get { return _allTables.Checked; }
        }

        public bool IncludeCaptionRow
        {
            get { return _includeCaption.Checked; }
        }

        public bool ConvertNumbers
        {
            get { return _convertNumbers.Checked; }
        }

        public bool IncludeSummary
        {
            get { return _includeSummary.Checked; }
        }

        /// <summary>Rang des tableaux décochés par l'utilisateur.</summary>
        public ICollection<int> ExcludedTables
        {
            get { return new HashSet<int>(_excluded); }
        }

        /// <summary>État de la case « Tout » (autotest).</summary>
        internal CheckState SelectAllState()
        {
            return _selectAll.State;
        }

        /// <summary>Clic sur « Tout » (autotest).</summary>
        internal void ToggleAll()
        {
            _selectAll.Toggle();
        }

        /// <summary>Coche ou décoche un tableau de la liste (autotest).</summary>
        internal void SetTableChecked(int tableIndex, bool check)
        {
            int row = _rowTables.FindIndex(t => t.Index == tableIndex);
            if (row >= 0) _preview.Rows[row].Cells[ColumnExport].Value = check;
        }

        private DataGridView BuildGrid(int width)
        {
            var grid = new DataGridView
            {
                Width = width,
                Height = 190,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = SystemColors.Window,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                Margin = new Padding(0, 2, 0, 0)
            };
            grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Tout", Width = 58, ToolTipText = "Cocher ou décocher tous les tableaux" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "N°", Width = 38, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Page début", Width = 74, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Page fin", Width = 62, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Feuille Excel", Width = 170, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Légende", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 120, ReadOnly = true });
            foreach (DataGridViewColumn c in grid.Columns) c.SortMode = DataGridViewColumnSortMode.NotSortable;
            grid.DataError += (s, e) => e.ThrowException = false;
            grid.CellValueChanged += (s, e) =>
            {
                if (_filling || e.ColumnIndex != ColumnExport || e.RowIndex < 0 || e.RowIndex >= _rowTables.Count) return;
                int index = _rowTables[e.RowIndex].Index;
                if (_selectAll.IsChecked(e.RowIndex)) _excluded.Remove(index);
                else _excluded.Add(index);
                RefreshSheetNames();
            };
            return grid;
        }

        /// <summary>Liste des tableaux proposés par l'option choisie (A ou B), cochés sauf ceux décochés auparavant.</summary>
        private void RefreshPreview()
        {
            _filling = true;
            try
            {
                _preview.Rows.Clear();
                _rowTables.Clear();
                foreach (var table in _tables.Where(t => ExportPlan.IsCandidate(t, AllTables)))
                {
                    string caption = table.HasCaption
                        ? (table.CaptionPosition == CaptionPosition.Below ? "↓ " : "↑ ") + table.Caption
                        : "(sans légende)";
                    int row = _preview.Rows.Add(!_excluded.Contains(table.Index), table.Index.ToString(CultureInfo.CurrentCulture),
                        Page(table.StartPage), Page(table.EndPage), string.Empty, caption);
                    if (!table.HasCaption) _preview.Rows[row].Cells[ColumnCaption].Style.ForeColor = SystemColors.GrayText;
                    _rowTables.Add(table);
                }
            }
            finally
            {
                _filling = false;
            }
            RefreshSheetNames();
        }

        /// <summary>Nom de la feuille de chaque tableau coché (les noms tiennent compte des tableaux décochés).</summary>
        private void RefreshSheetNames()
        {
            var plan = ExportPlan.Build(_tables, AllTables, _excluded);
            var names = plan.ToDictionary(p => p.Table.Index, p => p.SheetName);
            _filling = true;
            try
            {
                for (int i = 0; i < _rowTables.Count; i++)
                {
                    string name;
                    bool exported = names.TryGetValue(_rowTables[i].Index, out name);
                    var cell = _preview.Rows[i].Cells[ColumnSheet];
                    cell.Value = exported ? name : "(non exporté)";
                    cell.Style.ForeColor = exported ? SystemColors.WindowText : SystemColors.GrayText;
                }
            }
            finally
            {
                _filling = false;
            }
            _selectAll.Invalidate();
            _previewTitle.Text = string.Format(CultureInfo.CurrentCulture, "Feuilles Excel qui seront créées ({0})", plan.Count);
            _exportButton.Enabled = plan.Count > 0;
        }

        private static string Page(int page)
        {
            return page > 0 ? page.ToString(CultureInfo.CurrentCulture) : "\u2014";
        }

        private Label Section(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font(Font, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 14, 0, 2)
            };
        }

        private static Label Hint(string text, int width)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                MaximumSize = new Size(width - 26, 0),
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(26, 0, 0, 2)
            };
        }
    }
}
