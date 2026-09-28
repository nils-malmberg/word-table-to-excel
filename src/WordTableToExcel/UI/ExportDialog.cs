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
    /// Boîte de dialogue principale : choix entre l'option A (tableaux avec légende) et
    /// l'option B (tous les tableaux), aperçu des feuilles qui seront créées, options.
    /// </summary>
    internal sealed class ExportDialog : Form
    {
        private readonly IList<TableEntry> _tables;
        private readonly RadioButton _captionedOnly;
        private readonly RadioButton _allTables;
        private readonly CheckBox _includeCaption;
        private readonly CheckBox _convertNumbers;
        private readonly ListView _preview;
        private readonly Button _exportButton;
        private readonly Label _previewTitle;

        public ExportDialog(string documentName, IList<TableEntry> tables, CaptionPosition convention, bool allTables, bool includeCaption, bool convertNumbers)
        {
            _tables = tables;
            int total = tables.Count;
            int captioned = tables.Count(t => t.HasCaption);

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

            const int width = 540;
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

            string summary = string.Format(CultureInfo.CurrentCulture, "{0} tableau{1} trouvé{2}, dont {3} avec une légende",
                total, total > 1 ? "x" : string.Empty, total > 1 ? "s" : string.Empty, captioned);
            if (captioned > 0)
            {
                summary += convention == CaptionPosition.Below ? " (légendes placées sous les tableaux)." : " (légendes placées au-dessus des tableaux).";
            }
            else
            {
                summary += ".";
            }
            root.Controls.Add(new Label { Text = summary, AutoSize = true, MaximumSize = new Size(width, 0), Margin = new Padding(0, 0, 0, 4) });

            // --- Choix A / B
            root.Controls.Add(Section("Tableaux à exporter"));
            _captionedOnly = new RadioButton
            {
                Text = string.Format(CultureInfo.CurrentCulture, "A \u2014 Uniquement les tableaux qui ont une légende ({0})", captioned),
                AutoSize = true,
                Enabled = captioned > 0,
                Margin = new Padding(8, 2, 0, 0)
            };
            root.Controls.Add(_captionedOnly);
            root.Controls.Add(Hint(captioned > 0
                ? "Légende : champ SEQ « Tableau », « Table », « Tabla »… ou texte commençant par « Tabl… », juste au-dessus ou au-dessous du tableau. Chaque feuille porte le nom de sa légende."
                : "Aucune légende de tableau n'a été détectée dans ce document.", width));

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
            _preview = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                MultiSelect = false,
                HideSelection = true,
                Size = new Size(width, 150),
                Margin = new Padding(0, 2, 0, 0)
            };
            _preview.Columns.Add("N°", 40, System.Windows.Forms.HorizontalAlignment.Right);
            _preview.Columns.Add("Feuille Excel", 215);
            _preview.Columns.Add("Légende", width - 40 - 215 - 24);
            root.Controls.Add(_preview);

            // --- Options
            root.Controls.Add(Section("Options"));
            _includeCaption = new CheckBox
            {
                Text = "Écrire la légende complète en haut de chaque feuille (cellule A1)",
                AutoSize = true,
                Checked = includeCaption,
                Margin = new Padding(8, 2, 0, 0)
            };
            _convertNumbers = new CheckBox
            {
                Text = "Convertir les nombres en valeurs numériques Excel",
                AutoSize = true,
                Checked = convertNumbers,
                Margin = new Padding(8, 4, 0, 0)
            };
            root.Controls.Add(_includeCaption);
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

            if (captioned == 0 || allTables) _allTables.Checked = true;
            else _captionedOnly.Checked = true;

            _captionedOnly.CheckedChanged += (s, e) => RefreshPreview();
            _allTables.CheckedChanged += (s, e) => RefreshPreview();
            RefreshPreview();
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

        private void RefreshPreview()
        {
            var plan = ExportPlan.Build(_tables, AllTables);
            _preview.BeginUpdate();
            _preview.Items.Clear();
            foreach (var sheet in plan)
            {
                var item = new ListViewItem(sheet.Table.Index.ToString(CultureInfo.CurrentCulture));
                item.SubItems.Add(sheet.SheetName);
                string caption = sheet.Table.HasCaption
                    ? (sheet.Table.CaptionPosition == CaptionPosition.Below ? "↓ " : "↑ ") + sheet.Table.Caption
                    : "(sans légende)";
                item.SubItems.Add(caption);
                if (!sheet.Table.HasCaption) item.ForeColor = SystemColors.GrayText;
                _preview.Items.Add(item);
            }
            _preview.EndUpdate();
            _previewTitle.Text = string.Format(CultureInfo.CurrentCulture, "Feuilles Excel qui seront créées ({0})", plan.Count);
            _exportButton.Enabled = plan.Count > 0;
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
