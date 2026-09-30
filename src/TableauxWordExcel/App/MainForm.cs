using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using WordTableToExcel.Export;
using WordTableToExcel.Import;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.UI;
using WordTableToExcel.Word;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Fenêtre de l'application : en haut, l'export des tableaux d'un document Word (ouvert dans Word ou simple
    /// fichier) vers Excel ; en bas, l'import de tableaux Excel à l'endroit où se trouve le curseur dans Word.
    /// La fenêtre s'actualise d'elle-même quand l'utilisateur y revient après avoir cliqué dans Word.
    /// </summary>
    internal sealed class MainForm : Form
    {
        public const string AppTitle = "Tableaux Word ↔ Excel";

        private static readonly Color InfoBack = Color.FromArgb(0xEA, 0xF2, 0xFB);
        private static readonly Color InfoBorder = Color.FromArgb(0xB7, 0xCF, 0xEA);
        private static readonly Color WarningBack = Color.FromArgb(0xFF, 0xF4, 0xE0);
        private static readonly Color WarningBorder = Color.FromArgb(0xF0, 0xC8, 0x80);

        private readonly ListView _documents;
        private readonly ListViewGroup _openGroup;
        private readonly ListViewGroup _fileGroup;
        private readonly Label _emptyList;
        private readonly Button _addFile;
        private readonly Button _refresh;
        private readonly Button _export;
        private readonly Panel _insertionFrame;
        private readonly Label _insertion;
        private readonly Button _import;
        private readonly Label _status;
        private readonly CheckBox _topMost;
        private readonly ContextMenuStrip _fileMenu;
        private readonly Font _boldListFont;
        /// <summary>Word invisible des exports de fichiers non ouverts, gardé quelques minutes entre deux exports.</summary>
        private readonly HiddenWord _hiddenWord = new HiddenWord();
        private readonly Timer _idleTimer;

        private readonly List<string> _files = new List<string>();
        private readonly string[] _arguments;
        private List<DocumentEntry> _entries = new List<DocumentEntry>();
        private bool _busy;
        private bool _wordRunning;
        private int _lastRefresh;

        public MainForm(string[] arguments)
        {
            _arguments = arguments ?? new string[0];

            Text = AppTitle;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.Manual;
            MinimumSize = new Size(520, 500);
            Size = new Size(580, 600);
            KeyPreview = true;
            AllowDrop = true;
            Icon = LoadIcon();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 10,
                Padding = new Padding(12, 10, 12, 8)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // --- Export : Word → Excel
            root.Controls.Add(Heading("Word → Excel : exporter les tableaux d'un document"));
            root.Controls.Add(Hint("Choisissez un document ouvert dans Word, ou ajoutez un fichier Word. "
                + "Le document est seulement lu : il n'est jamais modifié."));

            _documents = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                ShowItemToolTips = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _documents.Columns.Add("Document", 210);
            _documents.Columns.Add("Tableaux", 70, System.Windows.Forms.HorizontalAlignment.Right);
            _documents.Columns.Add("Emplacement", 220);
            _openGroup = new ListViewGroup("open", "Ouverts dans Word");
            _fileGroup = new ListViewGroup("files", "Fichiers ajoutés (ouverts en arrière-plan, le temps de l'export)");
            _documents.Groups.Add(_openGroup);
            _documents.Groups.Add(_fileGroup);
            _documents.SelectedIndexChanged += (s, e) => UpdateButtons();
            _documents.ItemActivate += (s, e) => ExportSelected();
            _documents.KeyDown += OnListKeyDown;
            _documents.Resize += (s, e) => FitLastColumn();
            _documents.MouseUp += OnListMouseUp;
            _boldListFont = new Font(_documents.Font, FontStyle.Bold);

            _emptyList = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.GrayText,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand,
                UseMnemonic = false,
                Visible = false
            };
            _emptyList.Click += (s, e) => AddFilesFromDialog();

            var listPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4) };
            listPanel.Controls.Add(_documents);
            listPanel.Controls.Add(_emptyList);
            _emptyList.BringToFront();
            root.Controls.Add(listPanel);

            _addFile = new Button { Text = "Ajouter un fichier Word…", AutoSize = true, MinimumSize = new Size(0, 28), Margin = new Padding(0, 0, 6, 0) };
            _addFile.Click += (s, e) => AddFilesFromDialog();
            _refresh = new Button { Text = "Actualiser", AutoSize = true, MinimumSize = new Size(0, 28), Margin = new Padding(0) };
            _refresh.Click += (s, e) => RefreshFromWord(true);
            _export = MainButton("Exporter vers Excel…");
            _export.Click += (s, e) => ExportSelected();
            var exportButtons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, AutoSize = true, Margin = new Padding(0) };
            exportButtons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            exportButtons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            exportButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            exportButtons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            exportButtons.Controls.Add(_addFile, 0, 0);
            exportButtons.Controls.Add(_refresh, 1, 0);
            exportButtons.Controls.Add(_export, 3, 0);
            root.Controls.Add(exportButtons);

            root.Controls.Add(new Label { AutoSize = false, Height = 2, Dock = DockStyle.Fill, BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, 12, 0, 8) });

            // --- Import : Excel → Word
            root.Controls.Add(Heading("Excel → Word : importer des tableaux dans un document"));
            root.Controls.Add(Hint("1. Dans Word, cliquez à l'endroit où le tableau doit être inséré.\n"
                + "2. Revenez ici et cliquez sur « Importer depuis Excel… » (ou déposez un classeur Excel sur cette fenêtre)."));

            _insertion = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(6, 4, 6, 4),
                Margin = new Padding(1),
                UseMnemonic = false
            };
            _insertionFrame = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), Margin = new Padding(0, 6, 0, 6) };
            _insertionFrame.Controls.Add(_insertion);
            root.Controls.Add(_insertionFrame);

            _import = MainButton("Importer depuis Excel…");
            _import.Anchor = AnchorStyles.Right;
            _import.Click += (s, e) => ImportFrom(null);
            root.Controls.Add(_import);

            // --- Barre du bas
            _status = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                UseMnemonic = false,
                Margin = new Padding(0, 0, 8, 0)
            };
            var help = new LinkLabel { Text = "Aide", AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(0, 0, 12, 0) };
            help.LinkClicked += (s, e) => ShowHelp();
            _topMost = new CheckBox { Text = "Toujours visible", AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(0) };
            _topMost.CheckedChanged += (s, e) => TopMost = _topMost.Checked;
            var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.Controls.Add(_status, 0, 0);
            bottom.Controls.Add(help, 1, 0);
            bottom.Controls.Add(_topMost, 2, 0);
            root.Controls.Add(bottom);

            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // titre export
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // explication
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));          // liste
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // boutons export
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // séparateur
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // titre import
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // étapes
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, Font.Height * 4 + 34)); // point d'insertion (4 lignes)
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // bouton import
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));              // barre du bas
            Controls.Add(root);

            _fileMenu = new ContextMenuStrip();
            _fileMenu.Items.Add("Exporter vers Excel…", null, (s, e) => ExportSelected());
            _fileMenu.Items.Add("Retirer de la liste", null, (s, e) => RemoveSelectedFile());

            EnableDrop(this);
            RestoreWindow();
            ShowInsertion(null, false, false);
            UpdateButtons();

            _idleTimer = new Timer { Interval = 5000 };
            _idleTimer.Tick += (s, e) => CheckHiddenWord();
            _idleTimer.Start();
        }

        // ------------------------------------------------------------------ construction

        private Label Heading(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font(Font.FontFamily, Font.SizeInPoints + 1.5f, FontStyle.Bold),
                ForeColor = SystemInformation.HighContrast ? SystemColors.ControlText : Color.FromArgb(0x1F, 0x3A, 0x5F),
                Margin = new Padding(0, 0, 0, 2),
                UseMnemonic = false
            };
        }

        private static Label Hint(string text)
        {
            return new Label { Text = text, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 2), UseMnemonic = false };
        }

        private Button MainButton(string text)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                MinimumSize = new Size(190, 32),
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0)
            };
        }

        private static Icon LoadIcon()
        {
            try
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("TableauxWordExcel.ico"))
                {
                    if (stream != null) return new Icon(stream);
                }
            }
            catch (Exception)
            {
                // Icône par défaut.
            }
            return null;
        }

        private void EnableDrop(Control control)
        {
            control.AllowDrop = true;
            control.DragEnter += OnDragEnter;
            control.DragDrop += OnDragDrop;
            foreach (Control child in control.Controls) EnableDrop(child);
        }

        // ------------------------------------------------------------------ fenêtre

        private void RestoreWindow()
        {
            var settings = Settings.Load();
            _topMost.Checked = settings.AppTopMost;
            TopMost = settings.AppTopMost;

            Rectangle bounds;
            if (TryParseBounds(settings.AppBounds, out bounds) && Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(bounds)))
            {
                Bounds = bounds;
                return;
            }
            // Première ouverture : en bas à droite de l'écran, pour ne pas masquer le document Word.
            Rectangle area = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(Math.Max(area.Left, area.Right - Width - 24), Math.Max(area.Top, area.Bottom - Height - 24));
        }

        private static bool TryParseBounds(string text, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            if (string.IsNullOrEmpty(text)) return false;
            var parts = text.Split(',');
            if (parts.Length != 4) return false;
            var values = new int[4];
            for (int i = 0; i < 4; i++)
            {
                if (!int.TryParse(parts[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out values[i])) return false;
            }
            if (values[2] < 200 || values[3] < 200) return false;
            bounds = new Rectangle(values[0], values[1], values[2], values[3]);
            return true;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            FitLastColumn();
            RefreshFromWord(true);
            if (_arguments.Length > 0) BeginInvoke(new MethodInvoker(() => OpenFiles(_arguments)));
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            // Retour dans la fenêtre après un clic dans Word : documents et curseur à jour.
            if (!_busy) BeginInvoke(new MethodInvoker(() => RefreshFromWord(false)));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_busy)
            {
                e.Cancel = true;
                return;
            }
            try
            {
                var settings = Settings.Load();
                Rectangle bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                settings.AppBounds = string.Join(",", new[] { bounds.X, bounds.Y, bounds.Width, bounds.Height }.Select(v => v.ToString(CultureInfo.InvariantCulture)).ToArray());
                settings.AppTopMost = _topMost.Checked;
                settings.Save();
            }
            catch (Exception ex)
            {
                Log.Error("Enregistrement de la position de la fenêtre", ex);
            }
            _idleTimer.Stop();
            try
            {
                using (OleMessageFilter.Timeout(OleMessageFilter.RefreshTimeoutMs))
                {
                    _hiddenWord.Dispose(); // Word invisible refermé (ou rendu à l'utilisateur s'il contient un de ses documents)
                }
            }
            catch (Exception ex)
            {
                Log.Error("Fermeture du Word invisible", ex);
            }
            base.OnFormClosing(e);
        }

        /// <summary>Toutes les 5 secondes, hors opération : Word invisible refermé après inactivité, ou rendu à l'utilisateur.</summary>
        private void CheckHiddenWord()
        {
            if (_busy || !_hiddenWord.IsRunning) return;
            try
            {
                using (OleMessageFilter.Timeout(OleMessageFilter.RefreshTimeoutMs))
                {
                    _hiddenWord.CheckIdle();
                }
            }
            catch (Exception ex)
            {
                Log.Error("Surveillance du Word invisible", ex);
            }
            if (!_hiddenWord.IsRunning) RefreshFromWord(true); // Word rendu visible : ses documents apparaissent dans la liste
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.F5)
            {
                RefreshFromWord(true);
                e.Handled = true;
            }
        }

        // ------------------------------------------------------------------ actualisation

        /// <summary>Relit la liste des documents ouverts et la position du curseur dans Word.</summary>
        private void RefreshFromWord(bool force)
        {
            if (_busy || IsDisposed) return;
            int now = Environment.TickCount;
            if (!force && unchecked(now - _lastRefresh) < 500) return;
            _lastRefresh = now;

            var open = new List<DocumentEntry>();
            InsertionInfo insertion = null;
            bool running = false, busy = false;
            string version = null;
            _busy = true;
            try
            {
                using (OleMessageFilter.Timeout(OleMessageFilter.RefreshTimeoutMs))
                {
                    var instances = WordInstances.Find();
                    running = instances.Count > 0;
                    foreach (var instance in instances) open.AddRange(WordInstances.Documents(instance));
                    if (running)
                    {
                        version = WordCom.WordVersion(instances[0].Application);
                        insertion = WordInstances.ReadInsertionPoint(instances[0]);
                    }
                }
            }
            catch (Exception ex)
            {
                busy = ComErrors.IsBusy(ComErrors.HResultOf(ex));
                Log.Info("Actualisation depuis Word : " + ex.Message);
            }
            finally
            {
                _busy = false;
                ReleaseWordObjects();
            }

            if (busy)
            {
                // Word occupé (boîte de dialogue ouverte) : on garde la liste précédente.
                ShowInsertion(null, true, true);
                _status.Text = "Word est occupé (une boîte de dialogue est peut-être ouverte dans Word).";
                return;
            }

            _wordRunning = running;
            FillList(DocumentList.Merge(open, _files));
            ShowInsertion(insertion, running, false);
            int count = open.Count;
            _status.Text = !running ? "Word n'est pas ouvert."
                : "Word " + version + " · " + count.ToString(CultureInfo.CurrentCulture) + (count > 1 ? " documents ouverts" : " document ouvert");
        }

        private static void ReleaseWordObjects()
        {
            // Libère les références aux objets de Word : Word peut ensuite être fermé normalement par l'utilisateur.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void FillList(List<DocumentEntry> entries)
        {
            string selectedKey = SelectedEntry() != null ? SelectedEntry().Key : null;
            _entries = entries;

            _documents.BeginUpdate();
            _documents.Items.Clear();
            ListViewItem toSelect = null, active = null;
            foreach (var entry in entries)
            {
                string name = entry.Name + (entry.State == DocumentState.ProtectedView ? " (mode protégé)" : string.Empty);
                var item = new ListViewItem(name)
                {
                    Tag = entry,
                    Group = entry.State == DocumentState.File ? _fileGroup : _openGroup,
                    ToolTipText = entry.FullName
                };
                item.SubItems.Add(entry.TableCount >= 0 ? entry.TableCount.ToString(CultureInfo.CurrentCulture) : "—");
                item.SubItems.Add(entry.IsSaved ? entry.Folder : "(jamais enregistré)");
                if (entry.IsActive) item.Font = _boldListFont;
                if (entry.TableCount == 0) item.ForeColor = SystemColors.GrayText;
                _documents.Items.Add(item);
                if (selectedKey != null && entry.Key == selectedKey) toSelect = item;
                if (entry.IsActive && active == null) active = item;
            }
            _documents.EndUpdate();

            if (toSelect == null) toSelect = active ?? (_documents.Items.Count > 0 ? _documents.Items[0] : null);
            if (toSelect != null)
            {
                toSelect.Selected = true;
                toSelect.Focused = true;
                toSelect.EnsureVisible();
            }

            _emptyList.Text = (_wordRunning ? "Aucun document n'est ouvert dans Word." : "Word n'est pas ouvert.")
                + "\n\nOuvrez un document dans Word, ou cliquez ici pour choisir un fichier Word\n(vous pouvez aussi le déposer sur cette fenêtre).";
            // L'un ou l'autre, jamais superposés.
            _emptyList.Visible = entries.Count == 0;
            _documents.Visible = entries.Count > 0;
            UpdateButtons();
        }

        /// <summary>Contenu d'exemple, sans Word (autotest : capture de la fenêtre).</summary>
        internal void ShowPreview(List<DocumentEntry> entries, InsertionInfo insertion, string status)
        {
            _wordRunning = true;
            FillList(entries);
            ShowInsertion(insertion, true, false);
            _status.Text = status;
            FitLastColumn();
        }

        private void FitLastColumn()
        {
            if (_documents.Columns.Count < 3) return;
            int width = _documents.ClientSize.Width - _documents.Columns[0].Width - _documents.Columns[1].Width - 4;
            _documents.Columns[2].Width = Math.Max(120, width);
        }

        private void ShowInsertion(InsertionInfo info, bool running, bool wordBusy)
        {
            bool warning;
            if (wordBusy)
            {
                _insertion.Text = "Word est occupé : une boîte de dialogue est peut-être ouverte dans Word. Fermez-la, puis revenez dans cette fenêtre.";
                warning = true;
            }
            else
            {
                _insertion.Text = InsertionInfo.Describe(info, running);
                warning = info == null || !info.CanImport;
            }
            if (SystemInformation.HighContrast)
            {
                _insertionFrame.BackColor = SystemColors.WindowFrame;
                _insertion.BackColor = SystemColors.Window;
                _insertion.ForeColor = SystemColors.WindowText;
                return;
            }
            _insertionFrame.BackColor = warning ? WarningBorder : InfoBorder;
            _insertion.BackColor = warning ? WarningBack : InfoBack;
            _insertion.ForeColor = SystemColors.ControlText;
        }

        private DocumentEntry SelectedEntry()
        {
            return _documents.SelectedItems.Count > 0 ? _documents.SelectedItems[0].Tag as DocumentEntry : null;
        }

        private void UpdateButtons()
        {
            _export.Enabled = SelectedEntry() != null;
        }

        // ------------------------------------------------------------------ fichiers

        private void AddFilesFromDialog()
        {
            var settings = Settings.Load();
            string folder = !string.IsNullOrEmpty(settings.AppLastWordFolder) && Directory.Exists(settings.AppLastWordFolder)
                ? settings.AppLastWordFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            using (var dialog = new OpenFileDialog
            {
                Title = "Ajouter des documents Word",
                Filter = DocumentList.WordFileFilter,
                Multiselect = true,
                CheckFileExists = true,
                InitialDirectory = folder
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK || dialog.FileNames.Length == 0) return;
                settings.AppLastWordFolder = Path.GetDirectoryName(dialog.FileNames[0]);
                settings.Save();
                AddFiles(dialog.FileNames);
            }
        }

        private void AddFiles(IEnumerable<string> paths)
        {
            string first = null;
            foreach (var path in paths)
            {
                string full;
                try
                {
                    full = Path.GetFullPath(path);
                }
                catch (Exception)
                {
                    continue;
                }
                if (!File.Exists(full)) continue;
                if (first == null) first = full;
                if (!_files.Any(f => DocumentList.SamePath(f, full))) _files.Add(full);
            }
            if (first == null) return;
            RefreshFromWord(true);
            foreach (ListViewItem item in _documents.Items)
            {
                var entry = (DocumentEntry)item.Tag;
                if (!DocumentList.SamePath(entry.FullName, first)) continue;
                item.Selected = true;
                item.Focused = true;
                item.EnsureVisible();
                break;
            }
            _documents.Focus();
        }

        private void RemoveSelectedFile()
        {
            var entry = SelectedEntry();
            if (entry == null || entry.State != DocumentState.File) return;
            _files.RemoveAll(f => DocumentList.SamePath(f, entry.FullName));
            FillList(_entries.Where(e => e != entry).ToList());
        }

        private void OnListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                RemoveSelectedFile();
                e.Handled = true;
            }
        }

        private void OnListMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var item = _documents.GetItemAt(e.X, e.Y);
            if (item == null) return;
            item.Selected = true;
            var entry = (DocumentEntry)item.Tag;
            _fileMenu.Items[1].Visible = entry.State == DocumentState.File;
            _fileMenu.Show(_documents, e.Location);
        }

        private static List<string> DroppedFiles(IDataObject data)
        {
            try
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                if (files != null) return files.ToList();
            }
            catch (Exception)
            {
                // Contenu déposé illisible.
            }
            return new List<string>();
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            bool accepted = !_busy && DroppedFiles(e.Data).Any(f => DocumentList.IsWordFile(f) || DocumentList.IsExcelFile(f));
            e.Effect = accepted ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            var files = DroppedFiles(e.Data);
            // Traitement différé : l'Explorateur n'attend pas la fin de l'export ou de l'import.
            BeginInvoke(new MethodInvoker(() => OpenFiles(files)));
        }

        /// <summary>Fichiers déposés sur la fenêtre ou passés au lancement : Word → liste d'export, Excel → import.</summary>
        private void OpenFiles(IList<string> files)
        {
            if (_busy) return;
            var word = files.Where(DocumentList.IsWordFile).ToList();
            string excel = files.FirstOrDefault(DocumentList.IsExcelFile);
            if (word.Count > 0) AddFiles(word);
            if (excel != null) ImportFrom(excel);
            if (word.Count == 0 && excel == null && files.Count > 0)
            {
                Messages.Info(this, "Déposez un document Word (.docx, .doc…) pour l'exporter, ou un classeur Excel (.xlsx, .xls…) pour l'importer.");
            }
        }

        // ------------------------------------------------------------------ export

        private void ExportSelected()
        {
            var entry = SelectedEntry();
            if (entry == null || _busy) return;
            RunAction("L'export des tableaux a échoué.", true, () =>
            {
                if (entry.State == DocumentState.File)
                {
                    ExportFile(entry);
                    return;
                }
                object document = WordInstances.FindDocument(entry);
                if (document == null)
                {
                    Messages.Info(this, "« " + entry.Name + " » n'est plus ouvert dans Word.\n\nLa liste va être actualisée.");
                    return;
                }
                object application = ((dynamic)document).Application;
                new ExportService(application) { ReadContentFromXml = true }.Run(this, document);
            });
        }

        private void ExportFile(DocumentEntry entry)
        {
            if (!File.Exists(entry.FullName))
            {
                Messages.Warning(this, "Le fichier est introuvable :\n" + entry.FullName + "\n\nIl a peut-être été déplacé ou supprimé.");
                _files.RemoveAll(f => DocumentList.SamePath(f, entry.FullName));
                return;
            }

            object document = null;
            try
            {
                ProgressDialog.Run(this, "Ouverture du document", progress =>
                {
                    progress.Report(_hiddenWord.IsRunning
                        ? "Ouverture de « " + entry.Name + " » en arrière-plan (lecture seule)…"
                        : "Démarrage de Word en arrière-plan, puis ouverture de « " + entry.Name + " » (lecture seule)…", 0.3);
                    document = _hiddenWord.Open(entry.FullName);
                    if (progress.IsCancellationRequested) throw new OperationCanceledException();
                    progress.Report("Document ouvert.", 1);
                });
            }
            catch (ExportFailedException ex)
            {
                _hiddenWord.CloseDocument();
                if (ex.InnerException is OperationCanceledException) return;
                var inner = ex.InnerException ?? ex;
                Log.Error("Ouverture en arrière-plan de " + entry.FullName, inner);
                string friendly = ComErrors.FriendlyMessage(inner);
                Messages.Warning(this, "Word n'a pas pu ouvrir ce fichier :\n" + entry.FullName + "\n\n"
                    + (friendly ?? "Il est peut-être protégé par un mot de passe, endommagé, ou dans un format que Word ne sait pas lire.\n\nDétail : " + inner.Message)
                    + "\n\nVous pouvez aussi l'ouvrir vous-même dans Word : il apparaîtra alors dans la liste des documents ouverts.");
                return;
            }

            try
            {
                new ExportService(_hiddenWord.Application) { ReadContentFromXml = true }.Run(this, document);
            }
            finally
            {
                // Document refermé ; Word reste prêt quelques minutes pour l'export suivant.
                _hiddenWord.CloseDocument();
            }
        }

        // ------------------------------------------------------------------ import

        private void ImportFrom(string excelPath)
        {
            if (_busy) return;
            RunAction("L'import du classeur Excel a échoué.\nSi des tableaux ont déjà été insérés, Ctrl+Z (Annuler) dans Word les retire.", false, () =>
            {
                var instances = WordInstances.Find();
                if (instances.Count == 0)
                {
                    Messages.Info(this, "Word n'est pas ouvert.\n\nOuvrez votre document dans Word, cliquez à l'endroit où le tableau doit être inséré, "
                        + "puis revenez ici et cliquez sur « Importer depuis Excel… ».");
                    return;
                }
                var instance = instances[0];
                bool inserted = new ImportService(instance.Application, this).Run(excelPath);
                if (inserted) WordInstances.BringToFront(instance);
            });
        }

        // ------------------------------------------------------------------ exécution

        /// <summary>Exécute une action sur Word : une seule à la fois, erreurs de communication expliquées clairement.</summary>
        private void RunAction(string failure, bool documentUnchanged, Action action)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                using (OleMessageFilter.Timeout(OleMessageFilter.OperationTimeoutMs))
                {
                    action();
                }
            }
            catch (Exception ex)
            {
                var inner = ex is ExportFailedException && ex.InnerException != null ? ex.InnerException : ex;
                string friendly = ComErrors.FriendlyMessage(inner);
                if (friendly != null)
                {
                    Log.Error(failure, inner);
                    Messages.Warning(this, failure + "\n\n" + friendly);
                }
                else
                {
                    Messages.Error(this, failure, inner, documentUnchanged);
                }
            }
            finally
            {
                _busy = false;
                ReleaseWordObjects();
            }
            RefreshFromWord(true);
        }

        private void ShowHelp()
        {
            Messages.Info(this,
                "Exporter vers Excel\n"
                + "Sélectionnez un document ouvert dans Word, ou ajoutez un fichier Word (il est ouvert en arrière-plan, en lecture seule), "
                + "puis cliquez sur « Exporter vers Excel… ». Choisissez les tableaux, puis l'emplacement du classeur : "
                + "une feuille par tableau, avec la mise en forme. Le document Word n'est jamais modifié.\n\n"
                + "Importer depuis Excel\n"
                + "Dans Word, cliquez à l'endroit voulu, revenez dans cette fenêtre et cliquez sur « Importer depuis Excel… » "
                + "(ou déposez un classeur Excel sur la fenêtre). Chaque feuille choisie devient un tableau Word, avec sa mise en forme "
                + "et les valeurs telles qu'Excel les affiche, et une légende numérotée si vous le souhaitez. "
                + "Dans Word, Ctrl+Z annule l'import.\n\n"
                + "« Toujours visible » garde cette fenêtre au-dessus de Word. F5 actualise la liste.\n\n"
                + "Version " + Version + "\nJournal de diagnostic : " + Log.FilePath);
        }

        public static string Version
        {
            get
            {
                var v = Assembly.GetExecutingAssembly().GetName().Version;
                return v.Major + "." + v.Minor + "." + v.Build;
            }
        }
    }
}
