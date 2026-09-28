using System;
using System.Drawing;
using System.Windows.Forms;
using WordTableToExcel.Word;

namespace WordTableToExcel.UI
{
    /// <summary>
    /// Fenêtre de progression modale : Word est désactivé pendant l'export (l'utilisateur ne peut
    /// pas modifier le document pendant sa lecture), l'interface reste réactive et l'export
    /// peut être annulé.
    /// </summary>
    internal sealed class ProgressDialog : Form, IExportProgress
    {
        private readonly Label _message;
        private readonly ProgressBar _bar;
        private readonly Button _cancelButton;
        private readonly Action<IExportProgress> _work;
        private Exception _error;
        private bool _finished;
        private bool _cancelRequested;
        private int _lastPump;

        private ProgressDialog(string title, Action<IExportProgress> work)
        {
            _work = work;

            Text = title;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ControlBox = true;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12),
                Location = new Point(0, 0)
            };

            _message = new Label
            {
                AutoSize = false,
                Width = 380,
                Height = 36,
                Text = "Préparation…",
                Margin = new Padding(0, 0, 0, 8)
            };
            _bar = new ProgressBar
            {
                Width = 380,
                Height = 18,
                Minimum = 0,
                Maximum = 1000,
                Style = ProgressBarStyle.Continuous,
                Margin = new Padding(0, 0, 0, 12)
            };
            _cancelButton = new Button
            {
                Text = "Annuler",
                AutoSize = true,
                Anchor = AnchorStyles.Right
            };
            _cancelButton.Click += (s, e) => RequestCancel();

            layout.Controls.Add(_message);
            layout.Controls.Add(_bar);
            layout.Controls.Add(_cancelButton);
            Controls.Add(layout);
            CancelButton = _cancelButton;
        }

        /// <summary>Exécute <paramref name="work"/> pendant l'affichage de la fenêtre ; relance son exception éventuelle.</summary>
        public static void Run(IWin32Window owner, string title, Action<IExportProgress> work)
        {
            Exception error;
            using (var dialog = new ProgressDialog(title, work))
            {
                dialog.ShowDialog(owner);
                error = dialog._error;
            }
            if (error != null) throw new ExportFailedException(error);
        }

        public bool IsCancellationRequested
        {
            get
            {
                Pump(false);
                return _cancelRequested;
            }
        }

        public void Report(string message, double fraction)
        {
            if (message != null && _message.Text != message) _message.Text = message;
            int value = (int)Math.Round(Math.Max(0, Math.Min(1, fraction)) * _bar.Maximum);
            if (_bar.Value != value) _bar.Value = value;
            Pump(false);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // Laisse la fenêtre se dessiner avant de commencer.
            BeginInvoke(new MethodInvoker(Execute));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_finished)
            {
                e.Cancel = true;
                RequestCancel();
            }
            base.OnFormClosing(e);
        }

        private void Execute()
        {
            try
            {
                Pump(true);
                _work(this);
            }
            catch (Exception ex)
            {
                _error = ex;
            }
            finally
            {
                _finished = true;
                Close();
            }
        }

        private void RequestCancel()
        {
            if (_cancelRequested) return;
            _cancelRequested = true;
            _cancelButton.Enabled = false;
            _message.Text = "Annulation en cours…";
        }

        /// <summary>Traite les messages Windows au plus toutes les 100 ms (affichage, bouton Annuler).</summary>
        private void Pump(bool force)
        {
            int now = Environment.TickCount;
            if (!force && unchecked(now - _lastPump) < 100) return;
            _lastPump = now;
            Application.DoEvents();
        }
    }

    /// <summary>Encapsule l'exception survenue pendant l'exécution d'une tâche avec progression.</summary>
    internal sealed class ExportFailedException : Exception
    {
        public ExportFailedException(Exception inner)
            : base(inner.Message, inner)
        {
        }
    }
}
