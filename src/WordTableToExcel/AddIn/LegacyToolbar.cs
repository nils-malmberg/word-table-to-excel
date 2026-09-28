using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using WordTableToExcel.AddIn.Interop;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.AddIn
{
    /// <summary>
    /// Bouton « Tableaux → Excel » dans la barre d'outils Standard de Word 2000, 2002 (XP) et 2003,
    /// qui n'ont pas de ruban. Le modèle Normal n'est pas marqué comme modifié.
    /// </summary>
    internal sealed class LegacyToolbar : IDisposable
    {
        private const string Tag = "WordTableToExcel.Export";

        private readonly object _application;
        private object _button;
        private IConnectionPoint _connectionPoint;
        private int _cookie;

        private LegacyToolbar(object application)
        {
            _application = application;
        }

        public static LegacyToolbar Create(object application, Action onClick)
        {
            var toolbar = new LegacyToolbar(application);
            try
            {
                toolbar.Install(onClick);
            }
            catch (Exception ex)
            {
                Log.Error("Création du bouton de barre d'outils", ex);
            }
            return toolbar;
        }

        private void Install(Action onClick)
        {
            dynamic app = _application;
            WithNormalTemplateUntouched(() =>
            {
                RemoveExistingButtons(app);
                dynamic bar = app.CommandBars.Item("Standard");
                dynamic button = bar.Controls.Add(Type: 1, Temporary: true); // msoControlButton
                button.Caption = "Tableaux → Excel";
                button.Style = 2; // msoButtonCaption
                button.Tag = Tag;
                button.TooltipText = "Exporter les tableaux du document vers Excel";
                button.BeginGroup = true;
                button.OnAction = "!<" + Connect.ProgIdValue + ">";
                _button = button;
            });

            var container = (IConnectionPointContainer)_button;
            Guid iid = typeof(CommandBarButtonEvents).GUID;
            container.FindConnectionPoint(ref iid, out _connectionPoint);
            _connectionPoint.Advise(new ButtonEventSink(onClick), out _cookie);
        }

        public void Dispose()
        {
            try
            {
                if (_connectionPoint != null && _cookie != 0) _connectionPoint.Unadvise(_cookie);
            }
            catch (Exception ex)
            {
                Log.Error("Déconnexion du bouton", ex);
            }
            _connectionPoint = null;

            try
            {
                dynamic app = _application;
                WithNormalTemplateUntouched(() => RemoveExistingButtons(app));
            }
            catch (Exception ex)
            {
                Log.Error("Suppression du bouton", ex);
            }
            _button = null;
        }

        private static void RemoveExistingButtons(dynamic app)
        {
            for (int guard = 0; guard < 20; guard++)
            {
                dynamic existing = app.CommandBars.FindControl(Tag: Tag);
                if (existing == null) return;
                existing.Delete();
            }
        }

        /// <summary>Word enregistre les barres d'outils dans Normal.dot : on évite qu'il demande à l'enregistrer.</summary>
        private void WithNormalTemplateUntouched(Action action)
        {
            dynamic app = _application;
            dynamic normal = null;
            object previousContext = null;
            bool wasSaved = true;
            try
            {
                normal = app.NormalTemplate;
                wasSaved = Convert.ToBoolean(normal.Saved);
                previousContext = app.CustomizationContext;
                app.CustomizationContext = normal;
            }
            catch (Exception)
            {
                normal = null;
            }

            try
            {
                action();
            }
            finally
            {
                try
                {
                    if (previousContext != null) app.CustomizationContext = previousContext;
                    if (normal != null && wasSaved) normal.Saved = true;
                }
                catch (Exception)
                {
                    // Sans conséquence.
                }
            }
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        public sealed class ButtonEventSink : CommandBarButtonEvents
        {
            private readonly Action _onClick;

            public ButtonEventSink(Action onClick)
            {
                _onClick = onClick;
            }

            public void Click(object Ctrl, ref bool CancelDefault)
            {
                try
                {
                    _onClick();
                }
                catch (Exception ex)
                {
                    Log.Error("Clic sur le bouton", ex);
                }
            }
        }
    }
}
