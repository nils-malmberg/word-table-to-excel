using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using WordTableToExcel.AddIn.Interop;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.AddIn
{
    /// <summary>
    /// Boutons « Tableaux → Excel » et « Excel → Tableaux » dans la barre d'outils Standard de Word 2000,
    /// 2002 (XP) et 2003, qui n'ont pas de ruban. Le modèle Normal n'est pas marqué comme modifié.
    /// </summary>
    internal sealed class LegacyToolbar : IDisposable
    {
        private const string ExportTag = "WordTableToExcel.Export";
        private const string ImportTag = "WordTableToExcel.Import";
        private static readonly string[] Tags = { ExportTag, ImportTag };

        private sealed class ButtonHandle
        {
            public object Button;
            public IConnectionPoint ConnectionPoint;
            public int Cookie;
        }

        private readonly object _application;
        private readonly List<ButtonHandle> _buttons = new List<ButtonHandle>();

        private LegacyToolbar(object application)
        {
            _application = application;
        }

        public static LegacyToolbar Create(object application, Action onExport, Action onImport)
        {
            var toolbar = new LegacyToolbar(application);
            try
            {
                toolbar.Install(onExport, onImport);
            }
            catch (Exception ex)
            {
                Log.Error("Création des boutons de barre d'outils", ex);
            }
            return toolbar;
        }

        private void Install(Action onExport, Action onImport)
        {
            dynamic app = _application;
            object export = null, import = null;
            WithNormalTemplateUntouched(() =>
            {
                RemoveExistingButtons(app);
                dynamic bar = app.CommandBars.Item("Standard");
                export = AddButton(bar, "Tableaux → Excel", "Exporter les tableaux du document vers Excel", ExportTag, true);
                import = AddButton(bar, "Excel → Tableaux", "Importer des tableaux depuis un classeur Excel", ImportTag, false);
            });
            Connect(export, onExport);
            Connect(import, onImport);
        }

        private static object AddButton(dynamic bar, string caption, string tooltip, string tag, bool beginGroup)
        {
            dynamic button = bar.Controls.Add(Type: 1, Temporary: true); // msoControlButton
            button.Caption = caption;
            button.Style = 2; // msoButtonCaption
            button.Tag = tag; // les événements Click sont distribués par étiquette
            button.TooltipText = tooltip;
            button.BeginGroup = beginGroup;
            button.OnAction = "!<" + AddIn.Connect.ProgIdValue + ">";
            return button;
        }

        private void Connect(object button, Action onClick)
        {
            if (button == null) return;
            var handle = new ButtonHandle { Button = button };
            var container = (IConnectionPointContainer)button;
            Guid iid = typeof(CommandBarButtonEvents).GUID;
            container.FindConnectionPoint(ref iid, out handle.ConnectionPoint);
            handle.ConnectionPoint.Advise(new ButtonEventSink(onClick), out handle.Cookie);
            _buttons.Add(handle);
        }

        public void Dispose()
        {
            foreach (var handle in _buttons)
            {
                try
                {
                    if (handle.ConnectionPoint != null && handle.Cookie != 0) handle.ConnectionPoint.Unadvise(handle.Cookie);
                }
                catch (Exception ex)
                {
                    Log.Error("Déconnexion d'un bouton", ex);
                }
            }
            _buttons.Clear();

            try
            {
                dynamic app = _application;
                WithNormalTemplateUntouched(() => RemoveExistingButtons(app));
            }
            catch (Exception ex)
            {
                Log.Error("Suppression des boutons", ex);
            }
        }

        private static void RemoveExistingButtons(dynamic app)
        {
            foreach (var tag in Tags)
            {
                for (int guard = 0; guard < 20; guard++)
                {
                    dynamic existing = app.CommandBars.FindControl(Tag: tag);
                    if (existing == null) break;
                    existing.Delete();
                }
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
