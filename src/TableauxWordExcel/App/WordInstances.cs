using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WordTableToExcel.Infrastructure;
using WordTableToExcel.Word;

namespace WordTableToExcel.App
{
    /// <summary>Instance de Word en cours d'exécution.</summary>
    internal sealed class WordInstance
    {
        /// <summary>Processus Word (0 si inconnu).</summary>
        public int ProcessId;
        /// <summary>Objet Word.Application.</summary>
        public object Application;
        /// <summary>Fenêtre de Word la plus récemment utilisée (IntPtr.Zero si inconnue).</summary>
        public IntPtr Window;
    }

    /// <summary>
    /// Accès aux instances de Word ouvertes par l'utilisateur, depuis l'application (autre processus).
    /// Les fenêtres Word sont parcourues de la plus récente à la plus ancienne : la première instance trouvée est
    /// celle dans laquelle l'utilisateur a cliqué en dernier. Chaque fenêtre de document donne accès à son
    /// application (même mécanisme que les outils d'accessibilité), ce qui fonctionne aussi quand plusieurs
    /// instances de Word sont ouvertes ; à défaut, l'instance déclarée par Word dans la table des objets actifs.
    /// Seules des lectures sont faites ici : rien n'est modifié dans Word.
    /// </summary>
    internal static class WordInstances
    {
        private const string WordFrameClass = "OpusApp";
        private const string WordDocumentClass = "_WwG";
        private const int WdNoProtection = -1;
        private const int WdMainTextStory = 1;
        private const int WdActiveEndPageNumber = 3;
        private const int WdWithInTable = 12;
        private const int WdCollapseEnd = 0;

        private struct TopWindow
        {
            public IntPtr Handle;
            public int ProcessId;
        }

        /// <summary>Instances de Word visibles, la plus récemment utilisée en premier.</summary>
        public static List<WordInstance> Find()
        {
            var result = new List<WordInstance>();
            var resolved = new HashSet<int>();
            foreach (var window in WordWindows())
            {
                if (resolved.Contains(window.ProcessId)) continue;
                object application = ApplicationFromWindow(window.Handle);
                if (application == null) continue;
                resolved.Add(window.ProcessId);
                result.Add(new WordInstance { ProcessId = window.ProcessId, Application = application, Window = window.Handle });
            }

            if (result.Count == 0)
            {
                object application = RunningWord();
                if (application != null && IsVisible(application)) result.Add(new WordInstance { Application = application });
            }
            return result;
        }

        /// <summary>Fenêtres principales visibles de Word, dans l'ordre d'empilement (la plus récente d'abord).</summary>
        private static List<TopWindow> WordWindows()
        {
            var windows = new List<TopWindow>();
            NativeMethods.EnumWindowsProc callback = (hwnd, l) =>
            {
                if (NativeMethods.IsWindowVisible(hwnd) && NativeMethods.ClassName(hwnd) == WordFrameClass)
                {
                    uint pid;
                    NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                    windows.Add(new TopWindow { Handle = hwnd, ProcessId = (int)pid });
                }
                return true;
            };
            NativeMethods.EnumWindows(callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            return windows;
        }

        private static object ApplicationFromWindow(IntPtr frame)
        {
            IntPtr documentWindow = IntPtr.Zero;
            NativeMethods.EnumWindowsProc callback = (hwnd, l) =>
            {
                if (NativeMethods.ClassName(hwnd) != WordDocumentClass) return true;
                documentWindow = hwnd;
                return false;
            };
            NativeMethods.EnumChildWindows(frame, callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            if (documentWindow == IntPtr.Zero) return null;

            try
            {
                Guid iid = NativeMethods.IidIDispatch;
                object window;
                int hr = NativeMethods.AccessibleObjectFromWindow(documentWindow, NativeMethods.ObjIdNativeOm, ref iid, out window);
                if (hr < 0 || window == null) return null;
                return ((dynamic)window).Application;
            }
            catch (Exception ex)
            {
                Log.Info("Fenêtre Word inaccessible : " + ex.Message);
                return null;
            }
        }

        private static object RunningWord()
        {
            try
            {
                return Marshal.GetActiveObject("Word.Application");
            }
            catch (COMException)
            {
                return null; // Word n'est pas lancé
            }
            catch (Exception ex)
            {
                Log.Info("Word introuvable : " + ex.Message);
                return null;
            }
        }

        private static bool IsVisible(object application)
        {
            try
            {
                return WordCom.IsTrue(((dynamic)application).Visible);
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ------------------------------------------------------------------ documents

        /// <summary>Documents ouverts dans une instance de Word (y compris en mode protégé), sans les documents invisibles.</summary>
        public static List<DocumentEntry> Documents(WordInstance instance)
        {
            var list = new List<DocumentEntry>();
            dynamic app = instance.Application;
            string active = null;
            try
            {
                if (WordCom.AsInt(app.Documents.Count) > 0) active = WordCom.AsString(app.ActiveDocument.FullName);
            }
            catch (Exception)
            {
                active = null; // fenêtre en mode protégé au premier plan
            }

            foreach (dynamic document in app.Documents)
            {
                try
                {
                    if (!HasVisibleWindow((object)document)) continue;
                    DocumentEntry entry = Describe((object)document, DocumentState.Open, instance.ProcessId);
                    entry.IsActive = active != null && string.Equals(entry.FullName, active, StringComparison.OrdinalIgnoreCase);
                    list.Add(entry);
                }
                catch (Exception ex)
                {
                    Log.Info("Document illisible dans la liste : " + ex.Message);
                }
            }

            try
            {
                foreach (dynamic window in app.ProtectedViewWindows) // Word 2010 et suivants
                {
                    try
                    {
                        DocumentEntry entry = Describe((object)window.Document, DocumentState.ProtectedView, instance.ProcessId);
                        entry.IsActive = WordCom.IsTrue(window.Active);
                        list.Add(entry);
                    }
                    catch (Exception ex)
                    {
                        Log.Info("Document en mode protégé illisible : " + ex.Message);
                    }
                }
            }
            catch (Exception)
            {
                // Word 2007 : pas de mode protégé.
            }
            return list;
        }

        private static bool HasVisibleWindow(object documentObject)
        {
            try
            {
                dynamic document = documentObject;
                dynamic windows = document.Windows;
                int count = WordCom.AsInt(windows.Count);
                for (int i = 1; i <= count; i++)
                {
                    if (WordCom.IsTrue(windows.Item(i).Visible)) return true;
                }
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        private static DocumentEntry Describe(object documentObject, DocumentState state, int processId)
        {
            dynamic document = documentObject;
            var entry = new DocumentEntry
            {
                State = state,
                ProcessId = processId,
                Name = WordCom.AsString(document.Name),
                FullName = WordCom.AsString(document.FullName)
            };
            try
            {
                entry.Folder = WordCom.AsString(document.Path);
            }
            catch (Exception)
            {
                entry.Folder = string.Empty;
            }
            try
            {
                entry.TableCount = WordCom.AsInt(document.Tables.Count);
            }
            catch (Exception)
            {
                entry.TableCount = -1;
            }
            if (string.IsNullOrEmpty(entry.FullName)) entry.FullName = entry.Name;
            return entry;
        }

        /// <summary>Retrouve un document de la liste dans Word (il a pu être fermé entre-temps : null).</summary>
        public static object FindDocument(DocumentEntry entry)
        {
            foreach (var instance in Find())
            {
                if (entry.ProcessId != 0 && instance.ProcessId != 0 && instance.ProcessId != entry.ProcessId) continue;
                dynamic app = instance.Application;
                if (entry.State == DocumentState.ProtectedView)
                {
                    try
                    {
                        foreach (dynamic window in app.ProtectedViewWindows)
                        {
                            object document = window.Document;
                            if (Matches(document, entry)) return document;
                        }
                    }
                    catch (Exception)
                    {
                        // Word 2007 : pas de mode protégé.
                    }
                    continue;
                }
                foreach (object document in app.Documents)
                {
                    if (Matches(document, entry)) return document;
                }
            }
            return null;
        }

        private static bool Matches(object documentObject, DocumentEntry entry)
        {
            try
            {
                dynamic document = documentObject;
                string fullName = WordCom.AsString(document.FullName);
                if (string.IsNullOrEmpty(fullName)) fullName = WordCom.AsString(document.Name);
                return string.Equals(fullName, entry.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ------------------------------------------------------------------ point d'insertion

        /// <summary>
        /// Document actif et position du curseur dans l'instance de Word la plus récente ;
        /// null si aucun document n'est ouvert.
        /// </summary>
        public static InsertionInfo ReadInsertionPoint(WordInstance instance)
        {
            dynamic app = instance.Application;
            try
            {
                dynamic protectedWindow = app.ActiveProtectedViewWindow;
                if (protectedWindow != null && WordCom.IsTrue(protectedWindow.Active))
                {
                    string name = WordCom.AsString(protectedWindow.Document.Name);
                    return new InsertionInfo
                    {
                        DocumentName = name,
                        Problem = "« " + name + " » est ouvert en mode protégé : dans Word, cliquez sur « Activer la modification », "
                                + "puis cliquez à l'endroit où insérer le tableau."
                    };
                }
            }
            catch (Exception)
            {
                // Word 2007, ou aucune fenêtre en mode protégé.
            }

            if (WordCom.AsInt(app.Documents.Count) == 0) return null;
            dynamic document = app.ActiveDocument;
            var info = new InsertionInfo { DocumentName = WordCom.AsString(document.Name) };

            try
            {
                if (WordCom.IsTrue(document.Final))
                {
                    info.Problem = "« " + info.DocumentName + " » est marqué comme final (lecture seule) : dans Word, cliquez sur "
                                 + "« Modifier quand même » dans la barre jaune, puis cliquez à l'endroit où insérer le tableau.";
                    return info;
                }
            }
            catch (Exception)
            {
                // Propriété absente des anciennes versions.
            }
            try
            {
                if (WordCom.AsInt(document.ProtectionType) != WdNoProtection)
                {
                    info.Problem = "« " + info.DocumentName + " » est protégé contre les modifications (Révision › Restreindre la modification) : "
                                 + "retirez la protection dans Word pour pouvoir y insérer des tableaux.";
                    return info;
                }
            }
            catch (Exception)
            {
                // Protection inconnue : l'import vérifiera à nouveau.
            }

            try
            {
                dynamic selection = app.Selection;
                info.HasSelection = WordCom.AsInt(selection.Start) != WordCom.AsInt(selection.End);
                info.MainStory = WordCom.AsInt(selection.StoryType) == WdMainTextStory;
                dynamic point = selection.Range.Duplicate;
                point.Collapse(WdCollapseEnd); // l'import se fait à la fin de la sélection
                info.InTable = WordCom.IsTrue(point.Information(WdWithInTable));
                if (!info.InTable)
                {
                    dynamic paragraph = point.Paragraphs.Item(1).Range;
                    info.ParagraphText = WordCom.AsString(paragraph.Text);
                    info.AtParagraphStart = WordCom.AsInt(point.Start) == WordCom.AsInt(paragraph.Start);
                }
                int page = WordCom.AsInt(point.Information(WdActiveEndPageNumber));
                info.Page = page > 0 && !WordCom.IsUndefined(page) ? page : 0;
            }
            catch (Exception ex)
            {
                Log.Info("Position du curseur inconnue : " + ex.Message);
            }
            return info;
        }

        /// <summary>Remet Word au premier plan (après un import, pour voir le résultat).</summary>
        public static void BringToFront(WordInstance instance)
        {
            IntPtr hwnd = IntPtr.Zero;
            try
            {
                hwnd = new IntPtr(Convert.ToInt64(((dynamic)instance.Application).ActiveWindow.Hwnd)); // Word 2013+
            }
            catch (Exception)
            {
                hwnd = instance.Window;
            }
            NativeMethods.BringToFront(hwnd);
        }
    }
}
