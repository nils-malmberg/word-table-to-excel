using System;
using System.Runtime.InteropServices;

namespace WordTableToExcel.App
{
    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    /// <summary>
    /// Filtre de messages OLE du thread de l'interface. Word refuse les appels venant d'un autre programme
    /// quand il est occupé (boîte de dialogue ouverte, saisie en cours) : sans filtre, l'appel échoue aussitôt
    /// (« L'appel a été rejeté par l'appelé ») ; avec ce filtre, il est réessayé quelques secondes, puis abandonné
    /// proprement avec un message clair.
    /// </summary>
    internal sealed class OleMessageFilter : IOleMessageFilter, IDisposable
    {
        private const int ServerCallIsHandled = 0;
        private const int PendingMsgWaitDefProcess = 2;

        /// <summary>Attente maximale (ms) d'un Word occupé pendant une opération demandée par l'utilisateur.</summary>
        public const int OperationTimeoutMs = 10000;
        /// <summary>Attente maximale (ms) pendant l'actualisation automatique de la fenêtre.</summary>
        public const int RefreshTimeoutMs = 1500;

        [ThreadStatic]
        private static int _timeoutMs;

        private IOleMessageFilter _previous;
        private bool _registered;

        private OleMessageFilter()
        {
        }

        /// <summary>Installe le filtre sur le thread courant (thread STA de l'interface).</summary>
        public static OleMessageFilter Register()
        {
            var filter = new OleMessageFilter();
            IOleMessageFilter previous;
            int hr = CoRegisterMessageFilter(filter, out previous);
            filter._registered = hr == 0;
            filter._previous = previous;
            return filter;
        }

        public bool IsRegistered
        {
            get { return _registered; }
        }

        /// <summary>Change l'attente maximale le temps d'un bloc <c>using</c>.</summary>
        public static IDisposable Timeout(int milliseconds)
        {
            return new TimeoutScope(milliseconds);
        }

        private static int CurrentTimeout
        {
            get { return _timeoutMs > 0 ? _timeoutMs : OperationTimeoutMs; }
        }

        public void Dispose()
        {
            if (!_registered) return;
            _registered = false;
            IOleMessageFilter ignored;
            CoRegisterMessageFilter(_previous, out ignored);
            _previous = null;
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return ServerCallIsHandled;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            return ComErrors.RetryDelay(dwRejectType, dwTickCount, CurrentTimeout);
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            // Laisse l'interface se redessiner pendant que Word travaille.
            return PendingMsgWaitDefProcess;
        }

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        private sealed class TimeoutScope : IDisposable
        {
            private readonly int _previous;
            private bool _disposed;

            public TimeoutScope(int milliseconds)
            {
                _previous = _timeoutMs;
                _timeoutMs = milliseconds;
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _timeoutMs = _previous;
            }
        }
    }
}
