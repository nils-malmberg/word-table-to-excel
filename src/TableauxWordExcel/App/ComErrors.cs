using System;
using System.Runtime.InteropServices;

namespace WordTableToExcel.App
{
    /// <summary>
    /// Erreurs de communication avec Word (l'application pilote Word depuis un autre processus) :
    /// Word occupé par une boîte de dialogue, Word fermé entre-temps… traduites en messages clairs.
    /// </summary>
    internal static class ComErrors
    {
        /// <summary>RPC_E_CALL_REJECTED : Word refuse l'appel (occupé).</summary>
        public const int CallRejected = unchecked((int)0x80010001);
        /// <summary>RPC_E_SERVERCALL_RETRYLATER : Word demande de réessayer plus tard (boîte de dialogue ouverte).</summary>
        public const int RetryLater = unchecked((int)0x8001010A);
        /// <summary>RPC_E_DISCONNECTED : l'objet n'existe plus (document ou Word fermé).</summary>
        public const int Disconnected = unchecked((int)0x80010108);
        /// <summary>RPC_E_SERVER_DIED.</summary>
        public const int ServerDied = unchecked((int)0x80010007);
        /// <summary>RPC_E_SERVER_DIED_DNE.</summary>
        public const int ServerDiedNotExecuted = unchecked((int)0x80010012);
        /// <summary>RPC_S_SERVER_UNAVAILABLE : le processus Word a disparu.</summary>
        public const int ServerUnavailable = unchecked((int)0x800706BA);
        /// <summary>RPC_S_CALL_FAILED.</summary>
        public const int CallFailed = unchecked((int)0x800706BE);
        /// <summary>MK_E_UNAVAILABLE : Word n'est pas lancé.</summary>
        public const int Unavailable = unchecked((int)0x800401E3);

        /// <summary>SERVERCALL_RETRYLATER (IMessageFilter).</summary>
        public const int ServerCallRetryLater = 2;

        /// <summary>
        /// Décision du filtre de messages quand Word refuse un appel : délai avant nouvel essai (ms),
        /// ou -1 pour abandonner (appel refusé définitivement, ou délai maximal dépassé).
        /// </summary>
        /// <param name="rejectType">SERVERCALL_REJECTED (1) ou SERVERCALL_RETRYLATER (2).</param>
        /// <param name="elapsedMs">Temps écoulé depuis le premier essai.</param>
        /// <param name="timeoutMs">Délai maximal d'attente.</param>
        public static int RetryDelay(int rejectType, int elapsedMs, int timeoutMs)
        {
            if (rejectType != ServerCallRetryLater) return -1;
            if (elapsedMs < 0 || elapsedMs >= timeoutMs) return -1;
            return 200;
        }

        /// <summary>Code HRESULT de la première erreur COM de la chaîne d'exceptions ; 0 s'il n'y en a pas.</summary>
        public static int HResultOf(Exception ex)
        {
            for (var e = ex; e != null; e = e.InnerException)
            {
                var external = e as ExternalException;
                if (external != null) return external.ErrorCode;
            }
            return 0;
        }

        public static bool IsBusy(int hresult)
        {
            return hresult == CallRejected || hresult == RetryLater;
        }

        public static bool IsGone(int hresult)
        {
            return hresult == Disconnected || hresult == ServerDied || hresult == ServerDiedNotExecuted
                || hresult == ServerUnavailable || hresult == CallFailed;
        }

        /// <summary>Message clair pour les erreurs de communication connues ; null pour les autres erreurs.</summary>
        public static string FriendlyMessage(Exception ex)
        {
            int hr = HResultOf(ex);
            if (IsBusy(hr))
            {
                return "Word est occupé et ne répond pas : une boîte de dialogue est probablement ouverte dans Word "
                     + "(enregistrement, impression, mot de passe…) ou une saisie est en cours.\n\n"
                     + "Terminez-la dans Word, puis recommencez.";
            }
            if (IsGone(hr))
            {
                return "Word a été fermé ou ne répond plus.\n\nRouvrez le document dans Word, puis recommencez.";
            }
            return null;
        }
    }
}
