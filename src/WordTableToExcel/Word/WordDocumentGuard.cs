using System;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Garantit que l'export laisse le document exactement dans l'état où il était :
    /// l'extension ne fait que lire, mais certaines lectures (pagination, XML) peuvent
    /// marquer le document comme « modifié ». L'indicateur est alors rétabli, pour que Word
    /// ne propose pas d'enregistrer un document qui n'a pas changé.
    /// </summary>
    public sealed class WordDocumentGuard : IDisposable
    {
        private readonly dynamic _document;
        private readonly bool _wasSaved;
        private readonly bool _known;

        public WordDocumentGuard(object document)
        {
            _document = document;
            try
            {
                _wasSaved = WordCom.IsTrue(_document.Saved);
                _known = true;
            }
            catch (Exception)
            {
                _known = false;
            }
        }

        public void Dispose()
        {
            if (!_known || !_wasSaved) return;
            try
            {
                if (!WordCom.IsTrue(_document.Saved)) _document.Saved = true;
            }
            catch (Exception)
            {
                // Vue protégée / document en lecture seule : rien à rétablir.
            }
        }
    }
}
