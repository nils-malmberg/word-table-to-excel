using System.Collections.Generic;

namespace WordTableToExcel.Core.Model
{
    public enum CaptionPosition
    {
        None,
        Above,
        Below
    }

    /// <summary>Tableau Word extrait, prêt à être écrit dans une feuille Excel.</summary>
    public sealed class TableModel
    {
        /// <summary>Position du tableau dans le document (1 = premier tableau).</summary>
        public int DocumentIndex;
        /// <summary>Texte complet de la légende, ou null.</summary>
        public string Caption;
        public CaptionPosition CaptionPosition;
        /// <summary>Nom de la feuille Excel (renseigné par <see cref="Captions.SheetNameBuilder"/>).</summary>
        public string SheetName;

        public int RowCount;
        public int ColumnCount;
        public readonly List<CellModel> Cells = new List<CellModel>();

        /// <summary>Largeur des colonnes de la grille, en points (0 = inconnue).</summary>
        public double[] ColumnWidthsPt;
        /// <summary>Hauteur de ligne imposée par Word, en points (0 = automatique).</summary>
        public double[] RowHeightsPt;
        /// <summary>true si la hauteur de ligne Word est « exacte » (sinon « au moins »).</summary>
        public bool[] RowHeightExact;
        /// <summary>
        /// Nombre de lignes d'en-tête répétées en haut de chaque page : lues dans Word à l'export (figées dans Excel),
        /// écrites dans Word à l'import.
        /// </summary>
        public int HeaderRowCount;
        /// <summary>Pages du document Word où le tableau commence et se termine (0 : inconnues ; export, feuille Sommaire).</summary>
        public int StartPage;
        public int EndPage;
        /// <summary>Tableau de droite à gauche (feuille Excel en mode « de droite à gauche »).</summary>
        public bool RightToLeft;

        /// <summary>Remarques non bloquantes rencontrées lors de la lecture (affichées dans le rapport final).</summary>
        public readonly List<string> Warnings = new List<string>();
        /// <summary>Informations du tableau absentes du classeur (images, texte coupé…) : signalées en tête du rapport final.</summary>
        public readonly List<string> Omissions = new List<string>();

        public bool HasCaption
        {
            get { return !string.IsNullOrEmpty(Caption); }
        }
    }
}
