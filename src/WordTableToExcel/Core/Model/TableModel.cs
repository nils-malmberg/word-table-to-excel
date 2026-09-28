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

        /// <summary>Remarques non bloquantes rencontrées lors de la lecture (affichées dans le rapport final).</summary>
        public readonly List<string> Warnings = new List<string>();

        public bool HasCaption
        {
            get { return !string.IsNullOrEmpty(Caption); }
        }
    }
}
