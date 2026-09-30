using System.Collections.Generic;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Layout
{
    /// <summary>
    /// Cellule positionnée dans la grille du tableau, avec sa mise en forme de cellule
    /// (fond, bordures, alignement vertical, orientation) déjà résolue.
    /// </summary>
    public sealed class LayoutCell
    {
        public int Row;
        public int Column;
        public int RowSpan = 1;
        public int ColumnSpan = 1;

        /// <summary>Ligne Word (base 0) où commence la cellule.</summary>
        public int SourceRow;
        /// <summary>Rang de la cellule parmi les cellules visibles de sa ligne Word (base 0).</summary>
        public int SourceOrdinal;
        /// <summary>Rang de la cellule parmi toutes les cellules de sa ligne, y compris les continuations de fusion verticale (base 0).</summary>
        public int SourceIndex;
        /// <summary>Clé libre permettant de relier la cellule à sa source (utilisée par le mode de repli COM).</summary>
        public object SourceTag;

        public Rgb? Fill;
        /// <summary>Trame du premier paragraphe tramé de la cellule (utilisée si la cellule n'a pas de fond).</summary>
        public Rgb? ParagraphShading;
        public CellBorders Borders = new CellBorders();
        public VerticalAlignment VerticalAlignment = VerticalAlignment.Top;
        public int TextRotation;

        /// <summary>Texte et mise en forme lus dans le XML (null si le XML ne le permet pas).</summary>
        public XmlCellContent Content;

        public int LastRow
        {
            get { return Row + RowSpan - 1; }
        }

        public int LastColumn
        {
            get { return Column + ColumnSpan - 1; }
        }

        public override string ToString()
        {
            return string.Format("[{0},{1} {2}x{3}] src {4}/{5}", Row, Column, RowSpan, ColumnSpan, SourceRow, SourceOrdinal);
        }
    }

    /// <summary>Structure d'un tableau : grille régulière, fusions, dimensions.</summary>
    public sealed class TableLayout
    {
        public int RowCount;
        public int ColumnCount;
        public readonly List<LayoutCell> Cells = new List<LayoutCell>();
        /// <summary>Largeur de chaque colonne de grille en points (0 = inconnue).</summary>
        public double[] ColumnWidthsPt;
        /// <summary>Hauteur de ligne Word en points (0 = automatique).</summary>
        public double[] RowHeightsPt;
        public bool[] RowHeightExact;
        /// <summary>Nombre de cellules (y compris continuations de fusion verticale) de chaque ligne Word (lignes supprimées comprises).</summary>
        public int[] SourceCellCounts;
        /// <summary>Nombre de lignes du tableau dans Word, lignes supprimées en suivi des modifications comprises.</summary>
        public int SourceRowCount;
        /// <summary>Rang (base 0) des lignes supprimées en suivi des modifications, absentes de la grille.</summary>
        public readonly HashSet<int> DeletedRows = new HashSet<int>();
        /// <summary>true si fond et bordures proviennent de l'analyse XML (sinon ils restent à lire via COM).</summary>
        public bool HasCellFormatting;
        /// <summary>true si le texte et la mise en forme des caractères de chaque cellule ont été lus dans le XML.</summary>
        public bool HasContent;
        /// <summary>Texte visible du tableau selon le XML (pour vérification avec le texte renvoyé par Word).</summary>
        public string XmlText;
        /// <summary>
        /// Variantes acceptées du texte du tableau selon ce que Word inclut dans Range.Text : avec ou sans le texte
        /// supprimé en suivi des modifications, avec ou sans le texte masqué.
        /// </summary>
        public readonly List<string> XmlTextVariants = new List<string>();
        /// <summary>Images et objets (dessins, formes, zones de texte, objets OLE) du tableau : non exportés.</summary>
        public int ObjectCount;
        /// <summary>Lignes d'en-tête du haut du tableau répétées sur chaque page (w:tblHeader).</summary>
        public int HeaderRowCount;

        /// <summary>Cellules visibles qui commencent sur la ligne Word donnée, dans l'ordre.</summary>
        public List<LayoutCell> CellsStartingOnSourceRow(int sourceRow)
        {
            var list = new List<LayoutCell>();
            foreach (var cell in Cells)
            {
                if (cell.SourceRow == sourceRow) list.Add(cell);
            }
            list.Sort((a, b) => a.SourceOrdinal.CompareTo(b.SourceOrdinal));
            return list;
        }
    }
}
