using System.Collections.Generic;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Export
{
    /// <summary>Tableau du document et sa légende éventuelle.</summary>
    public sealed class TableEntry
    {
        /// <summary>Rang du tableau dans le document (base 1).</summary>
        public int Index;
        public string Caption;
        public CaptionPosition CaptionPosition;
        /// <summary>Page où commence le tableau (0 : inconnue).</summary>
        public int StartPage;
        /// <summary>Page où se termine le tableau (0 : inconnue).</summary>
        public int EndPage;

        public bool HasCaption
        {
            get { return !string.IsNullOrEmpty(Caption); }
        }
    }

    /// <summary>Tableau retenu pour l'export, avec le nom de sa feuille Excel.</summary>
    public sealed class PlannedSheet
    {
        public TableEntry Table;
        public string SheetName;
    }

    /// <summary>
    /// Sélection des tableaux (option A : uniquement ceux qui ont une légende ; option B : tous ; dans les deux cas
    /// sans les tableaux décochés) et attribution des noms de feuilles. Utilisé à l'identique pour l'aperçu et pour l'export.
    /// </summary>
    public static class ExportPlan
    {
        /// <summary>Tableaux proposés par l'option A ou B (avant les cases décochées).</summary>
        public static bool IsCandidate(TableEntry table, bool allTables)
        {
            return allTables || table.HasCaption;
        }

        /// <param name="excluded">Rang des tableaux décochés par l'utilisateur (null : aucun).</param>
        public static List<PlannedSheet> Build(IEnumerable<TableEntry> tables, bool allTables, ICollection<int> excluded = null)
        {
            var names = new SheetNameBuilder();
            var plan = new List<PlannedSheet>();
            foreach (var table in tables)
            {
                if (!IsCandidate(table, allTables)) continue;
                if (excluded != null && excluded.Contains(table.Index)) continue;
                string fallback = SheetNameBuilder.DefaultName(table.Index);
                plan.Add(new PlannedSheet
                {
                    Table = table,
                    SheetName = names.Reserve(table.HasCaption ? table.Caption : null, fallback)
                });
            }
            return plan;
        }
    }
}
