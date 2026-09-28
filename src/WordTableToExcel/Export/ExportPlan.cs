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
    /// Sélection des tableaux (option A : uniquement ceux qui ont une légende ; option B : tous)
    /// et attribution des noms de feuilles. Utilisé à l'identique pour l'aperçu et pour l'export.
    /// </summary>
    public static class ExportPlan
    {
        public static List<PlannedSheet> Build(IEnumerable<TableEntry> tables, bool allTables)
        {
            var names = new SheetNameBuilder();
            var plan = new List<PlannedSheet>();
            foreach (var table in tables)
            {
                if (!allTables && !table.HasCaption) continue;
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
