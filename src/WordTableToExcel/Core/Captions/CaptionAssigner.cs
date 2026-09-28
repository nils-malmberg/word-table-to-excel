using System.Collections.Generic;
using WordTableToExcel.Core.Model;

namespace WordTableToExcel.Core.Captions
{
    /// <summary>Paragraphe candidat au rôle de légende (voisin immédiat d'un tableau).</summary>
    public sealed class CaptionCandidate
    {
        public CaptionCandidate(int key, string text, bool fromSequenceField)
        {
            Key = key;
            Text = text;
            FromSequenceField = fromSequenceField;
        }

        /// <summary>Identifiant unique du paragraphe (sa position de début dans le document).</summary>
        public readonly int Key;
        public readonly string Text;
        public readonly bool FromSequenceField;
    }

    /// <summary>Légendes candidates au-dessus et au-dessous d'un tableau.</summary>
    public sealed class TableCaptionContext
    {
        public int TableIndex;
        public CaptionCandidate Above;
        public CaptionCandidate Below;
    }

    public sealed class CaptionAssignment
    {
        public static readonly CaptionAssignment NoCaption = new CaptionAssignment(null, CaptionPosition.None);

        public CaptionAssignment(CaptionCandidate caption, CaptionPosition position)
        {
            Caption = caption;
            Position = position;
        }

        public readonly CaptionCandidate Caption;
        public readonly CaptionPosition Position;
    }

    /// <summary>
    /// Associe chaque tableau à sa légende en déterminant automatiquement la convention du
    /// document (légendes au-dessus ou au-dessous). Une même légende n'est jamais attribuée
    /// à deux tableaux : entre deux tableaux consécutifs, elle revient à celui qui respecte
    /// la convention majoritaire.
    /// </summary>
    public static class CaptionAssigner
    {
        public static CaptionAssignment[] Assign(IList<TableCaptionContext> tables, out CaptionPosition convention)
        {
            convention = DetectConvention(tables);
            bool preferAbove = convention != CaptionPosition.Below;

            var result = new CaptionAssignment[tables.Count];
            var claimed = new HashSet<int>();

            // 1er passage : côté conforme à la convention.
            for (int i = 0; i < tables.Count; i++)
            {
                var candidate = preferAbove ? tables[i].Above : tables[i].Below;
                if (candidate != null && claimed.Add(candidate.Key))
                {
                    result[i] = new CaptionAssignment(candidate, preferAbove ? CaptionPosition.Above : CaptionPosition.Below);
                }
            }

            // 2e passage : l'autre côté, s'il n'est pas déjà pris.
            for (int i = 0; i < tables.Count; i++)
            {
                if (result[i] != null) continue;
                var candidate = preferAbove ? tables[i].Below : tables[i].Above;
                if (candidate != null && !IsPreferredCaptionOfAnotherTable(tables, i, candidate, preferAbove) && claimed.Add(candidate.Key))
                {
                    result[i] = new CaptionAssignment(candidate, preferAbove ? CaptionPosition.Below : CaptionPosition.Above);
                }
                else
                {
                    result[i] = CaptionAssignment.NoCaption;
                }
            }
            return result;
        }

        /// <summary>
        /// Convention majoritaire : on ne compte que les tableaux dont la légende est sans ambiguïté
        /// d'un seul côté. À égalité, « au-dessus » (valeur par défaut de Word pour les tableaux).
        /// </summary>
        public static CaptionPosition DetectConvention(IList<TableCaptionContext> tables)
        {
            int above = 0, below = 0;
            foreach (var t in tables)
            {
                if (t.Above != null && t.Below == null) above++;
                else if (t.Below != null && t.Above == null) below++;
            }
            if (above == 0 && below == 0)
            {
                // Tous ambigus : on privilégie les légendes issues de champs SEQ.
                foreach (var t in tables)
                {
                    if (t.Above != null && t.Above.FromSequenceField && (t.Below == null || !t.Below.FromSequenceField)) above++;
                    if (t.Below != null && t.Below.FromSequenceField && (t.Above == null || !t.Above.FromSequenceField)) below++;
                }
            }
            return below > above ? CaptionPosition.Below : CaptionPosition.Above;
        }

        private static bool IsPreferredCaptionOfAnotherTable(IList<TableCaptionContext> tables, int self, CaptionCandidate candidate, bool preferAbove)
        {
            for (int j = 0; j < tables.Count; j++)
            {
                if (j == self) continue;
                var preferred = preferAbove ? tables[j].Above : tables[j].Below;
                if (preferred != null && preferred.Key == candidate.Key) return true;
            }
            return false;
        }
    }
}
