using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableToExcel.Word
{
    /// <summary>
    /// Suivi des modifications et renvois de notes, lus via le modèle objet de Word : passages à ne pas exporter
    /// (texte supprimé, accepté ou non ; résultat d'un champ NOTEREF), sous forme d'intervalles de positions.
    /// </summary>
    public static class WordRevisions
    {
        /// <summary>wdRevisionDelete.</summary>
        public const int Delete = 2;
        /// <summary>wdRevisionMovedFrom (texte déplacé : son ancienne place).</summary>
        public const int MovedFrom = 14;
        /// <summary>wdRevisionConflictDelete.</summary>
        public const int ConflictDelete = 21;
        /// <summary>wdFieldNoteRef : renvoi vers une note de bas de page ou de fin.</summary>
        public const int FieldNoteRef = 72;

        /// <summary>
        /// Le document contient-il des modifications suivies (acceptées ou non) ? Une seule question à Word : sans
        /// modification suivie, toutes les recherches de suppressions (par tableau, par légende…) sont inutiles, et
        /// Word peut être lent à les faire. En cas de doute, on répond oui (les recherches ont lieu).
        /// </summary>
        public static bool DocumentHasRevisions(object documentObject)
        {
            try
            {
                return WordCom.AsInt(((dynamic)documentObject).Revisions.Count) > 0;
            }
            catch (Exception)
            {
                return true;
            }
        }

        /// <summary>Intervalles [début, fin[ supprimés en suivi des modifications dans la plage (acceptés ou non).</summary>
        public static List<int[]> DeletedIntervals(object rangeObject)
        {
            var list = new List<int[]>();
            dynamic range = rangeObject;
            dynamic revisions = range.Revisions;
            if (WordCom.AsInt(revisions.Count) == 0) return list;
            foreach (dynamic revision in revisions)
            {
                int type = WordCom.AsInt(revision.Type);
                if (type != Delete && type != MovedFrom && type != ConflictDelete) continue;
                dynamic r = revision.Range;
                int start = WordCom.AsInt(r.Start);
                int end = WordCom.AsInt(r.End);
                list.Add(new[] { start, end });
            }
            return Normalize(list);
        }

        /// <summary>Intervalles occupés par le résultat des renvois vers une note (champs NOTEREF) dans la plage.</summary>
        public static List<int[]> NoteReferenceIntervals(object rangeObject)
        {
            var list = new List<int[]>();
            dynamic range = rangeObject;
            dynamic fields = range.Fields;
            if (WordCom.AsInt(fields.Count) == 0) return list;
            foreach (dynamic field in fields)
            {
                if (WordCom.AsInt(field.Type) != FieldNoteRef) continue;
                dynamic result = field.Result;
                int start = WordCom.AsInt(result.Start);
                int end = WordCom.AsInt(result.End);
                list.Add(new[] { start, end });
            }
            return Normalize(list);
        }

        /// <summary>Texte de la plage sans les passages supprimés en suivi des modifications.</summary>
        public static string VisibleText(object rangeObject)
        {
            List<int[]> deleted;
            try
            {
                deleted = DeletedIntervals(rangeObject);
            }
            catch (Exception)
            {
                deleted = new List<int[]>();
            }
            return TextOutside(rangeObject, deleted);
        }

        /// <summary>
        /// Texte de la plage hors des intervalles exclus : découpé dans le texte de la plage quand chaque position
        /// correspond à un caractère (un seul appel à Word), sinon lu morceau par morceau.
        /// </summary>
        public static string TextOutside(object rangeObject, List<int[]> excluded)
        {
            dynamic range = rangeObject;
            string text = WordCom.AsString(range.Text);
            if (excluded == null || excluded.Count == 0) return text;
            int start = WordCom.AsInt(range.Start);
            int end = WordCom.AsInt(range.End);
            var pieces = Subtract(start, end, excluded);
            var sb = new StringBuilder();
            if (text.Length == end - start)
            {
                foreach (var piece in pieces) sb.Append(text, piece[0] - start, piece[1] - piece[0]);
                return sb.ToString();
            }
            dynamic document = range.Document;
            foreach (var piece in pieces) sb.Append(WordCom.AsString(document.Range(piece[0], piece[1]).Text));
            return sb.ToString();
        }

        /// <summary>
        /// Tableau supprimé en suivi des modifications (accepté ou non) : il porte des suppressions et, en dehors
        /// d'elles, il ne reste aucun texte (seulement les marques de cellules et de lignes).
        /// </summary>
        public static bool IsDeletedTable(object tableObject)
        {
            dynamic range = ((dynamic)tableObject).Range;
            var deleted = DeletedIntervals((object)range);
            if (deleted.Count == 0) return false;
            string remaining = TextOutside((object)range, deleted);
            return Core.Text.CellTextSanitizer.Comparable(remaining).Length == 0;
        }

        /// <summary>Vrai si les intervalles couvrent entièrement [début, fin[.</summary>
        public static bool Covers(List<int[]> intervals, int start, int end)
        {
            return end > start && Subtract(start, end, intervals).Count == 0;
        }

        /// <summary>Intervalles triés, sans chevauchement, non vides.</summary>
        public static List<int[]> Normalize(IEnumerable<int[]> intervals)
        {
            var sorted = intervals.Where(i => i != null && i.Length == 2 && i[1] > i[0]).OrderBy(i => i[0]).ToList();
            var result = new List<int[]>();
            foreach (var interval in sorted)
            {
                var last = result.Count > 0 ? result[result.Count - 1] : null;
                if (last != null && interval[0] <= last[1]) last[1] = Math.Max(last[1], interval[1]);
                else result.Add(new[] { interval[0], interval[1] });
            }
            return result;
        }

        /// <summary>Parties de [début, fin[ hors des intervalles exclus, dans l'ordre.</summary>
        public static List<int[]> Subtract(int start, int end, List<int[]> excluded)
        {
            var pieces = new List<int[]>();
            int position = start;
            foreach (var interval in Normalize(excluded ?? new List<int[]>()))
            {
                if (interval[1] <= position) continue;
                if (interval[0] >= end) break;
                if (interval[0] > position) pieces.Add(new[] { position, interval[0] });
                position = Math.Max(position, interval[1]);
                if (position >= end) break;
            }
            if (position < end) pieces.Add(new[] { position, end });
            return pieces;
        }
    }
}
