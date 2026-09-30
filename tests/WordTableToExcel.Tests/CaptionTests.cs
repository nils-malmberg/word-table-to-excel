using System.Collections.Generic;
using System.Linq;
using WordTableToExcel.Core.Captions;
using WordTableToExcel.Core.Model;
using WordTableToExcel.Export;
using Xunit;

namespace WordTableToExcel.Tests
{
    public class CaptionMatcherTests
    {
        private readonly CaptionMatcher _matcher = new CaptionMatcher();

        [Theory]
        [InlineData(" SEQ Tableau \\* ARABIC ", "Tableau")]
        [InlineData("SEQ Table \\* ARABIC \\s 1", "Table")]
        [InlineData("seq Tabla", "Tabla")]
        [InlineData(" SEQ \"Tableau\" \\* ROMAN", "Tableau")]
        [InlineData(" REF _Ref123 \\h ", null)]
        [InlineData("", null)]
        public void ParsesSequenceIdentifier(string code, string expected)
        {
            Assert.Equal(expected, CaptionMatcher.ParseSequenceIdentifier(code));
        }

        [Theory]
        [InlineData("Table")]
        [InlineData("Tableau")]
        [InlineData("TABLEAU")]
        [InlineData("Tabla")]
        [InlineData("Tabelle")]
        [InlineData("Tabella")]
        [InlineData("Tabel")]
        [InlineData("Tabela")]
        [InlineData("Tabell")]
        [InlineData("Tabulka")]
        [InlineData("Taulukko")]
        [InlineData("Táblázat")]
        [InlineData("Tablo")]
        [InlineData("Таблица")]
        [InlineData("Πίνακας")]
        [InlineData("表")]
        [InlineData("표")]
        [InlineData("Tableau_Annexe")]
        public void TableSequenceIdentifiers(string identifier)
        {
            Assert.True(_matcher.IsTableSequenceIdentifier(identifier), identifier);
        }

        [Theory]
        [InlineData("Figure")]
        [InlineData("Figura")]
        [InlineData("Abbildung")]
        [InlineData("Equation")]
        [InlineData("Illustration")]
        [InlineData("")]
        public void OtherSequenceIdentifiers(string identifier)
        {
            Assert.False(_matcher.IsTableSequenceIdentifier(identifier), identifier);
        }

        [Theory]
        [InlineData("Tableau 3 : Résultats", true)]
        [InlineData("Table 1. Results", true)]
        [InlineData("TABLE 1", true)]
        [InlineData("Tabla 2 - Datos", true)]
        [InlineData("Tabelle 4: Ergebnisse", true)]
        [InlineData("Tabl. 5 Données", true)]
        [InlineData("Tab. 3 Werte", true)]
        [InlineData("Tableau récapitulatif des coûts", true)]
        [InlineData("Tableau A-1 : Annexe", true)]
        [InlineData("Tableau IV - Synthèse", true)]
        [InlineData("Tableau n°12", true)]
        [InlineData("表1 销售额", true)]
        [InlineData("  Tableau\u00A02 : espace insécable", true)]
        [InlineData("Tablette de chocolat", false)]
        [InlineData("Tablespoon", false)]
        [InlineData("Tab", false)]
        [InlineData("Figure 2 : Carte", false)]
        [InlineData("Le tableau ci-dessous présente", false)]
        [InlineData("", false)]
        public void CaptionText(string text, bool expected)
        {
            Assert.Equal(expected, _matcher.LooksLikeCaptionText(text, false));
        }

        [Fact]
        public void CaptionStyleRelaxesNumberingRequirement()
        {
            Assert.False(_matcher.LooksLikeCaptionText("Tablespoon", false));
            Assert.True(_matcher.LooksLikeCaptionText("Tableaux comparatifs", true));
        }

        [Fact]
        public void ExtraLabels()
        {
            var matcher = new CaptionMatcher(new[] { "Annexe", " Récap. " });
            Assert.True(matcher.LooksLikeCaptionText("Annexe 2 : Coûts", false));
            Assert.True(matcher.IsTableSequenceIdentifier("Annexe"));
            Assert.True(matcher.LooksLikeCaptionText("Récap 1", false));
            Assert.False(_matcher.LooksLikeCaptionText("Annexe 2 : Coûts", false));
        }

        [Fact]
        public void CleanCaptionText_RemovesFieldCodesAndCollapsesSpaces()
        {
            string raw = "Tableau \u0013 SEQ Tableau \\* ARABIC \u00143\u0015\t:  Ventes\r";
            Assert.Equal("Tableau 3 : Ventes", CaptionMatcher.CleanCaptionText(raw));
        }
    }

    public class CaptionAssignerTests
    {
        private static CaptionCandidate C(int key, string text, bool seq = true)
        {
            return new CaptionCandidate(key, text, seq);
        }

        [Fact]
        public void CaptionsAbove()
        {
            var c1 = C(10, "Tableau 1");
            var c2 = C(50, "Tableau 2");
            var tables = new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Above = c1, Below = c2 }, // T1 et T2 se partagent c2
                new TableCaptionContext { TableIndex = 2, Above = c2 },
                new TableCaptionContext { TableIndex = 3, Above = C(90, "Tableau 3") }
            };

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(tables, out convention);

            Assert.Equal(CaptionPosition.Above, convention);
            Assert.Same(c1, result[0].Caption);
            Assert.Same(c2, result[1].Caption);
            Assert.Equal(CaptionPosition.Above, result[1].Position);
            Assert.Equal("Tableau 3", result[2].Caption.Text);
        }

        [Fact]
        public void CaptionsBelow()
        {
            var c1 = C(20, "Table 1");
            var c2 = C(60, "Table 2");
            var tables = new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Below = c1 },
                new TableCaptionContext { TableIndex = 2, Above = c1, Below = c2 }
            };

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(tables, out convention);

            Assert.Equal(CaptionPosition.Below, convention);
            Assert.Same(c1, result[0].Caption);
            Assert.Same(c2, result[1].Caption);
            Assert.Equal(CaptionPosition.Below, result[1].Position);
        }

        [Fact]
        public void CaptionIsNeverAssignedTwice()
        {
            var shared = C(30, "Tableau 1");
            var tables = new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Below = shared },
                new TableCaptionContext { TableIndex = 2, Above = shared }
            };

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(tables, out convention);

            int assigned = 0;
            foreach (var r in result) if (r.Caption != null) assigned++;
            Assert.Equal(1, assigned);
        }

        [Fact]
        public void MixedPositionsStillFindCaptions()
        {
            var tables = new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Above = C(1, "Tableau 1") },
                new TableCaptionContext { TableIndex = 2, Above = C(2, "Tableau 2") },
                new TableCaptionContext { TableIndex = 3, Below = C(3, "Tableau 3") },
                new TableCaptionContext { TableIndex = 4 }
            };

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(tables, out convention);

            Assert.Equal(CaptionPosition.Above, convention);
            Assert.Equal(CaptionPosition.Below, result[2].Position);
            Assert.Null(result[3].Caption);
            Assert.Equal(CaptionPosition.None, result[3].Position);
        }

        [Fact]
        public void AmbiguousEverywhere_PrefersSequenceFields()
        {
            var tables = new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Above = C(1, "Tableau récapitulatif", false), Below = C(2, "Tableau 1", true) }
            };

            CaptionPosition convention;
            var result = CaptionAssigner.Assign(tables, out convention);

            Assert.Equal(CaptionPosition.Below, convention);
            Assert.Equal("Tableau 1", result[0].Caption.Text);
        }

        /// <summary>
        /// Légende 1, tableau 1, tableau 2, légende 2, tableau 3 (texte entre les deux blocs), dans un document dont
        /// la plupart des légendes sont détectées au-dessous : la légende 2 revient au tableau 2 en automatique ; au
        /// tableau 3 si l'utilisateur impose « au-dessus ».
        /// </summary>
        private static List<TableCaptionContext> UserScenarioInBelowDocument(out CaptionCandidate l1, out CaptionCandidate l2)
        {
            l1 = C(1, "Tableau 1");
            l2 = C(4, "Tableau 2");
            return new List<TableCaptionContext>
            {
                new TableCaptionContext { TableIndex = 1, Above = l1 },
                new TableCaptionContext { TableIndex = 2, Below = l2 },
                new TableCaptionContext { TableIndex = 3, Above = l2 },
                new TableCaptionContext { TableIndex = 4, Below = C(8, "Tableau 4") },
                new TableCaptionContext { TableIndex = 5, Below = C(11, "Tableau 5") }
            };
        }

        [Fact]
        public void ForcedPosition_OverridesTheDetectedConvention()
        {
            CaptionCandidate l1, l2;
            var tables = UserScenarioInBelowDocument(out l1, out l2);

            CaptionPosition detected;
            var automatic = CaptionAssigner.Assign(tables, CaptionPosition.None, out detected);
            Assert.Equal(CaptionPosition.Below, detected);
            Assert.Same(l2, automatic[1].Caption);  // la confusion signalée par l'utilisateur
            Assert.Null(automatic[2].Caption);

            var above = CaptionAssigner.Assign(tables, CaptionPosition.Above, out detected);
            Assert.Equal(CaptionPosition.Below, detected); // la détection reste affichée telle quelle
            Assert.Same(l1, above[0].Caption);
            Assert.Null(above[1].Caption);
            Assert.Same(l2, above[2].Caption);
            Assert.Equal(CaptionPosition.Above, above[2].Position);
            // Légendes situées sous leur tableau et revendiquées par aucun autre : toujours trouvées.
            Assert.Equal(CaptionPosition.Below, above[3].Position);
            Assert.Equal("Tableau 5", above[4].Caption.Text);
        }

        [Fact]
        public void AssignCaptions_UpdatesTheEntriesAndReturnsTheDetectedConvention()
        {
            CaptionCandidate l1, l2;
            var contexts = UserScenarioInBelowDocument(out l1, out l2);
            var entries = contexts.Select(c => new TableEntry { Index = c.TableIndex, CaptionAbove = c.Above, CaptionBelow = c.Below }).ToList();

            Assert.Equal(CaptionPosition.Below, ExportPlan.AssignCaptions(entries, CaptionPosition.Above));
            Assert.Equal("Tableau 2", entries[2].Caption);
            Assert.Equal(CaptionPosition.Above, entries[2].CaptionPosition);
            Assert.False(entries[1].HasCaption);
            Assert.Equal("Tableau 2", ExportPlan.Build(entries, false).Single(p => p.Table.Index == 3).SheetName);

            Assert.Equal(CaptionPosition.Below, ExportPlan.AssignCaptions(entries, CaptionPosition.None));
            Assert.Equal("Tableau 2", entries[1].Caption);
            Assert.False(entries[2].HasCaption);
            Assert.Equal(CaptionPosition.None, entries[2].CaptionPosition);
        }
    }

    public class SheetNameBuilderTests
    {
        [Fact]
        public void InvalidCharactersAreReplaced()
        {
            string name = SheetNameBuilder.Sanitize("Tableau 1 : Coûts [2023]");
            Assert.Equal("Tableau 1 - Coûts -2023-", name);
            foreach (char c in "\\/?*[]:") Assert.DoesNotContain(c, SheetNameBuilder.Sanitize("a" + c + "b"));
        }

        [Fact]
        public void NamesAreTruncatedTo31Characters()
        {
            string name = SheetNameBuilder.Sanitize("Tableau 12 : Répartition des effectifs par région et par année");
            Assert.Equal(31, name.Length);
            Assert.Equal("Tableau 12 - Répartition des ef", name);
        }

        [Fact]
        public void DuplicatesAreNumbered_CaseInsensitive()
        {
            var builder = new SheetNameBuilder();
            Assert.Equal("Tableau 1", builder.Reserve("Tableau 1", "x"));
            Assert.Equal("tableau 1 (2)", builder.Reserve("tableau 1", "x"));
            Assert.Equal("Tableau 1 (3)", builder.Reserve("Tableau 1", "x"));

            string longName = new string('A', 40);
            string first = builder.Reserve(longName, "x");
            string second = builder.Reserve(longName, "x");
            Assert.Equal(31, first.Length);
            Assert.True(second.Length <= 31);
            Assert.EndsWith(" (2)", second);
        }

        [Fact]
        public void FallbackAndReservedNames()
        {
            var builder = new SheetNameBuilder();
            Assert.Equal("Tableau_4", builder.Reserve(null, SheetNameBuilder.DefaultName(4)));
            Assert.Equal("Tableau_5", builder.Reserve(" :: ", SheetNameBuilder.DefaultName(5)));
            Assert.Equal("History_", SheetNameBuilder.Sanitize("History"));
            Assert.Equal("Bilan", SheetNameBuilder.Sanitize("'Bilan'"));
            Assert.Equal("Ligne 1 Ligne 2", SheetNameBuilder.Sanitize("Ligne 1\r\nLigne 2\t"));
        }
    }
}
