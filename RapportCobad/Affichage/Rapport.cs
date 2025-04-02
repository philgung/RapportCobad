public record Rapport
{
    public List<Saison> Saisons { get; } = new();

    public IEnumerable<ClubDTO> GetAdherentsParClub() {
        return Saisons.SelectMany(s => s.Adherents)
            .GroupBy(a => a.SigleClub)
            .Select(g => new ClubDTO(g.Key, g.GroupBy(a => a.Saison)
                .Select(g => new SaisonDTO(g.Key, g.GroupBy(a => a.Categorie)
                    .Select(g => new CategorieDTO(g.Key, g.ToList())).ToList())
                ).ToList()));
    }
    // avoir liste des adhérents ayant toutes les saisons
    public IEnumerable<JoueurDTO> ObtenirJoueurs()
    {
        var lesSaisons = Saisons.Select(_ => _.NomDeLaSaison).Distinct();
        var premiereSaison = lesSaisons.Order().First().Substring(2, 5);
        var secondeSaison = lesSaisons.Order().Last().Substring(2, 5);
        var adherentsNonFiltree = Saisons.SelectMany(s => s.Adherents);

        return adherentsNonFiltree
            .GroupBy(a => a.NumeroDeLicence)
            .Select(grouped =>
            {
                var saisonPrecedente = grouped.FirstOrDefault(_ => _.Saison == premiereSaison);
                var saisonCourante = grouped.FirstOrDefault(_ => _.Saison == secondeSaison);

                if (saisonPrecedente != null && saisonCourante != null)
                {
                    if (saisonPrecedente.SigleClub == saisonCourante.SigleClub)
                    {
                        // est réinscrit dans le même club
                        return new JoueurDTO(saisonPrecedente.NumeroDeLicence,
                            saisonPrecedente.Nom, saisonPrecedente.Prenom, saisonPrecedente.Sexe,
                            saisonPrecedente.Categorie, saisonPrecedente.SigleClub, false, true, false, false);
                    }

                    // est parti dans un autre club du département
                    return new JoueurDTO(saisonPrecedente.NumeroDeLicence,
                        saisonPrecedente.Nom, saisonPrecedente.Prenom, saisonPrecedente.Sexe,
                        saisonPrecedente.Categorie, saisonPrecedente.SigleClub, false, false, true, false);
                }

                // est arrivé à la seconde saison
                if (saisonPrecedente == null && saisonCourante != null)
                {
                    return new JoueurDTO(saisonCourante.NumeroDeLicence,
                        saisonCourante.Nom, saisonCourante.Prenom, saisonCourante.Sexe,
                        saisonCourante.Categorie, saisonCourante.SigleClub, true, false, false, false);
                }

                // est parti
                if (saisonPrecedente != null && saisonCourante == null)
                {
                    return new JoueurDTO(saisonPrecedente.NumeroDeLicence,
                        saisonPrecedente.Nom, saisonPrecedente.Prenom, saisonPrecedente.Sexe,
                        saisonPrecedente.Categorie, saisonPrecedente.SigleClub, false, false, false, true);
                }

                throw new InvalidOperationException();

            });

    }

    public record ClubDTO(string Sigle, IList<SaisonDTO> Saisons);
    public record SaisonDTO(string Saison, IList<CategorieDTO> Categories);
    public record CategorieDTO(string Categorie, IList<Adherent> Adherents);
    public record JoueurDTO(string NumeroLicence, string Nom, string Prenom, string Sexe, string Categorie, string Club, bool EstNouveauJoueur, bool EstReinscrit, bool EstPartiDansUnAutreClubDuDepartement, bool EstParti);


    public decimal PourcentageDeCompetiteursAyantRenouveleDUneSaisonSurLAutre => Saisons
        .Zip(Saisons.Skip(1), (s1, s2) => (s1, s2))
        .SelectMany(t => t.s1.Competitions.SelectMany(c => c.Competiteurs)
            .Select(c => c.NumeroDeLicence)
            .Distinct()
            .Intersect(t.s2.Competitions.SelectMany(c => c.Competiteurs)
                .Select(c => c.NumeroDeLicence)
                .Distinct()))
        .Count() / (decimal)Saisons.Sum(s => s.Competitions.Sum(c => c.Competiteurs.Count));


}

