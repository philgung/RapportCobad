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
        var adherentsNonFiltree = Saisons.SelectMany(s => s.Adherents);

        return adherentsNonFiltree
            .GroupBy(a => a.NumeroDeLicence)
            .Select(g =>
                new JoueurDTO(g.Key, g.First().Nom, g.First().Prenom,
                    g.Select(_ => new JoueurParSaisonDTO(_.Saison, _.SigleClub, _.Categorie))));
    }

    public record ClubDTO(string Sigle, IList<SaisonDTO> Saisons);
    public record SaisonDTO(string Saison, IList<CategorieDTO> Categories);
    public record CategorieDTO(string Categorie, IList<Adherent> Adherents);
    public record JoueurDTO(string NumeroLicence, string Nom, string Prenom, IEnumerable<JoueurParSaisonDTO> Saisons);

    public record JoueurParSaisonDTO(string Saison, string Club, string Categorie);

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

