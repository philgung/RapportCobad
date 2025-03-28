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

    public record ClubDTO(string Sigle, IList<SaisonDTO> Saisons);
    public record SaisonDTO(string Saison, IList<CategorieDTO> Categories);
    public record CategorieDTO(string Categorie, IList<Adherent> Adherents);

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