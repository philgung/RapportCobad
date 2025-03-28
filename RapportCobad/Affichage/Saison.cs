public record Saison(string NomDeLaSaison)
{
    public IList<Competition> Competitions { get; } = new List<Competition>();

    public IList<Adherent> Adherents { get; } = new List<Adherent>();

    public decimal PourcentageDeCompetiteursUnique => NombreDeCompetiteursUnique / (decimal)NombreDeCompetiteurs;

    public decimal PourcentageDeCompetiteursAyantFaitNCompetitions(int n) => Competitions
        .SelectMany(c => c.Competiteurs)
        .GroupBy(c => c.NumeroDeLicence)
        .Count(g => g.Count() == n) / (decimal)NombreDeCompetiteurs;

    public int NombreDeCompetiteursUnique => Competitions
        .SelectMany(c => c.Competiteurs)
        .Select(c => c.NumeroDeLicence)
        .Distinct().Count();

    private int NombreDeCompetiteurs => Competitions.Sum(c => c.Competiteurs.Count);
}