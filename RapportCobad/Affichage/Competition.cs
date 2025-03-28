public record Competition(string NomDeLaCompetition)
{
    public IList<Competiteur> Competiteurs { get; init; } = new List<Competiteur>();
    public bool EstUnTNT => NomDeLaCompetition.Contains("TNT", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnPAD => NomDeLaCompetition.Contains("PAD", StringComparison.InvariantCultureIgnoreCase);
}