public record Adherent(
    string Sexe,
    string Nom,
    string Prenom,
    string NumeroDeLicence,
    string SigleClub,
    string Categorie,
    string Saison,
    string CompetiteurActif,
    string MeilleurPlume,
    string EstHandicape,
    string Handicap)
{
    public bool EstUnMinibad => Categorie.Contains("minibad", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnPoussin => Categorie.Contains("poussin", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnBenjamin => Categorie.Contains("benjamin", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnMinime => Categorie.Contains("minime", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnCadet => Categorie.Contains("cadet", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnJunior => Categorie.Contains("junior", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnSenior => Categorie.Contains("senior", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnVeteran => Categorie.Contains("veteran", StringComparison.InvariantCultureIgnoreCase);
}