// définir les joueurs nouveaux des joueurs déjà présents

using OfficeOpenXml;
using static Affichage.Parser;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


var rapport = CreerRapport(@"C:\Users\philippe.gung\CompetitionJeunes");

var deuxDernieresSaisons = rapport.Saisons.TakeLast(2);
var precedenteSaison = deuxDernieresSaisons.First();
var saison = deuxDernieresSaisons.Last();
ComparerSaisons(precedenteSaison, saison);



Console.WriteLine($"{saison.NomDeLaSaison} : {saison.Competitions.Count} compétitions");
Console.WriteLine($"{saison.PourcentageDeCompetiteursUnique:P} de compétiteurs uniques");
Console.WriteLine($"{saison.Competitions.Count(c => c.EstUnPAD)} PAD avec {saison.Competitions.Where(c => c.EstUnPAD).SelectMany(c => c.Competiteurs).Count()} compétiteurs");
Console.WriteLine($"{saison.Competitions.Count(c => c.EstUnTNT)} TNT avec {saison.Competitions.Where(c => c.EstUnTNT).SelectMany(c => c.Competiteurs).Count()} compétiteurs");

Console.WriteLine($"{rapport.Saisons.Count} saisons");

Console.WriteLine($"{saison.Adherents.Count} adherents");


void ComparerSaisons(Saison precedente, Saison courante)
{
    Console.WriteLine($"{precedente.NomDeLaSaison}");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnMinibad)} minibad");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnPoussin)} poussin");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnBenjamin)} benjamin");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnMinime)} minime");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnCadet)} cadet");
    Console.WriteLine($"{precedente.Adherents.Count(a => a.EstUnJunior)} junior");
    Console.WriteLine($"Ratio entre compétiteurs et adhérents : {precedente.NombreDeCompetiteursUnique} compétiteurs unique pour {precedente.Adherents.Count} adhérents");

    Console.WriteLine($"{courante.NomDeLaSaison}");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnMinibad)} minibad");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnPoussin)} poussin");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnBenjamin)} benjamin");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnMinime)} minime");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnCadet)} cadet");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnJunior)} junior");
    Console.WriteLine($"Ratio entre compétiteurs et adhérents : {courante.NombreDeCompetiteursUnique} compétiteurs pour {courante.Adherents.Count} adhérents");

    Console.WriteLine($"Taux de renouvellement : {precedente.Adherents.Count(a => courante.Adherents.Any(c => c.NumeroDeLicence == a.NumeroDeLicence)) / (decimal)precedente.Adherents.Count:P}");

}


public record Adherent(
    string Sexe,
    string Nom,
    string Prenom,
    string NumeroDeLicence,
    string SigleClub,
    string Categorie)
{
    public bool EstUnMinibad => Categorie.Contains("minibad", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnPoussin => Categorie.Contains("poussin", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnBenjamin => Categorie.Contains("benjamin", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnMinime => Categorie.Contains("minime", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnCadet => Categorie.Contains("cadet", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnJunior => Categorie.Contains("junior", StringComparison.InvariantCultureIgnoreCase);
}


public record Competiteur(string Nom, string Prenom, string NumeroDeLicence, string SigleClub, string Categorie);

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

public record Competition(string NomDeLaCompetition)
{
    public IList<Competiteur> Competiteurs { get; init; } = new List<Competiteur>();
    public bool EstUnTNT => NomDeLaCompetition.Contains("TNT", StringComparison.InvariantCultureIgnoreCase);
    public bool EstUnPAD => NomDeLaCompetition.Contains("PAD", StringComparison.InvariantCultureIgnoreCase);
}

public record Rapport
{
    public List<Saison> Saisons { get; } = new();
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