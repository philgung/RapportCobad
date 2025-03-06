// définir les joueurs nouveaux des joueurs déjà présents

using OfficeOpenXml;
using static Affichage.Parser;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


// ajouter par saison le nombre de compétiteurs par rapport aux autres licenciés jeunes

var rapport = CreerRapport(@"C:\Users\philippe.gung\CompetitionJeunes");



foreach (var saison in rapport.Saisons)
{
    Console.WriteLine($"{saison.NomDeLaSaison} : {saison.Competitions.Count} compétitions");
    Console.WriteLine($"{saison.PourcentageDeCompetiteursUnique:P} de compétiteurs uniques");
    Console.WriteLine($"{saison.PourcentageDeCompetiteursAyantFaitNCompetitions(1):P} compétiteurs ayant fait 1 compétition");
    Console.WriteLine($"{saison.PourcentageDeCompetiteursAyantFaitNCompetitions(2):P} compétiteurs ayant fait 2 compétitions");
    Console.WriteLine($"{saison.PourcentageDeCompetiteursAyantFaitNCompetitions(3):P} compétiteurs ayant fait 3 compétitions");
    Console.WriteLine($"{saison.PourcentageDeCompetiteursAyantFaitNCompetitions(4):P} compétiteurs ayant fait 4 compétitions");
}
Console.WriteLine($"{rapport.Saisons.Count} saisons");
Console.WriteLine($"{rapport.PourcentageDeCompetiteursAyantRenouveleDUneSaisonSurLAutre:P} de compétiteurs ayant renouvelé d'une saison sur l'autre");





public record Competiteur(string Nom, string Prenom, string NumeroDeLicence, string SigleClub, string Categorie);

public record Saison(string NomDeLaSaison)
{
    public IList<Competition> Competitions { get; } = new List<Competition>();

    public decimal PourcentageDeCompetiteursUnique => NombreDeCompetiteursUnique / (decimal)NombreDeCompetiteurs;

    public decimal PourcentageDeCompetiteursAyantFaitNCompetitions(int n) => Competitions
        .SelectMany(c => c.Competiteurs)
        .GroupBy(c => c.NumeroDeLicence)
        .Count(g => g.Count() == n) / (decimal)NombreDeCompetiteurs;

    private int NombreDeCompetiteursUnique => Competitions
        .SelectMany(c => c.Competiteurs)
        .Select(c => c.NumeroDeLicence)
        .Distinct().Count();

    private int NombreDeCompetiteurs => Competitions.Sum(c => c.Competiteurs.Count);
}

public record Competition(string NomDeLaCompetition)
{
    public IList<Competiteur> Competiteurs { get; init; } = new List<Competiteur>();
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