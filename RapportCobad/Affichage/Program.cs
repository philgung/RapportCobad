// définir les joueurs nouveaux des joueurs déjà présents

using OfficeOpenXml;
using static Affichage.Parser;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


const string baseUrl = @"C:\Users\philippe.gung\CompetitionJeunes";
var rapport = CreerRapport(baseUrl);

foreach (var saison in rapport.Saisons)
{
    Console.WriteLine($"{saison.NomDeLaSaison} : {saison.Competitions.Count} compétitions");
    foreach (var competition in saison.Competitions)
    {
        Console.WriteLine($"  {competition.NomDeLaCompetition} : {competition.Competiteurs.Count} compétiteurs");
    }
}
Console.WriteLine($"{rapport.Saisons.Count} saisons");





public record Competiteur(string Nom, string Prenom, string NumeroDeLicence, string sigleClub, string Categorie);

public record Saison(string NomDeLaSaison)
{
    public IList<Competition> Competitions { get; } = new List<Competition>();
}

public record Competition(string NomDeLaCompetition)
{
    public IList<Competiteur> Competiteurs { get; init; } = new List<Competiteur>();
}

public record Rapport
{
    public List<Saison> Saisons { get; } = new();
}