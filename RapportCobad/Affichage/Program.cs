// définir les joueurs nouveaux des joueurs déjà présents

using OfficeOpenXml;
using OfficeOpenXml.Table;
using static Affichage.Parser;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


var rapport = CreerRapport(@"C:\Users\philippe.gung\Rapport");

// Grouper par club, puis par saison, puis par catégorie d'age
var adherentsParClub = rapport.Saisons.SelectMany(s => s.Adherents)
    .GroupBy(a => a.SigleClub)
    .Select(g => new
    {
        Club = g.Key,
        Saisons = g.GroupBy(a => a.Saison)
            .Select(g => new
            {
                Saison = g.Key,
                Categories = g.GroupBy(a => a.Categorie)
                    .Select(g => new
                    {
                        Categorie = g.Key,
                        Adherents = g.ToList()
                    }).ToList()
            }).ToList()
    }).ToList();


// Create a new Excel package
using (var package = new ExcelPackage())
{
    var worksheet = package.Workbook.Worksheets.Add("Adherents");

    // Add merged headers for categories
    worksheet.Cells[1, 1].Value = "Club";
    worksheet.Cells[1, 2].Value = "Saison";

    ConfigureHeader(worksheet, "Minibad", 3);
    ConfigureHeader(worksheet, "Poussin", 5);
    ConfigureHeader(worksheet, "Benjamin", 7);
    ConfigureHeader(worksheet, "Minime", 9);
    ConfigureHeader(worksheet, "Cadet", 11);
    ConfigureHeader(worksheet, "Junior", 13);
    ConfigureHeader(worksheet, "Senior", 15);
    ConfigureHeader(worksheet, "Veteran", 17);
    worksheet.Cells[1, 19].Value = "Nombre d'adhérents";


    int row = 3;
    foreach (var club in adherentsParClub)
    {
        foreach (var saison in club.Saisons)
        {
            var adherents = saison.Categories.SelectMany(c => c.Adherents).ToList();
            worksheet.Cells[row, 1].Value = club.Club;
            worksheet.Cells[row, 2].Value = saison.Saison;
            worksheet.Cells[row, 3].Value = adherents.Count(a => a.EstUnMinibad && a.Sexe == "H");
            worksheet.Cells[row, 4].Value = adherents.Count(a => a.EstUnMinibad && a.Sexe == "F");
            worksheet.Cells[row, 5].Value = adherents.Count(a => a.EstUnPoussin && a.Sexe == "H");
            worksheet.Cells[row, 6].Value = adherents.Count(a => a.EstUnPoussin && a.Sexe == "F");
            worksheet.Cells[row, 7].Value = adherents.Count(a => a.EstUnBenjamin && a.Sexe == "H");
            worksheet.Cells[row, 8].Value = adherents.Count(a => a.EstUnBenjamin && a.Sexe == "F");
            worksheet.Cells[row, 9].Value = adherents.Count(a => a.EstUnMinime && a.Sexe == "H");
            worksheet.Cells[row, 10].Value = adherents.Count(a => a.EstUnMinime && a.Sexe == "F");
            worksheet.Cells[row, 11].Value = adherents.Count(a => a.EstUnCadet && a.Sexe == "H");
            worksheet.Cells[row, 12].Value = adherents.Count(a => a.EstUnCadet && a.Sexe == "F");
            worksheet.Cells[row, 13].Value = adherents.Count(a => a.EstUnJunior && a.Sexe == "H");
            worksheet.Cells[row, 14].Value = adherents.Count(a => a.EstUnJunior && a.Sexe == "F");
            worksheet.Cells[row, 15].Value = adherents.Count(a => a.EstUnSenior && a.Sexe == "H");
            worksheet.Cells[row, 16].Value = adherents.Count(a => a.EstUnSenior && a.Sexe == "F");
            worksheet.Cells[row, 17].Value = adherents.Count(a => a.EstUnVeteran && a.Sexe == "H");
            worksheet.Cells[row, 18].Value = adherents.Count(a => a.EstUnVeteran && a.Sexe == "F");
            worksheet.Cells[row, 19].Value = adherents.Count;
            row++;
        }
    }

    // Format as table
    var range = worksheet.Cells[2, 1, row - 1, 19];
    var table = worksheet.Tables.Add(range, "AdherentsTable");
    table.TableStyle = TableStyles.Medium9;

    // Save the package to a file
    var fileInfo = new FileInfo(@"C:\Users\philippe.gung\Rapport\Adherents.xlsx");
    package.SaveAs(fileInfo);
}


// var deuxDernieresSaisons = rapport.Saisons.TakeLast(2);
// var precedenteSaison = deuxDernieresSaisons.First();
// var saison = deuxDernieresSaisons.Last();
// ComparerSaisons(precedenteSaison, saison);
//
//
//
// Console.WriteLine($"{saison.NomDeLaSaison} : {saison.Competitions.Count} compétitions");
// Console.WriteLine($"{saison.PourcentageDeCompetiteursUnique:P} de compétiteurs uniques");
// Console.WriteLine($"{saison.Competitions.Count(c => c.EstUnPAD)} PAD avec {saison.Competitions.Where(c => c.EstUnPAD).SelectMany(c => c.Competiteurs).Count()} compétiteurs");
// Console.WriteLine($"{saison.Competitions.Count(c => c.EstUnTNT)} TNT avec {saison.Competitions.Where(c => c.EstUnTNT).SelectMany(c => c.Competiteurs).Count()} compétiteurs");
//
// Console.WriteLine($"{rapport.Saisons.Count} saisons");
//
// Console.WriteLine($"{saison.Adherents.Count} adherents");


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

void ConfigureHeader(ExcelWorksheet excelWorksheet, string categorie, int colInitial)
{
    excelWorksheet.Cells[1, colInitial, 1, colInitial + 1].Merge = true;
    excelWorksheet.Cells[1, colInitial].Value = categorie;
    excelWorksheet.Cells[2, colInitial].Value = "H";
    excelWorksheet.Cells[2, colInitial + 1].Value = "F";
}


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