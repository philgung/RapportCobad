using OfficeOpenXml;
using OfficeOpenXml.Table;
using static Affichage.Parser;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


var rapport = CreerRapport(@"C:\Users\philippe.gung\Rapport");

// Grouper par club, puis par saison, puis par catégorie d'age
var adherentsParClub = rapport.GetAdherentsParClub();


// Adhérents qui se réinscrive : Adhérents n-1 de tous les clubs dans toutes les catégories
// Nouveaux adhérents (H/F) par catégorie (poussin1/poussin2) par club
// Compétiteurs / non compétiteurs par catégorie par club
// adhérents qui se réinscrive / nouveaux adhérents sur compétiteurs et non compétiteurs par club par catégorie par sexe
// Etats des lieux sur mutés par club par catégorie par sexe





using (var package = new ExcelPackage())
{
    AjouterAdherents(package, adherentsParClub);

    // var previousSeason = rapport.Saisons.OrderByDescending(s => s.NomDeLaSaison).Skip(1).FirstOrDefault();
    // var currentSeason = rapport.Saisons.OrderByDescending(s => s.NomDeLaSaison).FirstOrDefault();
    //
    // if (previousSeason != null && currentSeason != null)
    // {
    //     var previousAdherents = previousSeason.Adherents.Select(a => a.NumeroDeLicence).ToHashSet();
    //
    // }


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

void AjouterAdherents(ExcelPackage excelPackage, IEnumerable<Rapport.ClubDTO> clubDtos)
{
    var worksheet = excelPackage.Workbook.Worksheets.Add("Adherents");

    // Add merged headers for categories
    worksheet.Cells[1, 1].Value = "Club";
    worksheet.Cells[1, 2].Value = "Saison";

    ConfigureHeader(worksheet, "Minibad", 3);
    ConfigureHeader(worksheet, "Poussin 1", 5);
    ConfigureHeader(worksheet, "Poussin 2", 7);
    ConfigureHeader(worksheet, "Benjamin 1", 9);
    ConfigureHeader(worksheet, "Benjamin 2", 11);
    ConfigureHeader(worksheet, "Minime 1", 13);
    ConfigureHeader(worksheet, "Minime 2", 15);
    ConfigureHeader(worksheet, "Cadet 1", 17);
    ConfigureHeader(worksheet, "Cadet 2", 19);
    ConfigureHeader(worksheet, "Junior 1", 21);
    ConfigureHeader(worksheet, "Junior 2", 23);
    ConfigureHeader(worksheet, "Senior", 25);
    ConfigureHeader(worksheet, "Veteran 1", 27);
    ConfigureHeader(worksheet, "Veteran 2", 29);
    ConfigureHeader(worksheet, "Veteran 3", 31);
    ConfigureHeader(worksheet, "Veteran 4+", 33);
    worksheet.Cells[1, 35].Value = "Nombre d'adhérents";

    var row = 3;
    foreach (var club in clubDtos)
    {
        foreach (var saison in club.Saisons)
        {
            var adherents = saison.Categories.SelectMany(c => c.Adherents).ToList();
            worksheet.Cells[row, 1].Value = club.Sigle;
            worksheet.Cells[row, 2].Value = saison.Saison;
            worksheet.Cells[row, 3].Value = adherents.Count(a => a.Categorie.Contains("Minibad") && a.Sexe == "H");
            worksheet.Cells[row, 4].Value = adherents.Count(a => a.Categorie.Contains("Minibad") && a.Sexe == "F");
            worksheet.Cells[row, 5].Value = adherents.Count(a => a.Categorie.Contains("Poussin 1") && a.Sexe == "H");
            worksheet.Cells[row, 6].Value = adherents.Count(a => a.Categorie.Contains("Poussin 1") && a.Sexe == "F");
            worksheet.Cells[row, 7].Value = adherents.Count(a => a.Categorie.Contains("Poussin 2") && a.Sexe == "H");
            worksheet.Cells[row, 8].Value = adherents.Count(a => a.Categorie.Contains("Poussin 2") && a.Sexe == "F");
            worksheet.Cells[row, 9].Value = adherents.Count(a => a.Categorie.Contains("Benjamin 1") && a.Sexe == "H");
            worksheet.Cells[row, 10].Value = adherents.Count(a => a.Categorie.Contains("Benjamin 1") && a.Sexe == "F");
            worksheet.Cells[row, 11].Value = adherents.Count(a => a.Categorie.Contains("Benjamin 2") && a.Sexe == "H");
            worksheet.Cells[row, 12].Value = adherents.Count(a => a.Categorie.Contains("Benjamin 2") && a.Sexe == "F");
            worksheet.Cells[row, 13].Value = adherents.Count(a => a.Categorie.Contains("Minime 1") && a.Sexe == "H");
            worksheet.Cells[row, 14].Value = adherents.Count(a => a.Categorie.Contains("Minime 1") && a.Sexe == "F");
            worksheet.Cells[row, 15].Value = adherents.Count(a => a.Categorie.Contains("Minime 2") && a.Sexe == "H");
            worksheet.Cells[row, 16].Value = adherents.Count(a => a.Categorie.Contains("Minime 2") && a.Sexe == "F");
            worksheet.Cells[row, 17].Value = adherents.Count(a => a.Categorie.Contains("Cadet 1") && a.Sexe == "H");
            worksheet.Cells[row, 18].Value = adherents.Count(a => a.Categorie.Contains("Cadet 1") && a.Sexe == "F");
            worksheet.Cells[row, 19].Value = adherents.Count(a => a.Categorie.Contains("Cadet 2") && a.Sexe == "H");
            worksheet.Cells[row, 20].Value = adherents.Count(a => a.Categorie.Contains("Cadet 2") && a.Sexe == "F");
            worksheet.Cells[row, 21].Value = adherents.Count(a => a.Categorie.Contains("Junior 1") && a.Sexe == "H");
            worksheet.Cells[row, 22].Value = adherents.Count(a => a.Categorie.Contains("Junior 1") && a.Sexe == "F");
            worksheet.Cells[row, 23].Value = adherents.Count(a => a.Categorie.Contains("Junior 2") && a.Sexe == "H");
            worksheet.Cells[row, 24].Value = adherents.Count(a => a.Categorie.Contains("Junior 2") && a.Sexe == "F");
            worksheet.Cells[row, 25].Value = adherents.Count(a => a.Categorie.Contains("Senior") && a.Sexe == "H");
            worksheet.Cells[row, 26].Value = adherents.Count(a => a.Categorie.Contains("Senior") && a.Sexe == "F");
            worksheet.Cells[row, 27].Value = adherents.Count(a => a.Categorie.Contains("Veteran 1") && a.Sexe == "H");
            worksheet.Cells[row, 28].Value = adherents.Count(a => a.Categorie.Contains("Veteran 1") && a.Sexe == "F");
            worksheet.Cells[row, 29].Value = adherents.Count(a => a.Categorie.Contains("Veteran 2") && a.Sexe == "H");
            worksheet.Cells[row, 30].Value = adherents.Count(a => a.Categorie.Contains("Veteran 2") && a.Sexe == "F");
            worksheet.Cells[row, 31].Value = adherents.Count(a => a.Categorie.Contains("Veteran 3") && a.Sexe == "H");
            worksheet.Cells[row, 32].Value = adherents.Count(a => a.Categorie.Contains("Veteran 3") && a.Sexe == "F");
            worksheet.Cells[row, 33].Value = adherents.Count(a => (a.Categorie.Contains("Veteran 4") || a.Categorie.Contains("Veteran 5") || a.Categorie.Contains("Veteran 6") ||
                                                                   a.Categorie.Contains("Veteran 7") || a.Categorie.Contains("Veteran 8")) && a.Sexe == "H");
            worksheet.Cells[row, 34].Value = adherents.Count(a => (a.Categorie.Contains("Veteran 4") || a.Categorie.Contains("Veteran 5") || a.Categorie.Contains("Veteran 6") ||
                                                                   a.Categorie.Contains("Veteran 7") || a.Categorie.Contains("Veteran 8")) && a.Sexe == "F");
            worksheet.Cells[row, 35].Value = adherents.Count;
            row++;
        }
    }

    // Format as table
    var range = worksheet.Cells[2, 1, row - 1, 35];
    var table = worksheet.Tables.Add(range, "AdherentsTable");
    table.TableStyle = TableStyles.Medium9;

    void ConfigureHeader(ExcelWorksheet excelWorksheet, string categorie, int colInitial)
    {
        excelWorksheet.Cells[1, colInitial, 1, colInitial + 1].Merge = true;
        excelWorksheet.Cells[1, colInitial].Value = categorie;
        excelWorksheet.Cells[2, colInitial].Value = "H";
        excelWorksheet.Cells[2, colInitial + 1].Value = "F";
    }
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