using OfficeOpenXml;
using static Affichage.Parser;
using static ExcelExtensions;

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
    Console.WriteLine(
        $"Ratio entre compétiteurs et adhérents : {precedente.NombreDeCompetiteursUnique} compétiteurs unique pour {precedente.Adherents.Count} adhérents");

    Console.WriteLine($"{courante.NomDeLaSaison}");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnMinibad)} minibad");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnPoussin)} poussin");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnBenjamin)} benjamin");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnMinime)} minime");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnCadet)} cadet");
    Console.WriteLine($"{courante.Adherents.Count(a => a.EstUnJunior)} junior");
    Console.WriteLine(
        $"Ratio entre compétiteurs et adhérents : {courante.NombreDeCompetiteursUnique} compétiteurs pour {courante.Adherents.Count} adhérents");

    Console.WriteLine(
        $"Taux de renouvellement : {precedente.Adherents.Count(a => courante.Adherents.Any(c => c.NumeroDeLicence == a.NumeroDeLicence)) / (decimal)precedente.Adherents.Count:P}");
}