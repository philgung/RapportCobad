using OfficeOpenXml;
using static Affichage.Parser;
using static ExcelExtensions;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


var rapport = CreerRapport(@"C:\Users\philippe.gung\Rapport");

var adherentsParClub = rapport.GetAdherentsParClub();

var joueurs = rapport.ObtenirJoueurs();
// Compétiteurs / non compétiteurs par catégorie par club
// adhérents qui se réinscrive / nouveaux adhérents sur compétiteurs et non compétiteurs par club par catégorie par sexe


using var package = new ExcelPackage();
AjouterAdherents(package, adherentsParClub);
AjouterReinscription(package, joueurs);

// Save the package to a file
var fileInfo = new FileInfo(@"C:\Users\philippe.gung\Rapport\Adherents.xlsx");
package.SaveAs(fileInfo);