using OfficeOpenXml;

namespace Affichage;

public static class Parser
{
    public static Rapport CreerRapport(string chemin)
    {
        var rapport = new Rapport();
        foreach (string repertoireSaison in Directory.GetDirectories(chemin))
        {
            if (repertoireSaison.EndsWith("Archive")) continue;
            var saison = CreerSaison(repertoireSaison);

            rapport.Saisons.Add(saison);
        }

        return rapport;
    }

    private static Saison CreerSaison(string repertoireSaison)
    {
        var saison = new Saison(Path.GetFileName(repertoireSaison));

        AjouterLesCompetitions(repertoireSaison, saison);
        AjouterLesAdherents(repertoireSaison, saison);

        return saison;
    }

    private static void AjouterLesAdherents(string repertoireSaison, Saison saison)
    {
        var fichierAdherents = Path.Combine(repertoireSaison, $"{saison.NomDeLaSaison}.xlsx");

        var fileInfo = new FileInfo(fichierAdherents);
        using var package = new ExcelPackage(fileInfo);
        var workbook = package.Workbook;
        var worksheet = workbook.Worksheets[0];
        var rowCount = worksheet.Dimension.Rows;

        for (var row = 2; row <= rowCount; row++)
        {
            var sexe = worksheet.Cells[row, 1].Text;
            var nom = worksheet.Cells[row, 2].Text;
            var prenom = worksheet.Cells[row, 3].Text;
            var licence = worksheet.Cells[row, 4].Text;
            var saisonLabel = worksheet.Cells[row, 5].Text;
            var sigle = worksheet.Cells[row, 6].Text;
            var categorie = worksheet.Cells[row, 7].Text;
            var competiteurActif = worksheet.Cells[row, 8].Text;
            var meilleurPlume = worksheet.Cells[row, 9].Text;
            var estHandicape = worksheet.Cells[row, 10].Text;
            var handicap = worksheet.Cells[row, 11].Text;

            if (string.IsNullOrWhiteSpace(nom) && string.IsNullOrWhiteSpace(prenom) &&
                string.IsNullOrWhiteSpace(licence) && string.IsNullOrWhiteSpace(sigle) &&
                string.IsNullOrWhiteSpace(categorie)) continue;
            saison.Adherents.Add(new Adherent(sexe, nom, prenom, licence, sigle, categorie, saisonLabel, competiteurActif, meilleurPlume, estHandicape, handicap));
        }
    }

    private static void AjouterLesCompetitions(string repertoireSaison, Saison saison)
    {
        foreach (var fichierCompetition in Directory.GetFiles(Path.Combine(repertoireSaison, "Competitions"), "*.xlsx"))
        {
            var competition = CreerCompetition(fichierCompetition);
            saison.Competitions.Add(competition);
        }
    }

    private static Competition CreerCompetition(string chemin)
    {
        var fileInfo = new FileInfo(chemin);
        Competition competition = null;
        using var package = new ExcelPackage(fileInfo);
        var workbook = package.Workbook;
        if (workbook == null || workbook.Worksheets.Count <= 0) return competition;
        var worksheet = workbook.Worksheets[0];
        var rowCount = worksheet.Dimension.Rows;

        competition = new Competition(worksheet.Cells[1, 1].Text);
        for (var row = 4; row <= rowCount; row++)
        {
            var nom = worksheet.Cells[row, 5].Text;
            var prenom = worksheet.Cells[row, 6].Text;
            var numeroDeLicence = worksheet.Cells[row, 7].Text;
            var sigleClub = worksheet.Cells[row, 3].Text;
            var categorie = worksheet.Cells[row, 11].Text;

            if (string.IsNullOrWhiteSpace(nom) && string.IsNullOrWhiteSpace(prenom) &&
                string.IsNullOrWhiteSpace(numeroDeLicence) && string.IsNullOrWhiteSpace(sigleClub) &&
                string.IsNullOrWhiteSpace(categorie)) continue;
            var competiteur = new Competiteur(nom, prenom, numeroDeLicence, sigleClub, categorie);
            competition.Competiteurs.Add(competiteur);
        }

        return competition;
    }
}

internal class AdherentDTO
{
    public string Sexe { get; set; }
    public string Nom { get; set; }
    public string Prenom { get; set; }
    public string Licence { get; set; }
    public string Sigle { get; set; }
    public string Categorie { get; set; }
}

