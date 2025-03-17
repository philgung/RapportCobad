using System.Globalization;
using CsvHelper;
using OfficeOpenXml;

namespace Affichage;

public static class Parser
{
    public static Rapport CreerRapport(string chemin)
    {
        var rapport = new Rapport();
        foreach (string repertoireSaison in Directory.GetDirectories(chemin))
        {
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
        foreach (var fichierAdherents in Directory.GetFiles(repertoireSaison, "*.csv"))
        {
            using (var reader = new StreamReader(fichierAdherents))
            using (var csv = new CsvReader(reader, new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
                   {
                       Delimiter = ";",
                   }))
            {
                var adherentDto = csv.GetRecords<AdherentDTO>();
                foreach (var dto in adherentDto)
                {
                    saison.Adherents.Add(new Adherent(dto.Sexe, dto.Nom, dto.Prenom, dto.Licence, dto.Sigle, dto.Categorie));
                }
            }
        }

    }

    private static void AjouterLesCompetitions(string repertoireSaison, Saison saison)
    {
        foreach (var fichierCompetition in Directory.GetFiles(repertoireSaison, "*.xlsx"))
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

