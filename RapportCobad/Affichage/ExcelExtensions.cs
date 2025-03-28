using OfficeOpenXml;
using OfficeOpenXml.Table;

public static class ExcelExtensions
{
    public static void AjouterAdherents(ExcelPackage excelPackage, IEnumerable<Rapport.ClubDTO> clubDtos)
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
                worksheet.Cells[row, 5].Value =
                    adherents.Count(a => a.Categorie.Contains("Poussin 1") && a.Sexe == "H");
                worksheet.Cells[row, 6].Value =
                    adherents.Count(a => a.Categorie.Contains("Poussin 1") && a.Sexe == "F");
                worksheet.Cells[row, 7].Value =
                    adherents.Count(a => a.Categorie.Contains("Poussin 2") && a.Sexe == "H");
                worksheet.Cells[row, 8].Value =
                    adherents.Count(a => a.Categorie.Contains("Poussin 2") && a.Sexe == "F");
                worksheet.Cells[row, 9].Value =
                    adherents.Count(a => a.Categorie.Contains("Benjamin 1") && a.Sexe == "H");
                worksheet.Cells[row, 10].Value =
                    adherents.Count(a => a.Categorie.Contains("Benjamin 1") && a.Sexe == "F");
                worksheet.Cells[row, 11].Value =
                    adherents.Count(a => a.Categorie.Contains("Benjamin 2") && a.Sexe == "H");
                worksheet.Cells[row, 12].Value =
                    adherents.Count(a => a.Categorie.Contains("Benjamin 2") && a.Sexe == "F");
                worksheet.Cells[row, 13].Value =
                    adherents.Count(a => a.Categorie.Contains("Minime 1") && a.Sexe == "H");
                worksheet.Cells[row, 14].Value =
                    adherents.Count(a => a.Categorie.Contains("Minime 1") && a.Sexe == "F");
                worksheet.Cells[row, 15].Value =
                    adherents.Count(a => a.Categorie.Contains("Minime 2") && a.Sexe == "H");
                worksheet.Cells[row, 16].Value =
                    adherents.Count(a => a.Categorie.Contains("Minime 2") && a.Sexe == "F");
                worksheet.Cells[row, 17].Value = adherents.Count(a => a.Categorie.Contains("Cadet 1") && a.Sexe == "H");
                worksheet.Cells[row, 18].Value = adherents.Count(a => a.Categorie.Contains("Cadet 1") && a.Sexe == "F");
                worksheet.Cells[row, 19].Value = adherents.Count(a => a.Categorie.Contains("Cadet 2") && a.Sexe == "H");
                worksheet.Cells[row, 20].Value = adherents.Count(a => a.Categorie.Contains("Cadet 2") && a.Sexe == "F");
                worksheet.Cells[row, 21].Value =
                    adherents.Count(a => a.Categorie.Contains("Junior 1") && a.Sexe == "H");
                worksheet.Cells[row, 22].Value =
                    adherents.Count(a => a.Categorie.Contains("Junior 1") && a.Sexe == "F");
                worksheet.Cells[row, 23].Value =
                    adherents.Count(a => a.Categorie.Contains("Junior 2") && a.Sexe == "H");
                worksheet.Cells[row, 24].Value =
                    adherents.Count(a => a.Categorie.Contains("Junior 2") && a.Sexe == "F");
                worksheet.Cells[row, 25].Value = adherents.Count(a => a.Categorie.Contains("Senior") && a.Sexe == "H");
                worksheet.Cells[row, 26].Value = adherents.Count(a => a.Categorie.Contains("Senior") && a.Sexe == "F");
                worksheet.Cells[row, 27].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 1") && a.Sexe == "H");
                worksheet.Cells[row, 28].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 1") && a.Sexe == "F");
                worksheet.Cells[row, 29].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 2") && a.Sexe == "H");
                worksheet.Cells[row, 30].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 2") && a.Sexe == "F");
                worksheet.Cells[row, 31].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 3") && a.Sexe == "H");
                worksheet.Cells[row, 32].Value =
                    adherents.Count(a => a.Categorie.Contains("Veteran 3") && a.Sexe == "F");
                worksheet.Cells[row, 33].Value = adherents.Count(a =>
                    (a.Categorie.Contains("Veteran 4") || a.Categorie.Contains("Veteran 5") ||
                     a.Categorie.Contains("Veteran 6") ||
                     a.Categorie.Contains("Veteran 7") || a.Categorie.Contains("Veteran 8")) && a.Sexe == "H");
                worksheet.Cells[row, 34].Value = adherents.Count(a =>
                    (a.Categorie.Contains("Veteran 4") || a.Categorie.Contains("Veteran 5") ||
                     a.Categorie.Contains("Veteran 6") ||
                     a.Categorie.Contains("Veteran 7") || a.Categorie.Contains("Veteran 8")) && a.Sexe == "F");
                worksheet.Cells[row, 35].Value = adherents.Count;
                row++;
            }
        }

        // Format as table
        var range = worksheet.Cells[2, 1, row - 1, 35];
        var table = worksheet.Tables.Add(range, "AdherentsTable");
        table.TableStyle = TableStyles.Medium9;
    }

    static void ConfigureHeader(ExcelWorksheet excelWorksheet, string categorie, int colInitial)
    {
        excelWorksheet.Cells[1, colInitial, 1, colInitial + 1].Merge = true;
        excelWorksheet.Cells[1, colInitial].Value = categorie;
        excelWorksheet.Cells[2, colInitial].Value = "H";
        excelWorksheet.Cells[2, colInitial + 1].Value = "F";
    }
}