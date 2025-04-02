using OfficeOpenXml;
using OfficeOpenXml.Table;

public static class ExcelExtensions
{
    static readonly string[] _categoriesVeteran4Plus = ["Veteran 4", "Veteran 5", "Veteran 6", "Veteran 7"];
    public static void AjouterReinscription(ExcelPackage package, IEnumerable<Rapport.JoueurDTO> joueurs)
    {
        var worksheet = package.Workbook.Worksheets.Add("Nouveaux Adherents/Réinscrits");
        worksheet.Cells[1, 1].Value = "Club";
        // compare la saison courante avec la saison précédente par avec le nombre d'adherents la saison précédente, le nombre de nouveaux adhérents
        // le nombre de départ et le nombre de réinscrits par sexe par catégorie et par club

        var categories = new[] { "Minibad", "Poussin 1", "Poussin 2", "Benjamin 1", "Benjamin 2", "Minime 1", "Minime 2", "Cadet 1", "Cadet 2", "Junior 1", "Junior 2", "Senior", "Veteran 1", "Veteran 2", "Veteran 3", "Veteran 4+" };


        var col = 2;
        foreach (var category in categories)
        {
            worksheet.Cells[1, col].Value = category;
            //worksheet.Cells[1, col, 1, col + 5].Merge = true;

            worksheet.Cells[2, col].Value = "H";
            //worksheet.Cells[2, col, 2, col + 3].Merge = true;
            worksheet.Cells[3, col].Value = "Nouveaux Adhérents";
            worksheet.Cells[3, col + 1].Value = "Réinscrits";
            worksheet.Cells[3, col + 2].Value = "Départs";
            worksheet.Cells[3, col + 3].Value = "Départs dans un autre club";

            worksheet.Cells[2, col + 4].Value = "F";
            //worksheet.Cells[2, col + 4, 2, col + 7].Merge = true;
            worksheet.Cells[3, col + 4].Value = "Nouveaux Adhérents";
            worksheet.Cells[3, col + 5].Value = "Réinscrits";
            worksheet.Cells[3, col + 6].Value = "Départs";
            worksheet.Cells[3, col + 7].Value = "Départs dans un autre club";

            col += 8;
        }

        int row = 4;
        // afficher le nombre de réinscrits par sexe par catégorie et par club

        var valeurs = joueurs.GroupBy(joueur => new {joueur.Club, joueur.Categorie, joueur.Sexe})
            .Select(g => new
            {
                g.Key.Club,
                g.Key.Categorie,
                g.Key.Sexe,
                NouveauxAdherents = g.Count(a => a.EstNouveauJoueur),
                Reinscrits = g.Count(a => a.EstReinscrit),
                Depart = g.Count(a => a.EstParti),
                DepartDansUnAutreClub = g.Count(a => a.EstPartiDansUnAutreClubDuDepartement)
            })
            .ToList();

        foreach (var groupeParClub in valeurs.GroupBy(v => v.Club))
        {
            worksheet.Cells[row, 1].Value = groupeParClub.Key;
            var colIndex = 2;
            foreach (var categorie in categories)
            {
                if (categorie == "Veteran 4+")
                {
                    var hommes = groupeParClub
                        .Where(v => _categoriesVeteran4Plus.Contains(v.Categorie) && v.Sexe == "H");
                    if (hommes != null)
                    {
                        worksheet.Cells[row, colIndex].Value = hommes.Sum(_ => _.NouveauxAdherents);
                        worksheet.Cells[row, colIndex + 1].Value = hommes.Sum(_ => _.Reinscrits);
                        worksheet.Cells[row, colIndex + 2].Value = hommes.Sum(_ => _.Depart);
                        worksheet.Cells[row, colIndex + 3].Value = hommes.Sum(_ => _.DepartDansUnAutreClub);
                    }

                    var femmes = groupeParClub
                        .Where(v => _categoriesVeteran4Plus.Contains(v.Categorie) && v.Sexe == "F");
                    if (femmes != null)
                    {
                        worksheet.Cells[row, colIndex + 4].Value = femmes.Sum(_ => _.NouveauxAdherents);
                        worksheet.Cells[row, colIndex + 5].Value = femmes.Sum(_ => _.Reinscrits);
                        worksheet.Cells[row, colIndex + 6].Value = femmes.Sum(_ => _.Depart);
                        worksheet.Cells[row, colIndex + 7].Value = femmes.Sum(_ => _.DepartDansUnAutreClub);
                    }
                }
                else
                {
                    var hommes = groupeParClub
                        .SingleOrDefault(v => v.Categorie == categorie && v.Sexe == "H");
                    if (hommes != null)
                    {
                        worksheet.Cells[row, colIndex].Value = hommes.NouveauxAdherents;
                        worksheet.Cells[row, colIndex + 1].Value = hommes.Reinscrits;
                        worksheet.Cells[row, colIndex + 2].Value = hommes.Depart;
                        worksheet.Cells[row, colIndex + 3].Value = hommes.DepartDansUnAutreClub;
                    }

                    var femmes = groupeParClub
                        .SingleOrDefault(v => v.Categorie == categorie && v.Sexe == "F");
                    if (femmes != null)
                    {
                        worksheet.Cells[row, colIndex + 4].Value = femmes.NouveauxAdherents;
                        worksheet.Cells[row, colIndex + 5].Value = femmes.Reinscrits;
                        worksheet.Cells[row, colIndex + 6].Value = femmes.Depart;
                        worksheet.Cells[row, colIndex + 7].Value = femmes.DepartDansUnAutreClub;
                    }
                }

                colIndex += 8;
            }

            row++;
        }



        // Format as table
        var range = worksheet.Cells[1, 1, row, col]; // Adjusted to start from row 1
        var table = worksheet.Tables.Add(range, "NouveauxAdherentsReinscritsTable");
        table.TableStyle = TableStyles.Medium9;

    }
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

                var categories = new[]
                {
                    new { Name = "Minibad", StartCol = 3 },
                    new { Name = "Poussin 1", StartCol = 5 },
                    new { Name = "Poussin 2", StartCol = 7 },
                    new { Name = "Benjamin 1", StartCol = 9 },
                    new { Name = "Benjamin 2", StartCol = 11 },
                    new { Name = "Minime 1", StartCol = 13 },
                    new { Name = "Minime 2", StartCol = 15 },
                    new { Name = "Cadet 1", StartCol = 17 },
                    new { Name = "Cadet 2", StartCol = 19 },
                    new { Name = "Junior 1", StartCol = 21 },
                    new { Name = "Junior 2", StartCol = 23 },
                    new { Name = "Senior", StartCol = 25 },
                    new { Name = "Veteran 1", StartCol = 27 },
                    new { Name = "Veteran 2", StartCol = 29 },
                    new { Name = "Veteran 3", StartCol = 31 },
                    new { Name = "Veteran 4+", StartCol = 33 }
                };

                foreach (var category in categories)
                {
                    if (category.Name == "Veteran 4+")
                    {
                        worksheet.Cells[row, category.StartCol].Value = adherents.Count(a => _categoriesVeteran4Plus.Contains(a.Categorie) && a.Sexe == "H");
                        worksheet.Cells[row, category.StartCol + 1].Value = adherents.Count(a => _categoriesVeteran4Plus.Contains(a.Categorie) && a.Sexe == "F");
                    }
                    else
                    {
                        worksheet.Cells[row, category.StartCol].Value = adherents.Count(a => a.Categorie.Contains(category.Name) && a.Sexe == "H");
                        worksheet.Cells[row, category.StartCol + 1].Value = adherents.Count(a => a.Categorie.Contains(category.Name) && a.Sexe == "F");
                    }
                }

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