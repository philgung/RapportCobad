using OfficeOpenXml;
using OfficeOpenXml.Table;

public static class ExcelExtensions
{
    static readonly string[] _categoriesVeteran4Plus = ["Veteran 4", "Veteran 5", "Veteran 6", "Veteran 7"];
    static readonly string[] Categories = ["Minibad", "Poussin 1", "Poussin 2", "Benjamin 1", "Benjamin 2", "Minime 1", "Minime 2", "Cadet 1", "Cadet 2", "Junior 1", "Junior 2", "Senior", "Veteran 1", "Veteran 2", "Veteran 3", "Veteran 4+"
    ];

    public static void AjouterReinscription(ExcelPackage package, IEnumerable<Rapport.JoueurDTO> joueurs)
    {
        var worksheet = package.Workbook.Worksheets.Add("Nouveaux Adherents/Réinscrits");
        worksheet.Cells[1, 1].Value = "Club";
        // compare la saison courante avec la saison précédente par avec le nombre d'adherents la saison précédente, le nombre de nouveaux adhérents
        // le nombre de départ et le nombre de réinscrits par sexe par catégorie et par club


        var col = 2;
        foreach (var category in Categories)
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

        var groupesParClub = joueurs.GroupBy(joueur => new {joueur.Club, joueur.Categorie, joueur.Sexe})
            .Select(g => new JoueurParClub
            (
                g.Key.Club,
                g.Key.Categorie,
                g.Key.Sexe,
                g.Count(a => a.EstNouveauJoueur),
                g.Count(a => a.EstReinscrit),
                g.Count(a => a.EstParti),
                g.Count(a => a.EstPartiDansUnAutreClubDuDepartement)
            ))
            .GroupBy(v => v.Club)
            .ToList();

        foreach (var groupeParClub in groupesParClub)
        {
            worksheet.Cells[row, 1].Value = groupeParClub.Key;
            var colIndex = 2;
            foreach (var categorie in Categories)
            {
                if (categorie == "Veteran 4+")
                {
                    var joueursVeteran4Plus = groupeParClub
                        .Where(v => _categoriesVeteran4Plus.Contains(v.Categorie));
                    RemplirJoueur(joueursVeteran4Plus, worksheet, row, colIndex);
                }
                else
                {
                    var joueursParCategorie = groupeParClub.Where(v => v.Categorie == categorie);
                    RemplirJoueur(joueursParCategorie, worksheet, row, colIndex);
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

    private static void RemplirJoueur(IEnumerable<JoueurParClub> joueurs, ExcelWorksheet worksheet, int row, int colIndex)
    {
        var hommes = joueurs.Where(v => v.Sexe == "H");
        if (hommes != null)
        {
            worksheet.Cells[row, colIndex].Value = hommes.Sum(_ => _.NouveauxAdherents);
            worksheet.Cells[row, colIndex + 1].Value = hommes.Sum(_ => _.Reinscrits);
            worksheet.Cells[row, colIndex + 2].Value = hommes.Sum(_ => _.Depart);
            worksheet.Cells[row, colIndex + 3].Value = hommes.Sum(_ => _.DepartDansUnAutreClub);
        }
        var femmes = joueurs.Where(v => v.Sexe == "F");
        if (femmes != null)
        {
            worksheet.Cells[row, colIndex + 4].Value = femmes.Sum(_ => _.NouveauxAdherents);
            worksheet.Cells[row, colIndex + 5].Value = femmes.Sum(_ => _.Reinscrits);
            worksheet.Cells[row, colIndex + 6].Value = femmes.Sum(_ => _.Depart);
            worksheet.Cells[row, colIndex + 7].Value = femmes.Sum(_ => _.DepartDansUnAutreClub);
        }
    }

    public static void AjouterAdherents(ExcelPackage excelPackage, IEnumerable<Rapport.ClubDTO> clubDtos)
    {
        var worksheet = excelPackage.Workbook.Worksheets.Add("Adherents");

        // Add merged headers for categories
        worksheet.Cells[1, 1].Value = "Club";
        worksheet.Cells[1, 2].Value = "Saison";

        var colIndex = 3;
        foreach (var categorie in Categories)
        {
            ConfigureHeader(worksheet, categorie, colIndex);
            colIndex += 2;
        }

        worksheet.Cells[1, 35].Value = "Nombre d'adhérents";

        var row = 3;
        foreach (var club in clubDtos)
        {
            foreach (var saison in club.Saisons)
            {
                var adherents = saison.Categories.SelectMany(c => c.Adherents).ToList();
                worksheet.Cells[row, 1].Value = club.Sigle;
                worksheet.Cells[row, 2].Value = saison.Saison;


                var categories = Categories.Select((categorie, index) => new { Name = categorie, StartCol = 3 + index * 2 })
                    .ToList();;

                foreach (var category in categories)
                {
                    if (category.Name == "Veteran 4+")
                    {
                        var joueursVeteran4Plus = adherents
                            .Where(a => _categoriesVeteran4Plus.Contains(a.Categorie)).ToList();

                        worksheet.Cells[row, category.StartCol].Value = joueursVeteran4Plus.Count(a => a.Sexe == "H");
                        worksheet.Cells[row, category.StartCol + 1].Value = joueursVeteran4Plus.Count(a => a.Sexe == "F");
                    }
                    else
                    {
                        var joueursParCategorie = adherents
                            .Where(a => a.Categorie == category.Name).ToList();
                        worksheet.Cells[row, category.StartCol].Value = joueursParCategorie.Count(a => a.Sexe == "H");
                        worksheet.Cells[row, category.StartCol + 1].Value = joueursParCategorie.Count(a => a.Sexe == "F");
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

internal record JoueurParClub(string Club, string Categorie, string Sexe, int NouveauxAdherents, int Reinscrits, int Depart, int DepartDansUnAutreClub);