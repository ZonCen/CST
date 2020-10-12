using OfficeOpenXml;
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace CST_With_only_excel
{
    class NewExcelHandler
    {
        FileInfo file = new FileInfo(@"C:\Users\ZonC\OneDrive\Dokument\Test\CST-20-Serie-2.xlsx");


        //Misc 
        public int GetRowNumber(string Name)
        {
            using (var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["S301"].Cells["A7:A80"]
                              where cell.Value != null && cell.Value.ToString() == Name
                              select cell.Start.Row).Last();

                return Convert.ToInt32(query1);
            }

        }

        public void test()
        {
            using (var package = new ExcelPackage(file))
            {
                var cells = (from cell in package.Workbook.Worksheets["Ranking"].Cells["A:A"]
                             where cell.Value != null && cell.Value.ToString() != "Ranking"
                             select cell);
                MessageBox.Show(cells.Last().ToString());
            }
        }

        //Section Scoretables
        public void UpdateScoreTable(string row, string row2, string match2row, string match2row2, int Player1Score1, int player2Score1, int player1Score2, int player2Score2)
        {
            int round = 1;
            using (var package = new ExcelPackage(file))
            {
                //Game 1
                package.Workbook.Worksheets["S30" + round].Cells[row].Value = Player1Score1;
                package.Workbook.Worksheets["S30" + round].Cells[row2].Value = player2Score1;

                //Game 2
                package.Workbook.Worksheets["S30" + round].Cells[match2row].Value = player1Score2;
                package.Workbook.Worksheets["S30" + round].Cells[match2row2].Value = player2Score2;

                package.Save();
            }
        }



        /* Skriv en funktion som flyttar spelare om i score tabellen efter alla matcher är avslutade:
         1. Spelare med placering 1 i divisionen går upp i division och hamnar på lägsta platsen i ovan division
         2. Spelare med placering 3 i divisionen går ner i division och hamnar på högsta platsen i nedan division
         3. Spelare med placering 2 i divisionen stannar i divisionen och hamnar i mitten av divisionen
         4. Detta skrivs in i S30 + (round+1) tabellen
            Undantag:
            Division 1 kan ej skicka spelare uppåt
            Lägsta divisionen kan ej skicka spelare neråt
            S306 är sista omgången och kan därför inte skicka det vidare till nästa scoretabell*/

        //Ranking tabellen

        //Skriv om nedan kod:
        /*Ska göra följande:
         1. Ta reda på ranken och skicka till spelareklassens rank (finns på A)
         2. Ta reda på resultat av omgångarna och lägg till i array "oldscore".*/
        public List<RankInfo> ImportRankingTable(List<Spelare> list)
        {
            List<RankInfo> tempRank = new List<RankInfo>();
            using (var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                              where cell.Value != null
                              select cell.Value);

                int i = 0;
                foreach (var p in query1)
                {
                    if (p.ToString() == list[i].fullname)
                    {
                        var playerRow = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                                         where cell.Text == list[i].fullname.Trim()
                                         select cell.Start.Row).First();

                        RankInfo newPlayer = new RankInfo();
                        if (package.Workbook.Worksheets["Ranking"].Cells["E" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång1 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["E" + playerRow].Value);
                        }
                        if (package.Workbook.Worksheets["Ranking"].Cells["F" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång2 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["F" + playerRow].Value);
                        }
                        if (package.Workbook.Worksheets["Ranking"].Cells["H" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång3 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["H" + playerRow].Value);
                        }
                        if (package.Workbook.Worksheets["Ranking"].Cells["J" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång4 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["J" + playerRow].Value);
                        }
                        if (package.Workbook.Worksheets["Ranking"].Cells["L" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång5 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["L" + playerRow].Value);
                        }
                        if (package.Workbook.Worksheets["Ranking"].Cells["N" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång6 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["N" + playerRow].Value);
                        }
                        tempRank.Add(newPlayer);

                    }
                    i++;
                }
                return tempRank;
            }

        }
    }
}
