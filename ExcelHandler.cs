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

    public class ExcelHandler
    {
        //FileInfo file = new FileInfo(@"C:\Users\ZonC\OneDrive\Dokument\Test\CST 2020 Serie 2.xlsx");
        FileInfo file = new FileInfo(@"C:\Users\ZonC\OneDrive\Dokument\Test\CST-20-Serie-2.xlsx");
        //Hämta information från excel
        public List<Spelare> importPlayers()
        {
            List<Spelare> temp = new List<Spelare>();

            using(var package = new ExcelPackage(file))
            {
                var cells = package.Workbook.Worksheets["Ranking"].Cells["C7:C19"].Value.ToString();

                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                              where cell.Value != null
                              select cell.Value);
                int i = 1;
                foreach(var p in query1)
                {
                    Spelare newSpelare = new Spelare(p.ToString(), i);
                    temp.Add(newSpelare);
                    i++;
                }
                    

            }

            return temp;
        }

        //Ska användas
        public void testUpdateTableWithNewPlayerClass(string row, string row2, string match2row, string match2row2, int Player1Score1, int player2Score1, int player1Score2, int player2Score2)
        {
            int round = 1;
            using(var package = new ExcelPackage(file))
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
        //ska användas
        public int GetRowNumber(string Name)
        {
            using(var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["S301"].Cells["A7:A80"]
                              where cell.Value != null && cell.Value.ToString() == Name
                              select cell.Start.Row).Last();

                return Convert.ToInt32(query1);
            }

        }

    public List<RankInfo> ImportRankingTable(List<Spelare> list)
        {
            List<RankInfo> tempRank = new List<RankInfo>();
            using(var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                              where cell.Value != null
                              select cell.Value);

                int i = 0;
                foreach(var p in query1)
                {
                    if(p.ToString() == list[i].fullname)
                    {
                        var playerRow = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                                 where cell.Text == list[i].fullname.Trim()
                                 select cell.Start.Row).First();

                        RankInfo newPlayer = new RankInfo();
                        if(package.Workbook.Worksheets["Ranking"].Cells["E" + playerRow].Value != null)
                        {
                            newPlayer.player = list[i];
                            newPlayer.omgång1 = Convert.ToInt32(package.Workbook.Worksheets["Ranking"].Cells["E" + playerRow].Value);
                        }
                        if(package.Workbook.Worksheets["Ranking"].Cells["F" + playerRow].Value != null)
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

        //Uppdatera tabellen
        public void UpdateTable(List<Spelare> ListOfDivisions)
        {
            List<Spelare> listOfDivisions = ListOfDivisions;
            int round = 1;
            using(var package = new ExcelPackage(file))
            {
                for(int i = 0; i < listOfDivisions.Count(); i++)
                {
                    int fakeRank = listOfDivisions[i].rank;
                    if(fakeRank > 3)
                        fakeRank = fakeRank - (3 * (listOfDivisions[i].division - 1));

                    if (listOfDivisions[i].division == 1)
                    {
                        if(fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A12"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A13"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A14"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A15"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A16"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A17"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    if (listOfDivisions[i].division == 2)
                    {
                        if (fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A22"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A23"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A24"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A25"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A26"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A27"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    if (listOfDivisions[i].division == 3)
                    {
                        if (fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A32"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A33"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A34"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A35"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A36"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A37"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    if (listOfDivisions[i].division == 4)
                    {
                        if (fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A50"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A51"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A52"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A53"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A54"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A55"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    if (listOfDivisions[i].division == 5)
                    {
                        if (fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A60"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A61"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A62"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A63"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A64"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A65"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    if (listOfDivisions[i].division == 6)
                    {
                        if (fakeRank == 1)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A70"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A71"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 2)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A72"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A73"].Value = listOfDivisions[i].lastname;
                        }
                        else if (fakeRank == 3)
                        {
                            package.Workbook.Worksheets["S30" + round].Cells["A74"].Value = listOfDivisions[i].name;
                            package.Workbook.Worksheets["S30" + round].Cells["A75"].Value = listOfDivisions[i].lastname;
                        }
                    }
                    package.Save();
                }


            }
        }

        //Uppdatera Ranking
        public void updateRank(List<Spelare> list, int Round, List<RankInfo> OldRankInfo)
        {
            var round = Round;
            List < RankInfo > oldRankInfo= OldRankInfo;
            using (var package = new ExcelPackage(file))
            {
                //Skriv alltid till E7, F7, H7, J7, L7, N7
                for (int i = 0; i < list.Count(); i++)
                {
                    if (round == 1)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rank;
                    }
                    else if (round == 2)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].rank;

                        foreach(var r in oldRankInfo)
                        {

                            if(list[i].fullname == r.player.fullname)
                            {
                                package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = r.omgång1;
                            }
                        }


                    }
                    else if (round == 3)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = list[i].rank;

                        foreach (var r in oldRankInfo)
                        {

                            if (list[i].fullname == r.player.fullname)
                            {
                                package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = r.omgång1;
                                package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = r.omgång2;
                            }
                        }
                    }
                    else if (round == 4)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = list[i].rank;

                        foreach (var r in oldRankInfo)
                        {

                            if (list[i].fullname == r.player.fullname)
                            {
                                package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = r.omgång1;
                                package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = r.omgång2;
                                package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = r.omgång3;
                            }
                        }
                    }
                    else if (round == 5)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["L" + (i + 7).ToString()].Value = list[i].rank;

                        foreach (var r in oldRankInfo)
                        {

                            if (list[i].fullname == r.player.fullname)
                            {
                                package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = r.omgång1;
                                package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = r.omgång2;
                                package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = r.omgång3;
                                package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = r.omgång4;
                            }
                        }
                    }
                    else if (round == 6)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastname.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["N" + (i + 7).ToString()].Value = list[i].rank;

                        foreach (var r in oldRankInfo)
                        {

                            if (list[i].fullname == r.player.fullname)
                            {
                                package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = r.omgång1;
                                package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = r.omgång2;
                                package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = r.omgång3;
                                package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = r.omgång4;
                                package.Workbook.Worksheets["Ranking"].Cells["L" + (i + 7).ToString()].Value = r.omgång5;
                            }
                        }
                    }
                }
                package.Save();
            }
        }

        //Uppdatera poäng
        public void updateScore(Spelare Player1, Spelare Player2, int resultat1Spelare1Match1, int resultatSpelare2Match1, int resultatSpelare1Match2, int resultatSpelare2Match2)
        {
            int round = 1;
            Spelare player1 = Player1;
            int rankSpelare1 = Player1.rank;
            Spelare player2 = Player2;
            int rankSpelare2 = Player2.rank;

            int p1Row;
            int p2Row;

            if (player1.rank > 3)
                rankSpelare1 = player1.rank - (3 * (player1.division -1));
            if (player2.rank > 3)
                rankSpelare2 = player2.rank - (3 * (player2.division - 1));

            using (var package = new ExcelPackage(file))
            {
                p1Row = (from cell in package.Workbook.Worksheets["S30" + round].Cells["A11:A57"]
                              where cell.Text == player1.name.Trim()
                              select cell.Start.Row).First();
                p2Row = (from cell in package.Workbook.Worksheets["S30" + round].Cells["A11:A57"]
                              where cell.Text == player2.name.Trim()
                              select cell.Start.Row).First();

                if (rankSpelare1 == 1)
                {

                    if (rankSpelare2 == 2)
                    {
                        //spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["E" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["E" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;


                        //Spelare2
                        package.Workbook.Worksheets["S30" + round].Cells["B" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["B" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;

                    }
                    else if (rankSpelare2 == 3)
                    {
                        //spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["H" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["H" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;

                        //spelare 2
                        package.Workbook.Worksheets["S30" + round].Cells["B" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["B" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;
                    }
                    package.Save();
                }
                else if (rankSpelare1 == 2)
                {
                    if (rankSpelare2 == 1)
                    {
                        //Spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["B" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["B" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;

                        //Spelare 2
                        package.Workbook.Worksheets["S30" + round].Cells["E" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["E" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;
                    }
                    else if (rankSpelare2 == 3)
                    {
                        //Spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["H" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["H" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;

                        //spelare 2
                        package.Workbook.Worksheets["S30" + round].Cells["E" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["E" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;
                    }
                    package.Save();
                }
                else if (rankSpelare1 == 3)
                {
                    if (rankSpelare2 == 1)
                    {
                        //Spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["B" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["B" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["D" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;

                        //spelare 2
                        package.Workbook.Worksheets["S30" + round].Cells["H" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["H" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;
                    }
                    else if (rankSpelare2 == 2)
                    {
                        //Spelare 1
                        package.Workbook.Worksheets["S30" + round].Cells["E" + p1Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + p1Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["E" + (p1Row + 1)].Value = resultatSpelare1Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["G" + (p1Row + 1)].Value = resultatSpelare2Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p1Row].Value = player1.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p1Row].Value = player1.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p1Row].Value = player1.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p1Row].Value = player1.rank;

                        //Spelare 2
                        package.Workbook.Worksheets["S30" + round].Cells["H" + p2Row].Value = resultatSpelare2Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + p2Row].Value = resultat1Spelare1Match1;
                        package.Workbook.Worksheets["S30" + round].Cells["H" + (p2Row + 1)].Value = resultatSpelare2Match2;
                        package.Workbook.Worksheets["S30" + round].Cells["J" + (p2Row + 1)].Value = resultatSpelare1Match2;

                        package.Workbook.Worksheets["S30" + round].Cells["N" + p2Row].Value = player2.matchWon;
                        package.Workbook.Worksheets["S30" + round].Cells["O" + p2Row].Value = player2.gameWon;
                        package.Workbook.Worksheets["S30" + round].Cells["P" + p2Row].Value = player2.pointDifference;
                        package.Workbook.Worksheets["S30" + round].Cells["Q" + p2Row].Value = player2.rank;
                    }
                    package.Save();
                }
            }
         
        }

        public void UpdatePlacement()
        {
            //Uppdatera placement efter alla matcher är spelade.

        }
    }
}
