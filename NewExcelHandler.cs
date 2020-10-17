using Microsoft.Win32;
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
        FileInfo file = new FileInfo(@"C:\Users\ZonC\OneDrive\Dokument\Test\CST 2020 Serie 3.xlsx");
        List<int> rows = new List<int>();
        public int round = 1;
        //Misc 

        public int checkRound()
        {
            using (var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["E7:O7"]
                              where cell.Value != null
                              select cell.Value);
                round = query1.Count() + 1;
                return round;
            }
        }

        public List<int> GetRounds(int row)
        {
            List<int> rounds = new List<int>();
            using (var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["E" + row + ":O" + row]
                              where cell.Value != null
                              select cell.Value);

                foreach (var q in query1)
                {
                    rounds.Add(Convert.ToInt32(q));
                }
            }
            return rounds;

        }


        public int GetRowNumber(string Name, int round)
        {
            using (var package = new ExcelPackage(file))
            {
                try
                {
                    var query1 = (from cell in package.Workbook.Worksheets["S30" + round].Cells["A12:A80"]
                                  where cell.Value != null && cell.Value.ToString() == Name
                                  select cell.Start.Row).Last();
                    return Convert.ToInt32(query1);
                }
                catch
                {
                    MessageBox.Show("Fel! i GetRowNumber " + Name);
                }

                return 1;
            }

        }

        public int GetRankingRowNumber(string Name, int round)
        {
            using (var package = new ExcelPackage(file))
            {
                try
                {
                    var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C24"]
                                  where cell.Value != null && cell.Value.ToString() == Name
                                  select cell.Start.Row).Last();
                    return Convert.ToInt32(query1);
                }
                catch
                {
                    MessageBox.Show("Fel! i GetRankingRowNumber " + Name);
                }

                return 1;
            }

        }

        public void test(string sheet, string column)
        {
            using (var package = new ExcelPackage(file))
            {
                var cells = (from cell in package.Workbook.Worksheets[sheet].Cells[column + ":" + column]
                             where cell.Value != null && cell.Value.ToString() != "Ranking"
                             select cell.Text);
            }
        }

        public void fillRows(List<Player> temp)
        {
            foreach (var p in temp)
            {
                rows.Add(p.row);
            }

            rows.Sort();
        }



        //Section Scoretables

        public void UpdatePlacement(List<Player> list)
        {
            using (var package = new ExcelPackage(file))
            {
                for (int i = 0; i < list.Count(); i++)
                {
                    package.Workbook.Worksheets["S30" + round].Cells["O" + list[i].row].Value = list[i].newPlacement.ToString();
                }
                package.Save();
            }
        }

        public void UpdateCurrentPlacement(List<Player> list)
        {
            list = list.OrderByDescending(x => x.wins).
                        ThenByDescending(y => y.game).
                        ThenByDescending(z => z.score).ToList();
            int i = 1;
            foreach (var p in list)
            {
                switch (p.division)
                {
                    case 1:
                        p.newPlacement = i;
                        i++;
                        break;
                    case 2:
                        p.newPlacement = i + 3;
                        i++;
                        break;
                    case 3:
                        p.newPlacement = i + 6;
                        i++;
                        break;
                    case 4:
                        p.newPlacement = i + 9;
                        i++;
                        break;
                    case 5:
                        p.newPlacement = i + 12;
                        i++;
                        break;
                    case 6:
                        p.newPlacement = i + 15;
                        i++;
                        break;
                }
            }
        }

        public void UpdateScoreTable(string row, string row2, string match2row, string match2row2, int Player1Score1, int player2Score1, int player1Score2, int player2Score2)
        {
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

        public int CheckPlacement(string column, int round)
        {
            int temp = 0;


            using (var package = new ExcelPackage(file))
            {
                temp = Convert.ToInt32(package.Workbook.Worksheets["S30" + round].Cells[column].Value);

            }
            return temp;
        }

        public int CheckLastPlacement(int round, string name)
        {
            int temp = 0;
            string column = "O" + GetRowNumber(name, round);
            temp = CheckPlacement(column, round);
            return temp;
        }

        public void nextWeekPlacement(List<Player> list)
        {
            foreach(var p in list)
            {
                if(p.newPlacement == 3 || p.newPlacement == 6 || p.newPlacement == 9 || p.newPlacement == 12)
                {
                    p.rank = p.newPlacement + 1;
                }
                else if(p.newPlacement == 4 || p.newPlacement == 7 || p.newPlacement == 10 || p.newPlacement == 13 || p.newPlacement == 16)
                {
                    p.rank = p.newPlacement - 1;
                }
                else
                {
                    p.rank = p.newPlacement;
                }
            }
        }

        public void UpdateNextWeek(int round, List<Player> list)
        {
            nextWeekPlacement(list);
            foreach (var p in list)
            {
                if (round != 6)
                {
                    using (var package = new ExcelPackage(file))
                    {
                        int row1 = 12;
                        if (p.rank > 1)
                        {
                            row1 = rows[p.rank - 1];
                        }
                        int row2 = row1 + 1;
                        package.Workbook.Worksheets["S30" + (round + 1)].Cells["A" + row1].Value = p.name;
                        package.Workbook.Worksheets["S30" + (round + 1)].Cells["A" + row2].Value = p.lastName;
                        package.Save();
                    }
                }
            }


        }

        public void UpdateWins(int row, int round, int wins, int games, int score)
        {
            using (var package = new ExcelPackage(file))
            {
                package.Workbook.Worksheets["S30" + (round)].Cells["L" + row].Value = wins;
                package.Workbook.Worksheets["S30" + (round)].Cells["M" + row].Value = games;
                package.Workbook.Worksheets["S30" + (round)].Cells["N" + row].Value = score;

                package.Save();
            }
        }

        //Ranking tabellen
        public List<Player> ImportPlayers()
        {
            List<Player> temp = new List<Player>();

            using (var package = new ExcelPackage(file))
            {
                var query1 = (from cell in package.Workbook.Worksheets["Ranking"].Cells["C7:C80"]
                              where cell.Value != null
                              select cell.Value);
                foreach (var p in query1)
                {
                    Player newSpelare = new Player(p.ToString(), round);
                    temp.Add(newSpelare);
                }
            }
            return temp;
        }
        public void CalculatePlacement(List<Player> list)
        {
            List<Player> Division1 = new List<Player>();
            List<Player> Division2 = new List<Player>();
            List<Player> Division3 = new List<Player>();
            List<Player> Division4 = new List<Player>();
            List<Player> Division5 = new List<Player>();
            List<Player> Division6 = new List<Player>();

            foreach (var p in list)
            {
                switch (p.division)
                {
                    case 1:
                        Division1.Add(p);
                        break;
                    case 2:
                        Division2.Add(p);
                        break;
                    case 3:
                        Division3.Add(p);
                        break;
                    case 4:
                        Division4.Add(p);
                        break;
                    case 5:
                        Division5.Add(p);
                        break;
                    case 6:
                        Division6.Add(p);
                        break;
                    default:
                        break;
                }

            }

            Division1 = Division1.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.score).ToList();
            Division2 = Division2.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.score).ToList();
            Division3 = Division3.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.score).ToList();
            Division4 = Division4.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.wins).ToList();
            Division5 = Division5.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.score).ToList();
            Division6 = Division6.OrderByDescending(x => x.wins).
                                  ThenByDescending(y => y.game).
                                  ThenByDescending(z => z.score).ToList();


            List<Player> players = new List<Player>();
            players.AddRange(Division1);
            players.AddRange(Division2);
            players.AddRange(Division3);
            players.AddRange(Division4);
            players.AddRange(Division5);
            players.AddRange(Division6);

            for (int i = 0; i < players.Count(); i++)
            {
                players[i].newPlacement = i + 1;
            }

            //foreach(var p in players)
            //{
            //    MessageBox.Show(p.name + " har placering " + p.newPlacement);
            //}
        }



        public void UpdateRankTable(List<Player> list)
        {
            using (var package = new ExcelPackage(file))
            {
                //Skriv alltid till E7, F7, H7, J7, L7, N7
                for (int i = 0; i < list.Count(); i++)
                {
                    if (round == 1)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].newPlacement;
                    }
                    else if (round == 2)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].newPlacement;

                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rounds[0];
                    }
                    else if (round == 3)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = list[i].newPlacement;

                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rounds[0];
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].rounds[1];
                    }
                    else if (round == 4)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = list[i].newPlacement;

                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rounds[0];
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].rounds[1];
                        package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = list[i].rounds[2];
                    }
                    else if (round == 5)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["L" + (i + 7).ToString()].Value = list[i].newPlacement;

                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rounds[0];
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].rounds[1];
                        package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = list[i].rounds[2];
                        package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = list[i].rounds[3];
                    }
                    else if (round == 6)
                    {
                        package.Workbook.Worksheets["Ranking"].Cells["C" + (i + 7).ToString()].Value = list[i].name.Trim() + Environment.NewLine + list[i].lastName.Trim();
                        package.Workbook.Worksheets["Ranking"].Cells["N" + (i + 7).ToString()].Value = list[i].newPlacement;

                        package.Workbook.Worksheets["Ranking"].Cells["E" + (i + 7).ToString()].Value = list[i].rounds[0];
                        package.Workbook.Worksheets["Ranking"].Cells["F" + (i + 7).ToString()].Value = list[i].rounds[1];
                        package.Workbook.Worksheets["Ranking"].Cells["H" + (i + 7).ToString()].Value = list[i].rounds[2];
                        package.Workbook.Worksheets["Ranking"].Cells["J" + (i + 7).ToString()].Value = list[i].rounds[3];
                        package.Workbook.Worksheets["Ranking"].Cells["L" + (i + 7).ToString()].Value = list[i].rounds[4];
                    }
                }
                package.Save();
            }
        }
    }
}
