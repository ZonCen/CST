using OfficeOpenXml;
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
    public class Player
    {
        NewExcelHandler test = new NewExcelHandler();
        public string name;
        public string lastName;
        public string fullName;

        public int rank;
        public int division;
        public int newPlacement;
        public int lastPlacement = 1;
        public int fakeRank;
        public int round = 1;
        public double average = 0;

        public int row = 0;
        public int rankingRow = 0;

        string[,] Columns = new string[10, 2];
        public string[] playedVs = new string[2];
        public List<int> rounds = new List<int>();
        public int gamePlayed = 0;

        //Score
        public int wins = 0;
        public int game = 0;
        public int score = 0;

        //Data for User
        public Player(string Fullname, int Round)
        {
            fullName = Fullname;

            string[] tokens = Fullname.Split(new[] { "\r\n" }, StringSplitOptions.None);
            name = tokens[0];
            lastName = tokens[1];
            round = test.checkRound();

            getRankingRow();
            GetRow();
            getRounds();

        }

        public void getRounds()
        {
            rounds.AddRange(test.GetRounds(rankingRow));
        }

        public void getRankingRow()
        {
            rankingRow = test.GetRankingRowNumber(fullName, round);
        }

        public void GetRow()
        {
            row = test.GetRowNumber(name, round);

            Columns[0, 0] = "B" + row;
            Columns[1, 0] = "D" + row;
            Columns[2, 0] = "E" + row;
            Columns[3, 0] = "G" + row;
            Columns[4, 0] = "H" + row;
            Columns[5, 0] = "J" + row;
            Columns[6, 0] = "L" + row;
            Columns[7, 0] = "M" + row;
            Columns[8, 0] = "N" + row;
            Columns[9, 0] = "O" + row;

            Columns[0, 1] = "B" + (row + 1);
            Columns[1, 1] = "D" + (row + 1);
            Columns[2, 1] = "E" + (row + 1);
            Columns[3, 1] = "G" + (row + 1);
            Columns[4, 1] = "H" + (row + 1);
            Columns[5, 1] = "J" + (row + 1);
            Columns[6, 1] = "L" + (row + 1);
            Columns[7, 1] = "M" + (row + 1);
            Columns[8, 1] = "N" + (row + 1);
            Columns[9, 1] = "O" + (row + 1);

            GetLastPlacement();

            if (lastPlacement == 1 ||  lastPlacement == 2 || lastPlacement == 4)
            {
                division = 1;
            }
            else if (lastPlacement == 5  || lastPlacement == 7 || lastPlacement == 3)
            {
                division = 2;
            }
            else if (lastPlacement == 8 || lastPlacement == 10 || lastPlacement == 6)
            {
                division = 3;
            }
            else if (lastPlacement == 11 || lastPlacement == 13 || lastPlacement == 9)
            {
                division = 4;
            }
            else if (lastPlacement == 14 || lastPlacement == 16 || lastPlacement == 12 || lastPlacement == 15)
            {
                division = 5;
            }
            //else if ((lastPlacement > 16 && lastPlacement < 18) || lastPlacement == 15)
            //{
            //    division = 6;
            //}
        }

        public void GetLastPlacement()
        {
            if (round > 1)
                lastPlacement = test.CheckLastPlacement((round - 1), name);
        }

        public void GetNewPlacement()
        {
            newPlacement = test.CheckPlacement(Columns[9, 0], round);
            rounds.Add(newPlacement);
        }

        public void UpdateVs(Player player2)
        {
            if (playedVs[0] != player2.fullName && playedVs[1] != player2.fullName)
            {
                if (playedVs[0] == null)
                {
                    playedVs[0] = player2.fullName;
                }
                else if (playedVs[1] == null)
                {
                    playedVs[1] = player2.fullName;
                }
            }
        }

        public void calculateAverage()
        {
            foreach (var a in rounds)
            {
                average += a;
            }

            average = average / rounds.Count;
        }


        public void Rapport(int Player1Score1, int player2Score1, int player1Score2, int player2Score2, int player2LastPlacement, int player2Division)
        {

            string firstColum = "";
            string secondColumn = "";
            string lowerFirstColumn = "";
            string lowerSecondColumn = "";
            lastPlacement = CalculateFakeRank(lastPlacement, division);

            int p2Rank = CalculateFakeRank(player2LastPlacement, player2Division);

            if ((lastPlacement == 1 || lastPlacement == 3) && p2Rank == 2)
            {
                firstColum = Columns[2, 0];
                secondColumn = Columns[3, 0];
                lowerFirstColumn = Columns[2, 1];
                lowerSecondColumn = Columns[3, 1];
            }
            else if ((lastPlacement == 1 || lastPlacement == 2) && p2Rank == 3)
            {
                firstColum = Columns[4, 0];
                secondColumn = Columns[5, 0];
                lowerFirstColumn = Columns[4, 1];
                lowerSecondColumn = Columns[5, 1];
            }
            else if ((lastPlacement == 2 || lastPlacement == 3) && p2Rank == 1)
            {
                firstColum = Columns[0, 0];
                secondColumn = Columns[1, 0];
                lowerFirstColumn = Columns[0, 1];
                lowerSecondColumn = Columns[1, 1];
            }
            else if ((lastPlacement == 1 || lastPlacement == 2) && p2Rank == 3)
            {
                firstColum = Columns[2, 0];
                secondColumn = Columns[3, 0];
                lowerFirstColumn = Columns[3, 1];
                lowerSecondColumn = Columns[3, 1];
            }

            test.UpdateScoreTable(firstColum, secondColumn, lowerFirstColumn, lowerSecondColumn, Player1Score1, player2Score1, player1Score2, player2Score2);
        }

        public int CalculateFakeRank()
        {

            return lastPlacement - 3 * (division - 1);

        }

        public int CalculateFakeRank(int p2LastPlacement, int p2Division)
        {
            if (p2LastPlacement > 2)
            {
                if(p2LastPlacement == p2Division*3 +1)
                {
                    return 3;
                }
                else if(p2LastPlacement == (p2Division - 1)*3)
                {
                    MessageBox.Show(((p2Division - 1) * 3).ToString());
                    return 1;
                }
                else
                {
                    return 2;
                }
            }
            return p2LastPlacement;
        }

        public void CalculateWins(int ScorePlayer1Game1, int ScorePlayer2Game1, int ScorePlayer1Game2, int ScorePlayer2Game2)
        {
            if (ScorePlayer1Game1 > ScorePlayer2Game1 && ScorePlayer1Game2 > ScorePlayer2Game2)
            {
                wins++;
                game += 2;
            }
            else if ((ScorePlayer1Game1 > ScorePlayer2Game1) || (ScorePlayer1Game2 > ScorePlayer2Game2))
            {
                game += 1;
            }

            int gamediffGame1 = ScorePlayer1Game1 - ScorePlayer2Game1;
            int gamediffGame2 = ScorePlayer1Game2 - ScorePlayer2Game2;
            score += gamediffGame1 + gamediffGame2;

            test.UpdateWins(row, round, wins, game, score);
        }
    }
}
