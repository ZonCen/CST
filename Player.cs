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
        public string name;
        public string lastName;
        public string fullName;

        public int rank;
        public int division;
        public int newPlacement;
        public int lastPlacement = 1;
        public int fakeRank =1;

        string[,] testArray = new string[10,2]; 


        public Player(string Fullname, int Rank)
        {
            fullName = Fullname;

            string[] tokens = Fullname.Split(new[] { "\n" }, StringSplitOptions.None);
            name = tokens[0];
            lastName = tokens[1];
            rank = Rank;

            if (lastPlacement >= 1 && lastPlacement <= 3)
            {
                division = 1;
            }
            else if (lastPlacement >= 4 && lastPlacement <= 6)
            {
                division = 2;
            }
            else if (lastPlacement >= 7 && lastPlacement <= 9)
            {
                division = 3;
            }
            else if (lastPlacement >= 10 && lastPlacement <= 12)
            {
                division = 4;
            }
            else if (lastPlacement >= 13 && lastPlacement <= 15)
            {
                division = 5;
            }
            else if (lastPlacement >= 16 && lastPlacement <= 18)
            {
                division = 6;
            }

            getRow();
        }

        public void getRow()
        {
            ExcelHandler test = new ExcelHandler();
            int row = test.GetRowNumber(name);

            testArray[0,0] = "B" + row;
            testArray[1,0] = "D" + row;
            testArray[2,0] = "E" + row;
            testArray[3,0] = "G" + row;
            testArray[4,0] = "H" + row;
            testArray[5,0] = "J" + row;
            testArray[6,0] = "L" + row;
            testArray[7,0] = "M" + row;
            testArray[8,0] = "N" + row;
            testArray[9,0] = "O" + row;

            testArray[0, 1] = "B" + (row + 1);
            testArray[1, 1] = "D" + (row+1);
            testArray[2, 1] = "E" + (row+1);
            testArray[3, 1] = "G" + (row+1);
            testArray[4, 1] = "H" + (row+1);
            testArray[5, 1] = "J" + (row+1);
            testArray[6, 1] = "L" + (row+1);
            testArray[7, 1] = "M" + (row+1);
            testArray[8, 1] = "N" + (row+1);
            testArray[9, 1] = "O" + (row+1);
        }

        public void Rapport(int Player1Score1, int player2Score1, int player1Score2, int player2Score2, int player2LastPlacement, int player2Division)
        {

            string firstColum = "";
            string secondColumn = "";
            string lowerFirstColumn = "";
            string lowerSecondColumn = "";
            if(division > 1)
                fakeRank = CalculateFakeRank();
            int p2Rank = CalculateFakeRank(player2LastPlacement, player2Division);

            if ((fakeRank == 1 || fakeRank == 3) && p2Rank == 2)
            {
                firstColum = testArray[2, 0];
                secondColumn = testArray[3, 0];
                lowerFirstColumn = testArray[2, 1];
                lowerSecondColumn = testArray[3, 1];
            }
            else if ((fakeRank == 1 || fakeRank == 2) && p2Rank == 3)
            {
                firstColum = testArray[4, 0];
                secondColumn = testArray[5, 0];
                lowerFirstColumn = testArray[4, 1];
                lowerSecondColumn = testArray[5, 1];
            }
            else if ((fakeRank == 2 || fakeRank == 3)&& p2Rank == 1)
            {
                firstColum = testArray[0, 0];
                secondColumn = testArray[1, 0];
                lowerFirstColumn = testArray[0, 1];
                lowerSecondColumn = testArray[1, 1];
            }

            ExcelHandler score = new ExcelHandler();
            score.testUpdateTableWithNewPlayerClass(firstColum, secondColumn, lowerFirstColumn, lowerSecondColumn, Player1Score1, player2Score1, player1Score2, player2Score2);
        }

        public int CalculateFakeRank()
        {
 
          return  lastPlacement - (3 * (division - 1));
 
        }

        public int CalculateFakeRank(int p2LastPlacement, int p2Division)
        {
            if (p2LastPlacement > 3)
            {
                int temp = p2LastPlacement - (3 * (p2Division - 1));
                return temp;
            }
            return p2LastPlacement;
        }

    }
}
