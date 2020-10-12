using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CST_With_only_excel
{
    public class Game
    {
        Spelare player1;
        Spelare player2;

        int scorePlayer1Game1;
        int scorePlayer2Game1;
        int scorePlayer1Game2;
        int scorePlayer2Game2;
        

        public void Rapport(Spelare p1, Spelare p2, int ScorePlayer1Game1, int ScorePlayer2Game1, int ScorePlayer1Game2, int ScorePlayer2Game2)
        {
            player1 = p1;
            player2 = p2;

            //Game 1
            scorePlayer1Game1 = ScorePlayer1Game1;
            scorePlayer2Game1 = ScorePlayer2Game1;

            //Game 2
            scorePlayer1Game2 = ScorePlayer1Game2;
            scorePlayer2Game2 = ScorePlayer2Game2;

            CalculateWin();
        }

        private void CalculateWin()
        {
            if (scorePlayer1Game1 > scorePlayer2Game1 && scorePlayer1Game2 > scorePlayer2Game2)
            {
                player1.gameWon += 2;
                player1.matchWon += 1;
            }
            else if (scorePlayer2Game1 > scorePlayer1Game1 && scorePlayer2Game2 > scorePlayer1Game2)
            {
                player2.gameWon += 2;
                player2.matchWon += 1;
            }
            else if (scorePlayer1Game1 > scorePlayer2Game2 || scorePlayer1Game2 > scorePlayer2Game2)
            {
                player1.gameWon += 1;
                player2.gameWon += 1;
            }
                

            player1.pointDifference += (scorePlayer1Game1 - scorePlayer2Game1) + (scorePlayer1Game2 - scorePlayer2Game2);
            player2.pointDifference += (scorePlayer2Game1 - scorePlayer1Game1) + (scorePlayer2Game2 - scorePlayer1Game2);

            ExcelHandler updateScore = new ExcelHandler();
            updateScore.updateScore(player1, player2, scorePlayer1Game1, scorePlayer2Game1, scorePlayer1Game2, scorePlayer2Game2);
        }

    }
}
