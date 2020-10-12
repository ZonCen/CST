using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CST_With_only_excel
{
    public partial class Form1 : Form
    {
        List<Spelare> players;
        List<RankInfo> oldRankinfo;
        public Form1()
        {
            InitializeComponent();

            ExcelHandler getData = new ExcelHandler();
            players = getData.importPlayers();
            oldRankinfo =  getData.ImportRankingTable(players);
            getData.UpdateTable(players);


            //Division testDivision = new Division(players, oldRankinfo);
            //testDivision.calculateNewRanks();

            //Funkar
            //Player newPlayer = new Player("Toni\nJak", 1);
            //Player newPlayer2 = new Player("Erik\nGullbrandsson", 2);

            //newPlayer.lastPlacement = 4;
            //newPlayer2.lastPlacement = 5;
            //newPlayer.division = 2;
            //newPlayer2.division = 2;
            //newPlayer.Rapport(11,3,11,4,newPlayer2.lastPlacement, newPlayer2.division);
            //newPlayer2.Rapport(3,11, 4,11, newPlayer.lastPlacement, newPlayer.division);


            //test hörna
            NewExcelHandler testExcel = new NewExcelHandler();
            testExcel.test();
            ReportDivision1.Enabled = false;
            ReportDivision2.Enabled = false;
            reportDivision3.Enabled = false;
            reportDivision4.Enabled = false;
            reportDivision5.Enabled = false;
            reportDivision6.Enabled = false;



            foreach (var p in players)
            {
                if (p.division == 1)
                {
                    Spelare1Division1.Items.Add(p.fullname);
                    Spelare2Division1.Items.Add(p.fullname);
                }
                if (p.division == 2)
                {
                    Spelare1Division2.Items.Add(p.fullname);
                    Spelare2Division2.Items.Add(p.fullname);
                }
                if (p.division == 3)
                {
                    Spelare1Division3.Items.Add(p.fullname);
                    Spelare2Division3.Items.Add(p.fullname);
                }
                if (p.division == 4)
                {
                    Spelare1Division4.Items.Add(p.fullname);
                    Spelare2Division4.Items.Add(p.fullname);
                }
                if (p.division == 5)
                {
                    Spelare1Division5.Items.Add(p.fullname);
                    Spelare2Division5.Items.Add(p.fullname);
                }
                if (p.division == 6)
                {
                    Spelare1Division6.Items.Add(p.fullname);
                    Spelare2Division6.Items.Add(p.fullname);
                }
            }

        }

        private void Spelare1Division1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division1.SelectedItem != Spelare2Division1.SelectedItem && (Spelare1Division1.SelectedItem != null && Spelare2Division1.SelectedItem != null))
            {
                ReportDivision1.Enabled = true;
            }
            else
            {
                ReportDivision1.Enabled = false;
            }
        }

        private void Spelare2Division1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division1.SelectedItem != Spelare2Division1.SelectedItem && (Spelare1Division1.SelectedItem != null && Spelare2Division1.SelectedItem != null))
            {
                ReportDivision1.Enabled = true;
            }
            else
            {
                ReportDivision1.Enabled = false;
            }
        }

        private void Spelare1Division2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division2.SelectedItem != Spelare2Division2.SelectedItem && (Spelare1Division2.SelectedItem != null && Spelare2Division2.SelectedItem != null))
            {
                ReportDivision2.Enabled = false;
            }
            else
            {
                ReportDivision2.Enabled = false;
            }
        }

        private void Spelare2Division2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division2.SelectedItem != Spelare2Division2.SelectedItem && (Spelare1Division2.SelectedItem != null && Spelare2Division2.SelectedItem != null))
            {
                ReportDivision2.Enabled = false;
            }
            else
            {
                ReportDivision2.Enabled = false;
            }
        }

        private void Spelare1Division3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division3.SelectedItem != Spelare2Division3.SelectedItem && (Spelare1Division3.SelectedItem != null && Spelare2Division3.SelectedItem != null))
            {
                reportDivision3.Enabled = false;
            }
            else
            {
                reportDivision3.Enabled = false;
            }
        }

        private void Spelare2Division3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division3.SelectedItem != Spelare2Division3.SelectedItem && (Spelare1Division3.SelectedItem != null && Spelare2Division3.SelectedItem != null))
            {
                reportDivision3.Enabled = false;
            }
            else
            {
                reportDivision3.Enabled = false;
            }
        }

        private void Spelare1Division4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division4.SelectedItem != Spelare2Division4.SelectedItem && (Spelare1Division4.SelectedItem != null && Spelare2Division4.SelectedItem != null))
            {
                reportDivision4.Enabled = false;
            }
            else
            {
                reportDivision4.Enabled = false;
            }
        }

        private void Spelare2Division4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division4.SelectedItem != Spelare2Division4.SelectedItem && (Spelare1Division4.SelectedItem != null && Spelare2Division4.SelectedItem != null))
            {
                reportDivision4.Enabled = false;
            }
            else
            {
                reportDivision4.Enabled = false;
            }
        }

        private void Spelare1Division5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division5.SelectedItem != Spelare2Division5.SelectedItem && (Spelare1Division5.SelectedItem != null && Spelare2Division5.SelectedItem != null))
            {
                reportDivision5.Enabled = false;
            }
            else
            {
                reportDivision5.Enabled = false;
            }
        }

        private void Spelare2Division5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division5.SelectedItem != Spelare2Division5.SelectedItem && (Spelare1Division5.SelectedItem != null && Spelare2Division5.SelectedItem != null))
            {
                reportDivision5.Enabled = false;
            }
            else
            {
                reportDivision5.Enabled = false;
            }
        }

        private void Spelare1Division6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division6.SelectedItem != Spelare2Division6.SelectedItem && (Spelare1Division6.SelectedItem != null && Spelare2Division6.SelectedItem != null))
            {
                reportDivision6.Enabled = false;
            }
            else
            {
                reportDivision6.Enabled = false;
            }
        }

        private void Spelare2Division6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Spelare1Division6.SelectedItem != Spelare2Division6.SelectedItem && (Spelare1Division6.SelectedItem != null && Spelare2Division6.SelectedItem != null))
            {
                reportDivision6.Enabled = false;
            }
            else
            {
                reportDivision6.Enabled = false;
            }
        }

        private void ReportDivision1_Click(object sender, EventArgs e)
        {
            Spelare player1 = new Spelare();
            Spelare player2 = new Spelare();

            int player1ScoreGame1 = 0;
            int player2ScoreGame1 = 0;
            int player1ScoreGame2 = 0;
            int player2ScoreGame2 = 0;

            int game1DiffPlayer1 = 0;
            int game2diffPLayer1 = 0;
            int game1DiffPlayer2 = 0;
            int game2diffPLayer2 = 0;

            //Player 1
            foreach (var p in players)
            {
                if (p.fullname == Spelare1Division1.SelectedItem.ToString())
                {
                    player1 = p;
                    player1ScoreGame1 = Convert.ToInt32(ScorePlayer1Division1.Value);
                    player1ScoreGame2 = Convert.ToInt32(ScorePlayer1Division1Game2.Value);
                }
            }
            //Player 2
            foreach (var p in players)
            {
                if (p.fullname == Spelare2Division1.SelectedItem.ToString())
                {
                    player2 = p;
                    player2ScoreGame1 = Convert.ToInt32(ScorePlayer2Division1.Value);
                    player2ScoreGame2 = Convert.ToInt32(ScorePlayer2Division1Game2.Value);
                }
            }

            game1DiffPlayer1 = player1ScoreGame1 - player2ScoreGame1;
            game2diffPLayer1 = player1ScoreGame2 - player2ScoreGame2;
            game1DiffPlayer2 = player2ScoreGame1 - player1ScoreGame1;
            game2diffPLayer2 = player2ScoreGame2 - player1ScoreGame2;
            MessageBox.Show("Spelare 1: " + player1.fullname + " Spelare 2: " + player2.fullname);

            MessageBox.Show("Diff game 1 för spelare 1: " + game1DiffPlayer1 + "\r\n" +
                "Diff game 2 för spelare 1: " + game2diffPLayer1 + "\r\n" +
                "Diff game 1 för spelare 2: " + game1DiffPlayer2 + "\r\n" +
                "Diff game 2 för spelare 2: " + game2diffPLayer2);

            if ((game1DiffPlayer1 <= 11 && game1DiffPlayer1 >= 2) || (game2diffPLayer1 <= 11 && game2diffPLayer1 >= 2))
            {
                Game newGame = new Game();
                newGame.Rapport(player1, player2, player1ScoreGame1, player2ScoreGame1, player1ScoreGame2, player2ScoreGame2);
            }
            else
            {
                MessageBox.Show("Poängskillnaden måste vara mellan 2 och 11");
            }
        }
    }
}
