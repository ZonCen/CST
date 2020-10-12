using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CST_With_only_excel
{
    public class Division
    {
        public List<Spelare> players = new List<Spelare>();
        public List<Spelare> division1 = new List<Spelare>();
        public List<Spelare> division2 = new List<Spelare>();
        public List<Spelare> division3 = new List<Spelare>();
        public List<Spelare> division4 = new List<Spelare>();
        public List<Spelare> division5 = new List<Spelare>();
        public List<Spelare> division6 = new List<Spelare>();
        public List<RankInfo> oldRankinfo = new List<RankInfo>();

        public Division(List<Spelare> list, List<RankInfo> OldRankInfo)
        {
            players.AddRange(list);
            oldRankinfo = OldRankInfo;
            foreach (var p in players)
            {
                if (p.division == 1)
                {
                    division1.Add(p);
                }
                if (p.division == 2)
                {
                    division2.Add(p);
                }
                if (p.division == 3)
                {
                    division3.Add(p);
                }
                if (p.division == 4)
                {
                    division4.Add(p);
                }
                if (p.division == 5)
                {
                    division5.Add(p);
                }
                if (p.division == 6)
                {
                    division6.Add(p);
                }
            }
            sortRanking();
        }

        public void sortRanking()
        {
            division1 = division1.OrderBy(x => x.rank).ToList();
            division2 = division2.OrderBy(x => x.rank).ToList();
            division3 = division3.OrderBy(x => x.rank).ToList();
            division4 = division4.OrderBy(x => x.rank).ToList();
            division5 = division5.OrderBy(x => x.rank).ToList();
            division6 = division6.OrderBy(x => x.rank).ToList();
        }

        public void newRanks(List<Spelare> division)
        {
            /* Ifall spelare får rank 1 i sin division (Med undantag division1) får den spelare -1 i rank, 
             * Ifall spelare får rank 3 i sin division (Med undantag sista icke tomma divisionen) får den spelare +1 i rank
             * Gör om nedan kod för att räkna ut vilka som får plats 1 och 3.
             *             if (player1.rank > 3)
                rankSpelare1 = player1.rank - (3 * (player1.division -1));
             */
            foreach (var p in division)
            {
                int rank = p.rank;
                int oldRank = p.rank;
                if (rank > 3)
                {
                    rank = p.rank - (3 * (p.division - 1));
                }

                if (p.rank != 1)
                {
                    if (rank == 1)
                    {
                        p.rank--;
                    }
                    else if (rank == 3)
                    {
                        p.rank++;
                    }
                }
            }
        }

        public void calculateNewRanks()
        {
            division1 = division1.OrderByDescending(x => x.matchWon).
                      ThenByDescending(y => y.gameWon).
                      ThenByDescending(z => z.pointDifference).ToList();
            division2 = division2.OrderByDescending(x => x.matchWon).
                    ThenByDescending(y => y.gameWon)
                    .ThenByDescending(z => z.pointDifference).ToList();
            division3 = division3.OrderByDescending(x => x.matchWon).
                ThenByDescending(y => y.gameWon)
                .ThenByDescending(z => z.pointDifference).ToList();
            division4 = division4.OrderByDescending(x => x.matchWon).
                    ThenByDescending(y => y.gameWon)
                    .ThenByDescending(z => z.matchWon).ToList();
            division5 = division5.OrderByDescending(x => x.matchWon).
                ThenByDescending(y => y.gameWon)
                .ThenByDescending(z => z.pointDifference).ToList();

            division6 = division6.OrderByDescending(x => x.matchWon).
                            ThenByDescending(y => y.gameWon)
                            .ThenByDescending(z => z.pointDifference).ToList();

            newRanks(division1);
            newRanks(division2);
            newRanks(division3);
            newRanks(division4);
            newRanks(division5);
            newRanks(division6);

            players = players.OrderBy(x => x.rank).ToList();

            UpdateRankingTable(players);
        }

        public void UpdateRankingTable(List<Spelare> list)
        {
            ExcelHandler updateRankTable = new ExcelHandler();
            updateRankTable.updateRank(list, 6, oldRankinfo);
        }
    }
}
