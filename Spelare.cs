using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CST_With_only_excel
{
    public class Spelare
    {
        public string name;
        public string lastname;
        public string fullname;

        public int rank;
        public int division;
        public int placement;
        public int lastPlacement;

        public int gameWon;
        public int matchWon;
        public int pointDifference;

        public Spelare()
        {

        }
        public Spelare(string Fullname, int Rank)
        {
            fullname = Fullname;
           
            string[] tokens = Fullname.Split(new[] { "\n" }, StringSplitOptions.None);
            name = tokens[0];
            lastname = tokens[1];
            rank = Rank;

            if(rank >= 1 && rank <= 3)
            {
                division = 1;
            }
            else if(rank >= 4 && rank <= 6)
            {
                division = 2;
            }
            else if(rank >= 7 && rank <= 9)
            {
                division = 3;
            }
            else if(rank >= 10 && rank <= 12)
            {
                division = 4;
            }
            else if(rank >= 13 && rank <= 15)
            {
                division = 5;
            }
            else if(rank >= 16 && rank <= 18)
            {
                division = 6;
            }
        }

        public Spelare(string Name, string Lastname, int Rank, int Division)
        {
            name = Name;
            lastname = Lastname;
            rank = Rank;
            division = Division;

            fullname = name.Trim() + " " + lastname.Trim();
        }
    }
}
