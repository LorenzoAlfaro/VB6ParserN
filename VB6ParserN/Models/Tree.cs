using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6ParserN.Models
{
    public class Tree
    {
        public int[] Trunk = new int[] { 0, 0 };

        public bool SplitGroup = false;

        public int GroupType; //0-IF,1 For i, 2 While, 3 Select, 4 Width

        public List<Tree> Branches;

        public Tree(int start, int end)
        {
            Branches = new List<Tree>();
            Trunk[0] = start;
            Trunk[1] = end;
        }
    }
}
