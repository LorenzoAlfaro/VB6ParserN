using System.Collections.Generic;
using System.Linq;
using static StringParserN.CharWorld;
using static VB6ParserN.VB6World;

namespace VB6ParserN.Models
{
    public class VBFunction
    {
        public string Name;
        public string Type;
        public int startLine;
        public int endLine;
        public int size;
        public string Privacy;

        public List<VBVariable> Arguments;

        public List<VBVariable> LocalVariables;

        public override string ToString()
        {
            return Name;
        }

        public VBFunction(string[] Lines, int start, int end)
        {
            LocalVariables = new List<VBVariable>();
            Arguments = new List<VBVariable>();
            startLine = start;
            endLine = end;
            size = end - start;
            Name =
                new string(getVBName(Lines[start].ToArray(), Lines[start].IndexOf('(')));
            Privacy = FunctionPrivacy(Lines[start]);
            List<int> indexes = returnVariables(Lines, start, end);
            foreach (int index in indexes)
            {
                VBVariable var = new VBVariable(Lines, index, start, end);
                LocalVariables.Add(var);
            }
            Arguments = ReturnArguments(Lines, start, end);
        }

        public VBFunction()
        {
            LocalVariables = new List<VBVariable>();
            Arguments = new List<VBVariable>();
        }

        public static List<VBVariable> ReturnArguments(string[] lines, int start, int end)
        {
            List<VBVariable> myVariables = new List<VBVariable>();
            List<string> arguments = getSubSections(lines[start].ToArray(), '(', ')', true, findClosing);

            if (arguments.Count != 0)
            {
                string Parenthesis = arguments[0];
                char[] charsToTrim = { '(', ')' };
                string TrimPar = Parenthesis.Trim(charsToTrim);

                if (TrimPar.Length > 0)
                {
                    string[] Arguments = TrimPar.Split(',');
                    foreach (string arg in Arguments)
                    {
                        string newArg = arg.Trim(' ');
                        VBVariable var = new VBVariable(lines, start, end, newArg);
                        myVariables.Add(var);
                    }
                }
            }
            return myVariables;
        }

    }
}
