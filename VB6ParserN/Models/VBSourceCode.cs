using System.Collections.Generic;
using static VB6ParserN.VB6World;

namespace VB6ParserN.Models
{
    public class VBSourceCode
    {
        public string RealName; //the name used in the file
        public string VBName; //The name used in the code;

        public List<childSourceCode> ChildForms; // The forms mentioned in my code
        public List<childSourceCode> ChildModules; // The forms mentioned in my code
        public List<childSourceCode> ChildClasses; // The forms mentioned in my code

        public List<VBFunction> Functions;
        public List<VBFunction> Subs;

        public override string ToString()
        {
            return RealName;
        }
        public string ToString(bool fileName)
        {
            if (fileName)
            {
                return RealName;
            }
            return VBName;
        }
        public VBSourceCode()
        {
            ChildForms = new List<childSourceCode>();
            ChildModules = new List<childSourceCode>();
            ChildClasses = new List<childSourceCode>();
        }
        public VBSourceCode(string[] Lines, string fileName)
        {
            RealName = fileName;
            ChildForms = new List<childSourceCode>();
            ChildModules = new List<childSourceCode>();
            ChildClasses = new List<childSourceCode>();
            Functions = new List<VBFunction>();
            Subs = new List<VBFunction>();

            List<string> SearchWords = new List<string>();
            SearchWords.Add("Function ");
            SearchWords.Add("Private Function");
            SearchWords.Add("Public Function");

            List<int[]> subsIndexes = getFunctions(Lines, SearchWords, "End Function");
            foreach (int[] index in subsIndexes)
            {
                Functions.Add(new VBFunction(Lines, index[0], index[1]));
            }
            SearchWords.Clear();
            SearchWords.Add("Sub ");
            SearchWords.Add("Private Sub");
            SearchWords.Add("Public Sub");
            subsIndexes = getFunctions(Lines, SearchWords, "End Sub");
            foreach (int[] index in subsIndexes)
            {
                Subs.Add(new VBFunction(Lines, index[0], index[1]));
            }
        }
    }
}
