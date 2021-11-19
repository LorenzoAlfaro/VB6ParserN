using System.Collections.Generic;

namespace VB6ParserN.Models
{
    public class childSourceCode
    {
        public string Name;
        public string Parent;
        public List<int> References;

        public childSourceCode()
        {
            References = new List<int>();
        }


        public override string ToString()
        {

            return this.Name;
        }
    }
}
