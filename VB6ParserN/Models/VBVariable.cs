using System.Collections.Generic;
using System.Linq;
using static StringParserN.CharWorld;
using static VB6ParserN.VB6World;

namespace VB6ParserN.Models
{
    public class VBVariable
    {
        public string Name;
        public string Type;
        public List<string> Instances;
        public int Line;

        public override string ToString()
        {
            return Name;
        }

        public VBVariable(string[] Lines, int index, int start, int end)
        {
            Name = VariableName(Lines[index]);
            Type = VariableType(Lines[index]);
            Line = index;
            Instances = getInstances(Lines, start, end, Name);
        }

        public VBVariable(string[] Lines, int start, int end, string preSplit)
        {
            Name = ArgumentName(preSplit);
            Type = ArgumentType(preSplit);
            Line = start;
            Instances = getInstances(Lines, start, end, Name);
        }

        public VBVariable()
        {

        }


        private static List<int> getPointers(string line, string name, int offset) //recursive
        {
            List<int> pointers = new List<int>();
            char[] Array = line.ToArray();
            int index = line.IndexOf(name);
            if (index == -1)
            {
                return pointers;
            }
            else
            {
                pointers.Add(index + offset);
                string NewLine = new string(getSubArray(index + name.Length, line.Length - 1, Array));
                pointers.AddRange(getPointers(NewLine, name, index + name.Length + offset));
                return pointers;
            }
        }

        private static List<int> getValidPointers(List<int> pointers, char[] array, string name)
        {
            List<int> validPointers = new List<int>();
            foreach (int pointer in pointers)
            {
                if (ValidPointerVariable(pointer, array, name))
                {
                    validPointers.Add(pointer);
                }
            }
            return validPointers;
        }

        public static List<string> getInstances(string[] lines, int start, int end, string name)
        {
            List<string> indexes = filterInstances(lines, getVariables(lines, start, end, name), name);
            return indexes;
        }

        private static List<string> filterInstances(string[] lines, List<int> hits, string name)
        {
            List<string> indexes = new List<string>();
            foreach (int index in hits)
            {
                string line = lines[index];
                List<int> pointers = getValidPointers(getPointers(line, name, 0), line.ToArray(), name);
                if (pointers.Count != 0)
                {
                    List<int> finalList = new List<int>();
                    finalList.Add(index);
                    foreach (int pointer in pointers)
                    {
                        indexes.Add(index.ToString() + "-" + pointer.ToString());
                    }
                }
            }
            return indexes;
        }

    }
}
