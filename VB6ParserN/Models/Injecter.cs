using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static TreeLogicN.TreeLogic;
using TreeLogicN;
using static VB6ParserN.VB6World;


namespace VB6ParserN.Models
{
    public static class Injecter
    {
        public static string inLineDivisor = "<%%%>"; //###

        public static void InjectFiles(string IgnoreInjectPath, Node myNody)
        {
            //openFileDialog1.ShowDialog();            
            List<string> Ignores = new List<string>();
            using (StreamReader sR = new StreamReader(IgnoreInjectPath))
            {
                while (!sR.EndOfStream) { Ignores.Add(sR.ReadLine()); }
            }

            foreach (Node FilePath in myNody.Branches)
            {
                string strFilePath = myNody.Trunk + "\\" + FilePath.Trunk;
                string[] lines = File.ReadAllLines(strFilePath);
                foreach (Node lineNumber in FilePath.Branches)
                {
                    if (!Checked(lineNumber.Branches[0].Branches[0].Trunk, Ignores))
                    {
                        lines[int.Parse(lineNumber.Trunk)] = injectLogging(lineNumber.Branches[0].Trunk,
                            int.Parse(lineNumber.Branches[0].Branches[0].Trunk));
                    }
                }
                using (StreamWriter sw = new StreamWriter(strFilePath))
                { foreach (string newLine in lines) { sw.WriteLine(newLine); } }
            }

        }
        public static bool Checked(string a, List<string> b)
        {
            foreach (string currentString in b)
            {
                if (currentString == a)
                {
                    return true;
                }
            }
            return false;
        }

        public static Node CreateIndexTree(string vb6ProjectPath, string IndexPath)
        {
            string[] lines = File.ReadAllLines(IndexPath);
            return createJSONTree(vb6ProjectPath, lines.ToList(), '*', false, inLineDivisor);
        }

        public static TreeNode LoadNodeTreeView(TreeView treeView, Node myNody)
        {
            return TranslateTree(myNody, new TreeNode());
            //treeView.Nodes.Add(treeViewNode);

            //return 0;
        }

        public static void cleanUpTree(TreeView treeView, TreeNode treeNode)
        {
            treeView.Nodes.Add(treeNode);
            treeView.Sort();
            treeView.Refresh();
            treeView.Update();
        }

        public static Task<TreeNode> LoadNodeTreeViewAsync(TreeView treeView, Node myNody)
        {
            return Task.FromResult(LoadNodeTreeView(treeView, myNody));
        }

        public static void addLogging2(string path, string inLineDivisor)
        {
            string[] lines = File.ReadAllLines(path); //all lines that need loggin
            int indexOfReferences = 1;
            foreach (string line in lines)
            {
                string pathToEdit = line.Split(new string[] { inLineDivisor }, StringSplitOptions.None)[0];
                int lineNumber = Int16.Parse(line.Split(new string[] { inLineDivisor }, StringSplitOptions.None)[1]);
                //string appendLine = createAppend(getLineCode(line));
                string injectedLine = injectLogging(line.Split(new string[] { inLineDivisor }, StringSplitOptions.None)[2], indexOfReferences);
                // Write the new file over the old file.
                string[] toEditlines = File.ReadAllLines(pathToEdit);
                StreamWriter sw = new StreamWriter(pathToEdit);
                using (sw)
                {
                    for (int currentLine = 0; currentLine < toEditlines.Length; ++currentLine)
                    {
                        if (currentLine == lineNumber)
                        {
                            sw.WriteLine(injectedLine); //add to the source code a command to display //inject middleMan
                        }
                        else
                        {
                            sw.WriteLine(toEditlines[currentLine]);
                        }
                    }
                }
                indexOfReferences += 1;
                //parse the info in line, code,linenumber, path, look for the file, open it, modify the line number and close
            }
        }
    }
}
