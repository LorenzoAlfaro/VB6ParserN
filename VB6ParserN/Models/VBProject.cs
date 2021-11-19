using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static VB6ParserN.VB6World;
using static StringParserN.CharWorld;

namespace VB6ParserN.Models
{
    public class VBProject
    {
        public string root;
        public List<VBSourceCode> Forms;
        public List<VBSourceCode> Modules;
        public List<VBSourceCode> Classes;
        public string Name;

        public VBProject()
        {
            Forms = new List<VBSourceCode>();
            Modules = new List<VBSourceCode>();
            Classes = new List<VBSourceCode>();
        }

        public VBProject(string path)
        {
            Forms = new List<VBSourceCode>();
            Modules = new List<VBSourceCode>();
            Classes = new List<VBSourceCode>();

            root = Path.GetDirectoryName(path);
            Name = Path.GetFileName(path);
            List<List<string>> myLists = new List<List<string>>();
            myLists = readVbProject(path);
            Forms = createSourceCodeList(myLists[0], root);
            Modules = createSourceCodeList(myLists[1], root);
            Classes = createSourceCodeList(myLists[2], root);
        }

        public static List<VBSourceCode> createSourceCodeList(List<string> paths, string root)
        {
            List<VBSourceCode> myForms = new List<VBSourceCode>();
            foreach (string path in paths)
            {
                StreamReader sr = new StreamReader(Path.Combine(root, path));
                VBSourceCode myForm = new VBSourceCode();
                myForm.VBName = vbNameModules(sr);
                myForm.RealName = path;

                sr.Close();

                myForms.Add(myForm);
            }

            return myForms;
        }
        public static Task<List<string>> findSentenceAsync(VBProject project, string searchString, string inLineDivisor, DrillBit bit, bool ignoreCasing = false)
        {
            return Task.FromResult(findSentence(project, searchString, inLineDivisor, bit, ignoreCasing));
        }
        public static List<string> findSentence(VBProject project, string searchString, string inLineDivisor, DrillBit bit, bool ignoreCasing = false)
        {
            //create an indexFile                       
            List<string> completePaths = new List<string>();
            foreach (VBSourceCode form in project.Forms)
            {
                completePaths.Add(form.RealName);
            }
            foreach (VBSourceCode mod in project.Modules)
            {
                completePaths.Add(mod.RealName);
            }
            foreach (VBSourceCode cls in project.Classes)
            {
                completePaths.Add(cls.RealName);
            }
            return searchMultipleFiles(project.root, completePaths, searchString, inLineDivisor, bit, ignoreCasing);
        }
        public static Task<List<List<string>>> readVBProjectAsync(string path)
        {
            return Task.FromResult(readVbProject(path));
        }
        public static List<string> searchMultipleFiles(string Root, List<string> paths, string searchString, string inLineDivisor, DrillBit bit, bool ignoreCasing = false)
        {
            List<string> results = new List<string>();
            StringComparison comp = ignoreCasing ? StringComparison.OrdinalIgnoreCase : StringComparison.CurrentCulture;
            foreach (string path in paths)
            {
                string[] lines = File.ReadAllLines(Path.Combine(Root, path));
                results.AddRange(lines
                    .Select((line, i) => bit(line, searchString, comp) ? i : -1)
                    .Where(i => i != -1)
                    .Select(index => path + inLineDivisor + index + inLineDivisor + lines[index]));
                //results.AddRange(lines.ToList().FindAll(item => bit(item, searchString)).Select(item => path + inLineDivisor +);
                //int[] v = lines.Select((b, i) => bit(b,searchString) ? i : -1).Where(i => i != -1).ToArray();
                //results.AddRange(Taladros.StreamReaderTaladro2(lines, searchString, path, inLineDivisor, ignoreCasing, bit));
            }
            return results;
        }
        public static List<List<string>> readVbProject(string path)
        {
            StreamReader stream = new StreamReader(path);

            int countForms = 0;
            int countModules = 0;
            int countClasses = 0;
            string line;
            List<List<string>> myLists = new List<List<string>>();
            List<string> forms = new List<string>();
            List<string> modules = new List<string>();
            List<string> classes = new List<string>();
            line = stream.ReadLine();
            while (line != null)
            {
                //write the line to console window
                Console.WriteLine(line);
                //Read the next line
                line = stream.ReadLine();
                if (line != null)
                {
                    if (line.StartsWith("Form"))
                    {
                        countForms += 1;
                        forms.Add(line.Split('=')[1]);
                    }
                    if (line.StartsWith("Module"))
                    {
                        countModules += 1;
                        string preSplit = line.Split('=')[1];
                        modules.Add(preSplit.Split(' ')[1]);
                    }
                    if (line.StartsWith("Class"))
                    {
                        string preSplit = line.Split('=')[1];
                        countClasses += 1;
                        classes.Add(preSplit.Split(' ')[1]);
                    }
                }
            }
            myLists.Add(forms);
            myLists.Add(modules);
            myLists.Add(classes);
            //close the file
            stream.Close();
            return myLists;
        }
        public static void scan(List<VBSourceCode> vbSourceCodeList, string root, List<VBSourceCode> searchingReferences)
        {
            List<string> directory = FormDirectory(searchingReferences); // a list of all the forms names
            foreach (VBSourceCode code in vbSourceCodeList)
            {
                code.ChildForms = findChildForms(Path.Combine(root, code.RealName), directory, code.VBName);
            }
        }
        private static List<childSourceCode> findChildForms(string path, List<string> directory, string formName) //change this to use TALADRO
        {
            //TODO: remove the .Contain and improve with char[] search code, too many false positives
            List<childSourceCode> childForms2 = new List<childSourceCode>();
            foreach (string childName in directory)
            {
                List<string> foundList = new List<string>();
                if (childName != formName)//dont add references to myself
                {
                    StreamReader sr = new StreamReader(path);
                    int lineNumber = 1;
                    string line = " ";
                    while (line != null)
                    {
                        line = sr.ReadLine();
                        if (line != null)
                        {
                            line = line + '\n';
                            string search1 = childName + " "; //might not be working with space Frm_Delivery_Models_To_Get.frm , Frm_Inv_Show
                            string search2 = childName + ".";
                            string search3 = childName + '\n';//more testing

                            if (line.Contains(search1) | line.Contains(search2) | line.Contains(search3))  //it is a bit too inclusev FRM_EMAIL and FRM_EMAIL_LIST are both counted, add a check for space at the end or '.'
                            {
                                if (!isItemPresent(childName, foundList))
                                {
                                    childSourceCode myChild = new childSourceCode();
                                    myChild.Name = childName;
                                    myChild.References.Add(lineNumber);
                                    childForms2.Add(myChild);//only add once
                                    foundList.Add(childName);
                                    //first appearance, create child and add line number
                                }
                                else
                                {
                                    childSourceCode myChild = findChildSimple(childName, childForms2);
                                    myChild.Name = childName;
                                    myChild.References.Add(lineNumber);
                                    //just add line number to the already created child
                                }
                            }
                        }
                        lineNumber += 1;
                    }
                    sr.Close();
                }
            }
            return childForms2;
        }
        private static List<string> FormDirectory(List<VBSourceCode> vbForms)
        {
            List<string> directory = new List<string>();
            foreach (VBSourceCode form in vbForms)
            {
                directory.Add(form.VBName); //use the the code name, not the file name
            }
            return directory;
        }
    }
}
