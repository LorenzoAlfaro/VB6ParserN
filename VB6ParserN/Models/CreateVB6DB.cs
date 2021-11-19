using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

using static VB6ParserN.Models.CreateSourceDB;
using static VB6ParserN.VB6World;

namespace VB6ParserN.Models
{
    public static class CreateVB6DB
    {
        // These are functions use for the SourceDB creation/initialization
        // A lot of VB6 source code parsing is done here!
        // Level 5 Loop/Loop/Loop/Loop/        
        public static void UpdateDB2(string path, SqlConnection conn)
        {
            string DirectoryPath = Path.GetDirectoryName(path);
            int ProjectId = CreateSourceDB.VBProject(DirectoryPath, conn);
            List<List<string>> sources;
            sources = VBProject.readVbProject(path); //Forms(0), Module(1), Class(2) paths

            //TODO: for the other TYPE tables, get the string Type and look it up
            //if it exists use the ID, if not, insert the type and get the new ID
            //write a function for that
            foreach (string sourcePath in sources[0])
            {
                CreateSource(ProjectId, sourcePath, DirectoryPath, 1, conn);
            }
            foreach (string sourcePath in sources[1])
            {
                CreateSource(ProjectId, sourcePath, DirectoryPath, 2, conn);
            }
            foreach (string sourcePath in sources[2])
            {
                CreateSource(ProjectId, sourcePath, DirectoryPath, 3, conn);
            }
        }
        public static void CreateSource(int ProjectId, string sourcePath, string DirectoryPath, int sourceType, SqlConnection conn)
        {
            int SourceId = SourceFile(sourcePath, ProjectId, "", sourceType, 1, conn); //Update Sources Table             
            string[] Lines = File.ReadAllLines(Path.Combine(DirectoryPath, sourcePath));

            List<string> SearchWords = new List<string>();
            SearchWords.Add("Sub ");
            SearchWords.Add("Private Sub");
            SearchWords.Add("Public Sub");

            List<int[]> Functions = getFunctions(Lines, SearchWords, "End Sub"); //creates Subs
            foreach (int[] function in Functions)
            {
                CreateFunction(SourceId, Lines, function, 0, conn);//subs
            }

            SearchWords.Clear();
            SearchWords.Add("Function ");
            SearchWords.Add("Private Function");
            SearchWords.Add("Public Function");

            Functions = getFunctions(Lines, SearchWords, "End Function"); //creates Function
            foreach (int[] function in Functions)
            {
                CreateFunction(SourceId, Lines, function, 1, conn); //functions
            }
        }
        public static void CreateFunction(int SourceId, string[] Lines, int[] function, int SubOrFunction, SqlConnection conn)
        {
            string FunctionName =
                new string(getVBName(
                    Lines[function[0]].ToCharArray(),
                    Lines[function[0]].IndexOf('(')));
            //StringFilter.Parser.FunctionName(Lines[function[0]],'(');

            int ReturnType = ReturnVarType(FunctionReturnType(Lines[function[0]]), conn);
            int PrivacyType = ReturnPrivacyType(FunctionPrivacy(Lines[function[0]]), conn);

            int FunctionId = Function(FunctionName, SourceId, ReturnType, PrivacyType, SubOrFunction, 1, function[0], function[1], conn);//Update Function Table DB

            List<int> Variables = returnVariables(Lines, function[0], function[1]);
            foreach (int Variable in Variables) //Create Variables
            {
                if (Lines[Variable].Contains(","))
                {
                    Console.WriteLine(Lines[Variable]);
                }

                if (Lines[Variable].Contains(":"))
                {
                    Console.WriteLine(Lines[Variable]);
                }

                List<string[]> subVariables = getSubVariables(Lines[Variable]);
                foreach (string[] subVariable in subVariables)
                {
                    int varType = ReturnVarType(subVariable[1], conn);
                    CreateVariable(FunctionId, Lines, subVariable[0], function, 0, varType, conn);
                }
            }

            List<string[]> Arguments = ReturnArguments2(Lines[function[0]].ToCharArray());
            foreach (string[] Argument in Arguments) //Create Arguments
            {
                int ArgType = ReturnVarType(Argument[1], conn);
                CreateVariable(FunctionId, Lines, Argument[0], function, 0, ArgType, conn);
            }

        }
        public static void CreateVariable(int FunctionId, string[] Lines, string VariableName, int[] function, int ArgumentOrVariable, int VarType, SqlConnection conn)
        {
            int VariableId = Variable(VariableName, FunctionId, VarType, ArgumentOrVariable, 1, conn); //Update Variables Table DB            
            List<string> Instances = VBVariable.getInstances(Lines, function[0], function[1], VariableName);
            foreach (string instance in Instances)
            {
                CreateInstance(VariableId, instance, conn);
            }
        }
        public static void CreateInstance(int VariableId, string instance, SqlConnection conn)
        {
            string[] array = instance.Split('-');
            int LineNumber = Int16.Parse(array[0]);
            int ColumnNumber = Int16.Parse(array[1]);
            int InstanceId = Instance(VariableId, LineNumber, ColumnNumber, conn);//Update Instances Table DB
        }
        public static int ReturnVarType(string type, SqlConnection conn)
        {
            int typeId = GetReturnTypeId("ReturnType_Table", "ReturnType", type, conn);
            if (typeId == -1) //did not found it
            {
                typeId = InsertNewType("ReturnType_Table", "ReturnType", type, conn);//insert new record
            }
            return typeId;
        }
        public static int ReturnPrivacyType(string type, SqlConnection conn)
        {
            int typeId = GetReturnTypeId("PrivacyType_Table", "PrivacyType", type, conn);
            if (typeId == -1) //did not found it
            {
                typeId = InsertNewType("PrivacyType_Table", "PrivacyType", type, conn);//insert new record
            }
            return typeId;
        }
        public static int Privacy(string type)
        {
            return 0;
        }
    }
}
