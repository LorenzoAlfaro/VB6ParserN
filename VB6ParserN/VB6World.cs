using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VB6ParserN.Models;
using static StringParserN.CharWorld;

namespace VB6ParserN
{
    public static class VB6World
    {
        // comments in VB6 start with '
        public static int startComment(char[] array, int index, int commentChar)
        {
            for (int i = index; i < array.Length; i++) //TODO: change type int to char, also this function is the same as stringparser ->findNextChar
            {
                if (array[i] == commentChar)//check the ' is not part of a string "dfdf'"
                {
                    return i;
                }
            }
            return -1;
        }
        public static List<int[]> offZones(char[] array)
        {
            List<int[]> stringZones = getSubSectionsIndexes(array, '"', '"', endString);
            int[] commentZones = { startComment(array, 0, '\''), array.Length };
            if (commentZones[0] != -1)
            {
                stringZones.Add(commentZones);
            }
            return stringZones;
        }
        public static bool SafeZone(int pointer, char[] array)
        {
            bool flag = true;
            List<int[]> redZone = offZones(array);
            foreach (int[] zone in redZone)
            {
                if (pointer < zone[0] & pointer > zone[1])
                {
                    flag = true;
                }
                else
                {
                    return false;
                }
            }

            return flag;
        }
        public static bool ValidPointerVariable(int pointer, char[] array, string name)
        {
            if (checkEnding(pointer, array, name) & checkStart(pointer, array) & SafeZone(pointer, array))
            {
                return true;
            }
            return false;
        }
        // TODO: Factorize out the if logic
        public static bool checkEnding(int pointer, char[] array, string name)
        {
            char[] AllowedEndings = { '&', '*', '>', '<', '%', '=', '/', '-', '+', ':', ' ', '\t', '.', '!', ')', '(', '[', ']', ',', '\n' };
            int lastChar = pointer + name.Length;

            if (lastChar < array.Length)
            {
                return CharFound(array[lastChar], AllowedEndings);// the character after the variable name
            }
            else
            {
                return true;
            }
        }
        public static bool checkEndingLine(int pointer, char[] array, string name)
        {
            char[] AllowedEndings = { '\n' };
            int lastChar = pointer + name.Length;

            if (lastChar < array.Length)
            {
                return CharFound(array[lastChar], AllowedEndings);
            }
            else
            {
                return true;
            }
        }
        public static bool checkStart(int pointer, char[] array)
        {
            char[] AllowedStarts = { '&', '*', '>', '<', '%', '=', '/', '-', '+', ':', ' ', '\t', '(', '[', ',' };
            int firstChar = pointer - 1;
            if (pointer > 0)
            {
                return CharFound(array[firstChar], AllowedStarts);
            }
            else
            {
                return true;
            }
        }
        public static bool checkLabel(char myChar)//just passed index -1
        {
            char[] AllowedStarts = { ' ', '\t', '.', '!', ')', '(', '[', ']', ',', '\n' };

            return CharFound(myChar, AllowedStarts);
        }

        public static bool ValidLabel(char[] array, int index)
        {
            //index = get the line.IndexOf(":") -1
            for (int i = index; i >= 0; i--)
            {
                if (checkLabel(array[i]))
                {
                    return false;
                }
            }
            return true;
        }

        public static int beginVariableName(char[] array, int index, int LineNumber = 0)
        {   //I'm expected to be call when you found the index of the beginning of a method Rs.Open and you want to find
            //the start of the Object Name, that is why I go backwards until I find a space. I'm expecting a simple object
            // not Records.Rs["hola " + "saludos"].Open will break for the space            
            //while (array[index] != ' ')            
            for (int i = index; i < array.Length; i--)
            {
                if (checkStart(i, array))
                {
                    return i;
                }
            }
            return -1;
        }

        public static int endParameter(char[] array, int index, int LineNumber)
        {
            //Context: there is a Sub and I want to find the end of its first parameter

            //I'm expecting my Index to start after a parameter is expected Rs.Open   ^"SQL ~~" ,myAdoDB,~~~ or Rs.Execute(^"sql")
            //and will find the index where that parameter ends
            //Rs.Open   "SQL ~~"^ ,myAdoDB,~~~ 
            //Rs.Execute("sql"^)
            //cant call a sub within a sub 

            if (array[index] == ' ')
            {
                index = skipSpaces(array, index, ' ');
            }
            else if (array[index] == '"')
            {
                index = endString(array, index + 1, '"', '"');//forward the index
            }
            else if (array[index] == '(')
            {
                index = findClosing(array, index, '(', ')');    //weird but possible rs.open ("Select")            
            }
            else
            {
                index = findEnd(array, index, LineNumber) + 1; //a variable, function or and array or an operator & + |
            }
            //check the forwarded index
            if (index == array.Length | index == -1)
            {
                return -1;
            }
            else if (array[index] == ',' | array[index] == '\'' | index == array.Length - 1)
            {
                return index;
            }
            else
            {
                return endParameter(array, index, LineNumber);
            }
        }
        public static int findEnd(char[] array, int index, int lineNumber) //TODO use TALADRO
        {
            //I'm expecting to be called when you found a variable or function or and array or a operator
            int offSet = index;
            for (int i = index; i < array.Length; i++)
            {
                if (array[i] == ' ' | array[i] == '(' | array[i] == '[' | array[i] == ',' | array[i] == '\'')
                {
                    offSet = i;
                }
            }

            if (array[offSet] == ' ' | array[offSet] == ',' | array[offSet] == '\'')
            {
                return index; //could be a variable name, or an operator
            }
            else if (array[offSet] == '(')
            {
                return findClosing(array, index, '(', ')'); //return end of a function
            }
            else //if (array[currentIndex] == '[')
            {
                return findClosing(array, index, '[', ']');  //return end of an array[1]
            }

        }

        public static char[] getVBName(char[] array, int index) //index of starting (
        {
            if (index != -1)
            {
                return getSubArray(beginVariableName(array, index), index - 1, array);
            }
            return new char[] { };
        }
        public static char[] returnComments(char[] array, int index)
        {
            if (index != -1)
            {
                return getSubArray(index, array.Length - 1, array);
            }
            return new char[] { };
        }
        //inject methods
        public static string AddParameter(string code, string originalMethod, int LineNumber = 0)
        {
            //Assuming he doesnt do TWO ReadTable Values in the same line
            char[] Search = code.ToArray();
            int injectPoint = code.IndexOf(originalMethod) + originalMethod.Length; //"ReadTableValue(^"
            string Cara = new string(getSubArray(0, injectPoint, Search));
            string Cola = new string(getSubArray(injectPoint, Search.Count(), Search)); //I don't know why end is not Search.Count()-1 in this case 

            if (Cola.IndexOf(originalMethod) > 0)
            {
                Cola = AddParameter(Cola, originalMethod, LineNumber); //multiple instances of the same function in the same line
            }
            return Cara + LineNumber.ToString() + ", " + Cola;
        }
        public static string EncloseVariable(string code, string originalMethod, string EnclosingFunction, int LineNumber)
        {
            char[] Search = code.ToArray();

            //find the variable if any Rs.Update or .Update
            int end = code.IndexOf(originalMethod); //".Update"
            int start = beginVariableName(Search, end, LineNumber);
            char[] VariableName = getSubArray(start, end - 1, Search); //ese plus 1 might not work anymore
            string Cara = new string(getSubArray(0, start, Search));
            string Cola = new string(getSubArray(end + 7, Search.Count() - 1, Search)); //-1? removed the -1,


            if (VariableName.Count() != 0)
            {
                string VarName = new string(VariableName);
                return Cara + " " + VarName + ".Update: Call " + EnclosingFunction + LineNumber + "," + VarName + "," + VarName + ".Bookmark" + ") " + Cola;
                //return Cara + " Call " + EnclosingFunction + LineNumber + "," + VarName + ") " + Cola;
            }
            else // With case
            {
                //Console.WriteLine(code);
                //turn Cara + "Set genericRS = .Clone" + ": " + "Call " + EnclosingFunction + LineNumber + "," + "genericRS" + ")" + Cola;// this can be simplfied
                return Cara + " .Update: Call " + EnclosingFunction + LineNumber + "," + ".Clone" + ", .Bookmark" + ")" + Cola;// this can be simplfied
            }
        }
        public static string wrapParameter(string code, string originalMethod, string injectedFunction, int LineNumber)
        {
            int start;
            int end;
            //find the parameter
            start = code.IndexOf(originalMethod) + originalMethod.Length; // index + 5 char for .open
            end = endParameter(code.ToArray(), start, LineNumber);
            if (end == -1)// corner case Rs.Open cmd^
            {
                end = code.ToArray().Length;
            }
            return MergeParts(start, end, code.ToArray(), injectedFunction, LineNumber.ToString());
        }
        public static string MergeParts(int start, int end, char[] array, string InjectFunction, string injectParameter)
        {
            //if no parameter return same line
            string str = new string(getSubArray(start, end, array));
            if (str == " ")
            {
                return new string(array);
            }

            string cola = "";
            if (end < array.Length)
            {
                cola = new string(getSubArray(end, array.Length - 1, array));
            }
            string cara = new string(getSubArray(0, start, array));

            return cara + " " + InjectFunction + injectParameter + "," + str + ")" + cola;
        }

        public static string VariableName(string line)
        {
            //Dim MyCnt As Control
            //Dim TmpStr(1000, 100) As String
            //string[] parts = trimLine.Split(new string[] { "Dim " }, StringSplitOptions.None);                      
            string[] parts2 = line.Split(new string[] { " As " }, StringSplitOptions.None);
            if (parts2.Length > 1)
            {
                string[] parts3 = parts2[0].Split(new string[] { "(" }, StringSplitOptions.None);
                return parts3[0].Trim();
            }
            else if (parts2.Length == 1)
            {
                string[] parts4 = parts2[0].Split(new string[] { "(" }, StringSplitOptions.None);
                return parts4[0].Trim();
            }
            else
            {
                return "Undefined";
            }

        }
        public static string VariableType(string line)
        {
            //Dim MyCnt As Control
            //Dim TmpStr(1000, 100) As String
            //Dim Frm As New Frm_Utility_Mass_Changes_Model_Number //TODO
            //Dim crAppl As New CRAXDRT.Application
            //Dim RptEngine As New RptEngine
            //Dim ReturnVal As Long, PathBuffSize As Long, PathBuff As String
            //Dim j As Integer, s As String: s = ""            

            //string[] parts = trimLine.Split(new string[] { "Dim " }, StringSplitOptions.None);
            string[] parts2 = line.Split(new string[] { " As " }, StringSplitOptions.None);

            if (parts2.Length > 1)
            {
                string[] parts3 = parts2[1].Split(' '); //TODO: Add logic for //Dim Frm As New Frm_Utility_Mass_Changes_Model_Number

                if (parts3[0] == "New")
                {
                    return parts3[1].Trim();
                }

                return parts3[0].Trim();
            }
            else
            {
                return "Undefined";
            }
        }
        public static string FunctionReturnType(string line)
        {
            List<string> arguments = getSubSections(line.ToArray(), '(', ')', true, findClosing);

            if (arguments.Count > 0)
            {
                string[] parts = line.Split(new string[] { arguments[0] }, StringSplitOptions.None); //function(line as string) As bool ' comment //Private Sub CmdGetLatLngZone_Click()  ' // Get Lat, Lng, Zone button on Delivery Zones form
                string TrimSpace = parts[1].Trim();
                string[] parts2 = TrimSpace.Split(' ');

                if (parts2.Count() > 1)
                {
                    if (parts2[0] == "As")
                    {
                        return parts2[1];
                    }
                }
                else
                {
                    return "Undefined";
                }
            }
            else
            {
                return "Undefined";
            }
            return "Undefined";
        }
        public static string FunctionPrivacy(string line)
        {
            if (line.StartsWith("Private"))
            {
                return "Private";
            }
            else if (line.StartsWith("Public"))
            {
                return "Public";
            }
            else
            {
                return "Undefined";
            }
        }
        public static string ArgumentType(string line)
        {
            //Dim FirstSpace As Integer, Slash As Integer TODO: two variable definitions in ONE LINE
            //Dim MyCnt As Control
            //Dim TmpStr(1000, 100) As String
            //Spaces have been Trimmed passing to this function
            //Optional MyCnt As Control //break case, Optional ByVal VariableName As Variant

            //MyCnt, returnUndefined
            //MyCnt As Control, return2
            //Optional ShowComd = vbNormalNoFocus returnUndefined
            //Optional ShowComd As Long = vbNormalNoFocus return3
            //Optional ByVal WorkDir As Variant  return4
            //Optional ByVal WorkDir As Variant = fakeValue return4
            //Optional ByVal WorkDir = fakeValue returnUndefined
            //ShowComd = vbNormalNoFocus returnUndefined
            //ShowComd As Long = vbNormalNoFocus return2
            //ByVal WorkDir As Variant  return3
            //ByVal WorkDir As Variant = fakeValue return3
            //ByVal WorkDir = fakeValue returnUndefined

            //keywords Optional,ByRef,ByVal,As, =
            //in that order

            string[] parts = line.Split(new string[] { " As " }, StringSplitOptions.None);
            if (parts.Length > 1)
            {
                //remove other stuff
                string[] parts2 = parts[1].Split(' ');
                return parts2[0].Trim();
            }
            else
            {
                return "Undefined";
            }
        }
        public static string ArgumentName(string line)
        {
            //MyCnt As Control
            string[] parts = line.Split(' ');

            int index = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] != "Optional" & parts[i] != "ByRef" & parts[i] != "ByVal" & parts[i] != " ")
                {
                    index = i;
                    break;
                }
            }
            return parts[index];
        }
        public static string[] getRecordSetVariable(string line)
        {
            string[] Names = new string[] { "", "" };

            int index = line.IndexOf("![");
            int start = beginVariableName(line.ToArray(), index, 0);
            string name = new string(getSubArray(start, index, line.ToArray()));
            Names[0] = name;
            Names[1] = getSubSections(line.ToArray(), '[', ']', true, findClosing)[0];

            return Names;
        }
        public static string injectLogging(string code, int LineNumber) //TODO: Improved the search, dont rely on .Contains
        {
            string functionPart1 = "QueryLoggin(";
            string UpdatefunctionPart1 = "UpdateLoggin(";

            string FinalString = code;

            if (code.Contains(".Open ") & !code.Contains("ExecuteScalar") & !code.Contains(".OpenDocument"))
            {
                FinalString = wrapParameter(code, ".Open", functionPart1, LineNumber);
            }

            if (code.Contains(".Execute "))
            {
                FinalString = wrapParameter(FinalString, ".Execute", functionPart1, LineNumber);
            }

            if (code.Contains(".Source = ") & !code.Contains(".Source = Cmd"))
            {
                FinalString = wrapParameter(FinalString, ".Source =", functionPart1, LineNumber);
            }

            if (code.Contains(".CommandText ="))
            {
                FinalString = wrapParameter(FinalString, ".CommandText =", functionPart1, LineNumber);
            }

            if (code.Contains(".Update") & !code.Contains("'") & !code.Contains(".UpdatePrice") & !code.Contains(".Configuration.Fields.Update") & !code.Contains(".UpdateTextCarriers") & !code.Contains(".UpdateExpenseHeader"))
            {

                FinalString = EncloseVariable(FinalString, ".Update", UpdatefunctionPart1, LineNumber); //Can't use .Update with a trailing space because some lines END with the .Update
            }

            if (code.Contains("ReadTableValue("))
            {
                FinalString = AddParameter(FinalString, "ReadTableValue(", LineNumber);
            }

            return FinalString; //nothing to change
        }
        //
        public static List<string> VariableNames3(string line, int LineNumber, string inLineDivisor)
        {
            List<string> myVariables = new List<string>();
            string code = line.Split(new string[] { inLineDivisor }, StringSplitOptions.None)[0];
            char[] firstParameter;
            if (code.Contains(".Open"))
            {
                firstParameter = code.Split(new string[] { ".Open" }, StringSplitOptions.None)[1].ToCharArray();
            }
            else
            {
                firstParameter = code.Split(new string[] { ".Execute" }, StringSplitOptions.None)[1].ToCharArray();
            }

            for (int i = 0; i < firstParameter.Length; i++)
            {
                if (firstParameter[i] == '"')//skip all the characters in the string
                {
                    i = endString(firstParameter, i, '"', '"');
                }
                else if (firstParameter[i] == ',')
                {
                    break; //end of the first parameter
                }
                else if (firstParameter[i] == '&')
                {
                    i += 1;
                    i = skipSpaces(firstParameter, i, ' ');

                    if (firstParameter[i] == '"')
                    {
                        i = endString(firstParameter, i, '"', '"');
                    }
                    else
                    {
                        int start = i; //find end

                        i = findEnd(firstParameter, i, LineNumber);
                        int end = i;
                        char[] var = getSubArray(start, end, firstParameter);
                        string str = new string(var);
                        myVariables.Add(str);
                    }
                }
                else if (firstParameter[i] == '\'')
                {
                    break; //found the start of a comment
                }
            }

            List<string> cleanVariables = new List<string>();
            foreach (string item in myVariables)
            {
                string cleanItem;
                if (item.StartsWith("Mod_SingleQuoteSQLHandler("))
                {

                    cleanItem = item.Replace("Mod_SingleQuoteSQLHandler(", "");
                    cleanItem = cleanItem.Remove(cleanItem.Length - 1, 1);
                }
                else
                {
                    cleanItem = item;
                }
                cleanVariables.Add(cleanItem);
            }
            return cleanVariables;
        }
        public static List<string[]> getSubVariables(string line)
        {
            //Dim lProcesses(1000, 100) As Long, lModules(1000, 100) As Long, n As Long, lRet As Long, hProcess As Long: 'hey more stuff!
            //Dim j As Integer, s As String: s = ""
            List<string[]> myVariables = new List<string[]>();

            //first split Dim
            //second split : 
            //third split ,
            //fourth split As
            //fith split ' '
            //sixth check New

            string[] Parts1 = line.Split(new string[] { "Dim " }, StringSplitOptions.None);
            string[] Parts2 = Parts1[1].Split(new string[] { ":" }, StringSplitOptions.None);

            List<string> arguments = getSubSections(Parts2[0].ToArray(), '(', ')', true, findClosing);

            if (arguments.Count != 0)
            {
                foreach (string argument in arguments)
                {
                    Parts2[0] = Parts2[0].Replace(argument, ""); //eliminate the lProcesses(1000, 100) -> lProcesses
                }
            }

            string[] Parts3 = Parts2[0].Split(new string[] { "," }, StringSplitOptions.None); ////Dim TmpStr(1000, 100) As String

            foreach (string subVar in Parts3)
            {
                string newArg = subVar.Trim(' ');
                string[] Argument = new string[] { VariableName(newArg), VariableType(newArg) };

                myVariables.Add(Argument);
            }
            return myVariables;
        }
        //
        public static List<string[]> ReturnArguments2(char[] line)
        {
            List<string[]> myVariables = new List<string[]>();
            List<string> arguments = getSubSections(line, '(', ')', true, findClosing);

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
                        string[] Argument = new string[] { ArgumentName(newArg), ArgumentType(newArg) };

                        myVariables.Add(Argument);
                    }
                }
            }
            return myVariables;
        }

        ///////////////////////
        ///



        public static bool bitFieldMapping(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.Length == 0)
            {
                return false;
            }

            int index = LineToCheck.IndexOf("] =", comparison);
            if (index != -1)
            {
                if (SafeZone(index, LineToCheck.ToArray())) //check this is not inside a string or comment
                {
                    List<string> fields = getSubSections(LineToCheck.ToArray(), '[', ']', true, findClosing);
                    if (fields.Count == 2)
                    {
                        if (fields[0] == fields[1])
                        {
                            return false; //equal field names
                        }
                    }
                    else//many fields [][][]
                    {
                        return false;
                    }
                    int count = LineToCheck.Count(c => c == '!');
                    if (count > 1) //at leat two rs![],rs1![]
                    {
                        return true;
                    }
                    return false;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool bitFunctionStart(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.StartsWith("Public Function")
                    | LineToCheck.StartsWith("Private Function")
                    | LineToCheck.StartsWith("Function ")
                    | LineToCheck.StartsWith("Public Sub")
                    | LineToCheck.StartsWith("Private Sub")
                    | LineToCheck.StartsWith("Sub "))
            {
                return true;
            }
            return false;
        }
        public static bool bitSafeZone(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.Contains(searchWord, comparison))
            {
                int withIndex = LineToCheck.IndexOf(searchWord, comparison);

                if (SafeZone(withIndex, LineToCheck.ToArray()))
                {
                    return true;
                }
            }
            return false;
        }
        public static bool bitValidPointer(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.Contains(searchWord, comparison))
            {
                int withIndex = LineToCheck.IndexOf(searchWord, comparison);
                if (ValidPointerVariable(withIndex, LineToCheck.ToArray(), searchWord))
                {
                    return true;
                }
            }
            return false;
        }
        public static bool bitStartEnd(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.Contains(searchWord, comparison))
            {
                int withIndex = LineToCheck.IndexOf(searchWord, comparison);
                if (checkStart(withIndex, LineToCheck.ToArray()) & checkEnding(withIndex, LineToCheck.ToArray(), searchWord))
                {
                    return true;
                }
            }
            return false;
        }
        public static bool bitWordNewLIne(string LineToCheck, string searchWord, StringComparison comparison)
        {
            if (LineToCheck.Contains(searchWord, comparison))
            {
                int withIndex = LineToCheck.IndexOf(searchWord, comparison);
                if (checkStart(withIndex, LineToCheck.ToArray()) & checkEndingLine(withIndex, LineToCheck.ToArray(), searchWord))
                {
                    return true;
                }
            }
            return false;
        }
        public static bool startWithAnyOf(string line, List<string> searchWords)
        {
            bool result = false;
            foreach (string Word in searchWords)
            {
                if (line.StartsWith(Word))
                {
                    result = true;
                    break;
                }
            }
            return result;
        }
        public static bool bitValidLabel(string LineToCheck, string LabelMarker, StringComparison comparison)
        {
            if (LineToCheck.Contains(LabelMarker, comparison))
            {
                int index = LineToCheck.IndexOf(':') - 1;

                return ValidLabel(LineToCheck.ToArray(), index);
            }
            return false;
        }

        // New section Taladros

        public static List<int> getVariables(string[] lines, int start, int end, string name) //uses .Contains to easily find
        {
            List<int> indexes = new List<int>();

            for (int i = start; i < end; i++)
            {
                if (lines[i].Contains(name)) //TODO: add case for when there are spaces before Dim
                {
                    indexes.Add(i);
                }
            }
            return indexes;
        }
        public static List<int[]> getFunctions(string[] lines, List<string> SearchWords, string Ender) //uses .Contains
        {
            List<int[]> myResults = new List<int[]>();
            for (int i = 0; i < lines.Length; i++)
            {
                if (startWithAnyOf(lines[i], SearchWords) & lines[i].Contains("(") & lines[i].Contains(")"))//Change this to use the Taladro
                {
                    int end = getEndFunction(lines, i, Ender);
                    if (end == -1)
                    {
                        break;
                    }
                    myResults.Add(new int[] { i, end });
                }
            }
            return myResults;
        }
        public static int getEndFunction(string[] lines, int index, string Ender) //Change this to use the Taladro
        {
            for (int i = index; i < lines.Length; i++)
            {
                if (lines[i].StartsWith(Ender))
                {
                    return i;
                }
            }
            return -1;
        }
        //Level 1.5

        public static List<int> returnVariables(string[] lines, int start, int end)
        {
            List<int> indexes = new List<int>();
            for (int i = start; i < end; i++)
            {
                string trimLine = lines[i].TrimStart(' ');
                if (trimLine.StartsWith("Dim ")) //TODO: Dim FirstSpace As Integer, Slash As Integer : 2 variables in one line
                {
                    indexes.Add(i);
                }
            }
            return indexes;
        }
        public static string vbNameModules(StreamReader sr)
        {
            string name = "";
            string newName = "";
            string line = " ";
            while (line != null)
            {
                line = sr.ReadLine();
                if (line != null)
                {
                    if (line.StartsWith("Attribute VB_Name")) //Attribute VB_Name = "clsExtendedMatching"
                    {
                        name = line.Split(' ')[3]; //Begin VB.Form Frm_Service_Part_List_Suggested
                        newName = name.Replace("\"", "");
                        break;
                    }
                }
            }
            return newName;
        }





        // filters

        public static bool isItemPresent(string searchWord, List<string> myList)//generic return match from list
        {
            bool present = false;
            foreach (string item in myList)
            {
                if (item == searchWord)
                {
                    present = true;
                    break;
                }
            }
            return present;
        }

        public static VBSourceCode returnSourceCode(string realName, List<VBSourceCode> sourceCodes) //generic return match from list
        {
            VBSourceCode mySourceCode = new VBSourceCode();
            foreach (VBSourceCode sc in sourceCodes)
            {
                if (sc.RealName == realName)
                {
                    mySourceCode = sc;
                }
            }
            return mySourceCode;
        }
        public static VBFunction returnFunction(string realName, List<VBFunction> functions)//generic return match from list
        {
            VBFunction myFunction = new VBFunction();
            foreach (VBFunction function in functions)
            {
                if (function.Name == realName)
                {
                    myFunction = function;
                }
            }
            return myFunction;
        }
        public static VBVariable returnVariable(string realName, List<VBVariable> variables)//generic return match from list
        {
            VBVariable myVar = new VBVariable();
            foreach (VBVariable var in variables)
            {
                if (var.Name == realName)
                {
                    myVar = var;
                }
            }
            return myVar;
        }
        public static childSourceCode findChildSimple(string realName, List<childSourceCode> childs)//generic return match from list
        {
            childSourceCode myChild = new childSourceCode();
            foreach (childSourceCode child in childs)
            {
                if (child.Name == realName)
                {
                    myChild = child;
                }
            }
            return myChild;
        }



    }
}
