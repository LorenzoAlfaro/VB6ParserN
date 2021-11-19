using System;
using System.Collections.Generic;
using System.Linq;
using static StringParserN.CharWorld;

namespace VB6ParserN.Models
{
    public static class LogicToTree
    {
        public static Tree getTrunk4(string[] FunctionLines, int start, int end, DrillBit bit, string[] startKeywords, string[] endKeywords, string[] splitKeywords)
        {
            //string[] startKeywords = new string[] { "IF THEN", "FOR", "DO WHILE", "SELECT CASE", "DO UNTIL", "WHILE", "DO", "FUNCTION"};
            //string[] endKeywords = new string[] { "END IF", "NEXT", "LOOP", "END SELECT", "LOOP", "WEND", "LOOP_WHILE", "END FUNCTION"};
            //string[] splitKeywords = new string[] { "ELSE", "", "", "Case", "", "", "", ""};

            Tree myTree = new Tree(start - 1, end);
            List<Tree> myBranches = new List<Tree>();

            for (int i = start; i < end; i++)
            {
                int[] startBranch = TrunkStart3(FunctionLines, i, end, bit, startKeywords); //index, type
                bool completeTree = false;
                if (startBranch[0] == -1)
                {
                    break;
                }
                else
                {
                    int endBranch = TrunkFinalEnd3(FunctionLines, startBranch[0] + 1, end, bit, startKeywords, endKeywords, startBranch[1]);
                    Tree subTree = new Tree(startBranch[0], endBranch);
                    if (endBranch != -1)
                    {
                        completeTree = true;
                        switch (startBranch[1])
                        {
                            //Console.WriteLine("Select case");
                            case 3:
                            case 0: //IF

                                List<int> Splits = new List<int>();
                                Splits.Add(startBranch[0]);
                                Splits.AddRange(subRanges(FunctionLines, startBranch[0] + 1, endBranch, bit, startKeywords, splitKeywords[startBranch[1]], endKeywords, 0));
                                Splits.Add(endBranch);

                                if (Splits.Count() == 2)
                                {
                                    subTree = getTrunk4(FunctionLines, startBranch[0] + 1, endBranch, bit, startKeywords, endKeywords, splitKeywords);
                                }
                                else
                                {
                                    List<Tree> subBranches = new List<Tree>();
                                    for (int j = 1; j < Splits.Count(); j++)
                                    {
                                        subBranches.Add(getTrunk4(FunctionLines, Splits[j - 1] + 1, Splits[j], bit, startKeywords, endKeywords, splitKeywords));
                                    }
                                    subTree.Branches = subBranches;
                                    subTree.GroupType = startBranch[1];
                                    subTree.SplitGroup = true;
                                }
                                break;
                            case 1: //FOR                                                                
                            case 2: //DO WHILE
                            //case 3: //Select End Select
                            case 4: //DO UNTIL
                            case 5: //WHILE
                            case 6: //DO
                            case 7: //Function End Function

                                subTree = getTrunk4(FunctionLines, startBranch[0] + 1, endBranch, bit, startKeywords, endKeywords, splitKeywords);
                                subTree.GroupType = startBranch[1];
                                Console.WriteLine("Case No Split");
                                break;
                            default:
                                Console.WriteLine("Default case");
                                break;
                        }
                        i = endBranch; //>>>>>>>>>>>Forwarded
                        if (completeTree)
                        {
                            myBranches.Add(subTree);
                        }
                    }
                }
            }
            myTree.Branches = myBranches;
            return myTree;
        }
        public static int[] TrunkStart3(string[] FunctionLines, int current, int end, DrillBit bit, string[] startKeywords)
        {
            // index, and boolean Else? End If
            for (int i = current; i < end; i++)
            {
                int type = StartType(FunctionLines[i].TrimStart(), startKeywords, bit);
                if (type != -1)
                {
                    return new int[] { i, type };
                }

            }
            return new int[] { -1, -1 };
        }
        public static int StartType(string line, string[] startKeywords, DrillBit bit)
        {
            for (int i = 0; i < startKeywords.Length; i++)
            {
                if (bit(line, startKeywords[i])) //If then else end if
                {
                    //where & what
                    return i;
                }
            }
            return -1;
        }
        public static int TrunkFinalEnd3(string[] FunctionLines, int current, int endSearch, DrillBit bit, string[] startKeywords, string[] endKeywords, int type)// will ignore subBranches 'ELSE, else if'
        {
            for (int i = current; i < endSearch; i++)
            {
                int StarType2 = StartType(FunctionLines[i].TrimStart(), startKeywords, bit);//may go off the function
                int offSet = 0;
                int subEnd = 0;


                if (StarType2 != -1) //Found a new start of whatever
                {
                    subEnd = TrunkFinalEnd3(FunctionLines, i + 1, endSearch, bit, startKeywords, endKeywords, StarType2);//find the end of the new struct
                    if (subEnd == -1) { break; }

                    i = subEnd; //>>>>>>>>>>Forwarded
                    offSet = 1;

                }

                if (bit(FunctionLines[i + offSet].TrimStart(), endKeywords[type]))
                {
                    return i + offSet;
                }

            }
            return -1;
        }
        public static List<int> subRanges(string[] FunctionLines, int current, int end, DrillBit bit, string[] startKeywords, string splitKeyword, string[] endKeywords, int type)
        {
            List<int> SubRanges = new List<int>();//the first int, the start 'If dfdfdf then '//middle ones 'else if' 'else if'//before the closing, 'else'//last 'End If//if List only has one, one line if,//if List only has 2, If- End If// If List > 3, If ElseIF Else if                   
            for (int i = current; i < end; i++)
            {
                int StarType2 = StartType(FunctionLines[i].TrimStart(), startKeywords, bit);//may go off the function
                int offSet = 0;
                if (StarType2 != -1)
                {
                    int finalEndIf = TrunkFinalEnd3(FunctionLines, i + 1, end, bit, startKeywords, endKeywords, StarType2);
                    if (finalEndIf != -1)
                    {
                        i = finalEndIf; //>>>Forward
                        offSet = 1;
                    }
                    else
                    {
                        break;
                    }
                }

                if (bit(FunctionLines[i + offSet].TrimStart(), splitKeyword)) //Else, else if
                {
                    SubRanges.Add(i + offSet);
                }
            }
            return SubRanges;
        }
        public static string[] tabFile(Tree logicTree, string[] lines, bool tab)
        {

            if (!logicTree.SplitGroup)
            {
                for (int i = logicTree.Trunk[0] + 1; i < logicTree.Trunk[1]; i++)
                {
                    lines[i] = '\t' + lines[i];
                }
            }
            else
            {
                if (logicTree.GroupType == 3) //'Select Case' special case of format
                {
                    foreach (Tree branch in logicTree.Branches)
                    {
                        lines[branch.Trunk[1]] = '\t' + lines[branch.Trunk[1]];
                        tabFile(branch, lines, true);
                    }

                    lines[logicTree.Trunk[1]] = lines[logicTree.Trunk[1]].Remove(0, 1);
                }
            }

            foreach (Tree branch in logicTree.Branches)
            {
                tabFile(branch, lines, true);
            }

            return lines;
        }
    }
}
