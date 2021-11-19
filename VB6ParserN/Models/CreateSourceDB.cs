using System.Data.SqlClient;

namespace VB6ParserN.Models
{
    public static class CreateSourceDB
    {
        // SourceDB creation functions
        // TODO: This should be rewritten using Stored Procedures
        public static int VBProject(string ProjectPath, SqlConnection cnn, string Table = "ProjectFile_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(ProjectPath)" + "output INSERTED.ID " +
                " values(" +
                "'" + ProjectPath + "'" + ")";
            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar(); //when adding records (INSERT), the amount of fields have to match    
        }

        public static int SourceFile(string SourcePath, int Project, string VBName, int Type, int Inject, SqlConnection cnn, string Table = "SourceFile_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(SourcePath, ProjectFile_Table_ID, VBName, SourceType_Table_ID, Inject)" + "output INSERTED.ID " +
                " values(" +
                 "'" + SourcePath + "'" + "," +
                 Project.ToString() + "," +
                "'" + VBName + "'" + "," +
                Type.ToString() + "," +
                Inject.ToString() + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }

        public static int Function(string FunctionName, int SourceId, int ReturnType, int PrivacyType, int Sub, int Inject, int LineStart, int LineEnd, SqlConnection cnn, string Table = "Function_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(FunctionName, SourceFile_Table_ID, ReturnType_Table_ID, PrivacyType_Table_ID, Sub, Inject, LineStart, LineEnd)" + "output INSERTED.ID " +
                " values(" +
                 "'" + FunctionName + "'" + "," +
                SourceId.ToString() + "," +
                ReturnType.ToString() + "," +
                PrivacyType.ToString() + "," +
                Sub.ToString() + "," +
                Inject.ToString() + "," +
                LineStart.ToString() + "," +
                LineEnd.ToString() + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }

        public static int InstanceFunction(int FunctionId, int SourceFile, int LineNumber, int ColumnNumber, SqlConnection cnn, string Table = "FunctionInstances_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(Function_Table_ID, SourceFile_Table_ID, LineNumber, ColumnNumber)" + "output INSERTED.ID " +
                " values(" +
                FunctionId.ToString() + "," +
                SourceFile.ToString() + "," +
                LineNumber.ToString() + "," +
                ColumnNumber.ToString() + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }

        public static int Variable(string VariableName, int FunctionId, int VariableType, int Argument, int Inject, SqlConnection cnn, string Table = "Variable_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(VariableName, Function_Table_ID, ReturnType_Table_ID, Argument, Inject)" + "output INSERTED.ID " +
                " values(" +
                 "'" + VariableName + "'" + "," +
                FunctionId.ToString() + "," +
                VariableType.ToString() + "," +
                Argument.ToString() + "," +
                Inject.ToString() + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }

        public static int Instance(int VariableId, int LineNumber, int ColumnNumber, SqlConnection cnn, string Table = "VariableInstances_Table")
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(Variable_Table_ID, LineNumber, ColumnNumber)" + "output INSERTED.ID " +
                " values(" +
                VariableId.ToString() + "," +
                LineNumber.ToString() + "," +
                ColumnNumber.ToString() + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }

        public static int GetReturnTypeId(string Table, string columnName, string Type, SqlConnection cnn)
        {
            SqlCommand command;
            SqlDataReader dataReader;
            string sql;

            sql = "SELECT [ID] FROM [" + Table + "] Where " + columnName + " = '" + Type + "'";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            int returnValue = -1;
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    returnValue = (int)dataReader.GetValue(0); //asuming there is only one match!                    
                }
                dataReader.Close();
            }
            else
            {
                returnValue = -1;
                dataReader.Close();
            }

            return returnValue;
        }

        public static int InsertNewType(string Table, string Column, string NewType, SqlConnection cnn)
        {
            SqlCommand command;
            string sql2 = "INSERT into " +
                Table + "(" + Column + ")" + "output INSERTED.ID " +
                " values(" + "'" + NewType + "'" + ")";

            command = new SqlCommand(sql2, cnn);
            return (int)command.ExecuteScalar();
        }
    }
}
