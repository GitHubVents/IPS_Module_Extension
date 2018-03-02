using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace AddFeatureContextMenu
{
    class ExcelFunctionality
    {
        static private OleDbConnection exCon = null;
        static private string pathToLocalFile;

        public ExcelFunctionality(string PathTofile)
        {
            if (PathTofile != string.Empty && PathTofile != null)
            {
                pathToLocalFile = PathTofile;
            }
        }
        static private OleDbConnection ExcelConnection
        {
            get
            {
                if (exCon == null)
                {
                    exCon = MakeConnection(pathToLocalFile);
                }
                return exCon;
            }
        }
        private static OleDbConnection MakeConnection(string pathToFile)
        {
            exCon = new OleDbConnection();
            exCon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                pathToFile + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'";
            return exCon;
        }


        public static DataSet GetTotalSheetsData()
        {
            DataSet dsData = new DataSet();
            string[] allSheets = GetSheetNames();

            foreach (var item in allSheets)
            {
                dsData.Tables.Add( GetDataFromExcel(item));
            }
            return dsData;
        }
        private static string[] GetSheetNames()
        {
            DataTable dtTables = new DataTable();
            ExcelConnection.Open();
            dtTables = ExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            ExcelConnection.Close();
            string[] excelSheets = null;
            if ((dtTables != null))
            {
                excelSheets = new string[dtTables.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dtTables.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
            }
            return excelSheets;
        }
        public static DataTable GetDataFromExcel(string sheetName)
        {
            string command = "SELECT * FROM " + "[" + sheetName + "]";
            DataTable dt = new DataTable();

            string tempName = sheetName.Replace("$", string.Empty); // имя таблицы совпадает с именем листа


            //проверка являеться ли первый/ последний символ имени листа буквой
            Char[] eachSymb = tempName.ToCharArray();
            Char firstLetter = eachSymb[0];
            Char lastLetter = eachSymb[eachSymb.Length - 1];
            if (!Char.IsLetter(firstLetter) && !Char.IsNumber(firstLetter))
            {
                tempName = tempName.Remove(0, 1);
            }
            if (!Char.IsLetter(lastLetter) && !Char.IsNumber(lastLetter))
            {
                tempName = tempName.Remove(tempName.Length - 1);
            }

            dt.TableName = tempName;


            using (OleDbCommand comm = new OleDbCommand())
            {
                comm.CommandText = command;

                comm.Connection = ExcelConnection;

                using (OleDbDataAdapter da = new OleDbDataAdapter())
                {
                    da.SelectCommand = comm;
                    da.Fill(dt);
                    da.Dispose();

                    return dt;
                }
            }
        }
    }
}