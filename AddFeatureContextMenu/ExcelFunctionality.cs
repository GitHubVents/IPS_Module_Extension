using Intermech.Imbase;
using Intermech.Interfaces;
using Intermech.Kernel.Search;
using System;
using System.Collections.Generic;
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
            pathToLocalFile = PathTofile;
        }
        static private OleDbConnection ExcelConnection
        {
            get
            {
                if (exCon == null)
                {
                    exCon = MakeConnection(pathToLocalFile);
                    MessageBox.Show("First read");
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


        private void Fill_ImbaseTable(long tableID, DataTable dt)
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds;
                DataTable tb2;
                DataRow row1;

                try
                {
                    ds = TableLoadHelper.GetTables(keeper.Session, tableID, true);
                    tb2 = ds.Tables[Consts.IMS_DATA];

                    foreach (var item in tb2.Columns)
                    {
                        MessageBox.Show(((DataColumn)item).ColumnName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Таблицы с указанным названием не найдено");
                    throw ex;
                }

                Dictionary<string, Guid> columnNamesAndGuids = new Dictionary<string, Guid>();

                byte columnsCount = Convert.ToByte(dt.Columns.Count); // количество колонок в табллице
                List<string> columnNamesinExcel = new List<string>(columnsCount);
                for (byte i = 0; i < columnsCount; i++)
                {
                    MessageBox.Show("columnsCount   " + dt.Columns[i].ColumnName);
                    columnNamesinExcel.Add(dt.Columns[i].ColumnName); // создали список наименований колонок в таблице
                    columnNamesAndGuids.Add(dt.Columns[i].ColumnName, MetaDataHelper.GetAttributeByTypeNameGuid(dt.Columns[i].ColumnName)); // имя аттрибута и его Guid
                }
                //Intermech.SystemGUIDs

                string tempGuid;
                string tempName;      
                foreach (var item in dt.AsEnumerable()) // по всем рядкам
                {
                    row1 = tb2.NewRow();

                    for (byte i = 0; i < columnsCount; i++) // по каждой колонке
                    {
                        tempName = columnNamesinExcel[i];
                        tempGuid = "" + columnNamesAndGuids[tempName] + "";

                        row1.SetField(tempGuid, item[tempName].ToString());
                    }
                    tb2.Rows.Add(row1);
                    tb2.AcceptChanges();
                }
                
                TableLoadHelper.StoreData(keeper.Session, tableID, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
            }
        }


        public void MainMethodForFillingImbaseTable()
        {
            DataTable allTables = GetImbaseTables();

            DataSet ds = GetTotalSheetsData();

            foreach (var item in ds.Tables)
            {
                DataTable dt = (DataTable)item;
                MessageBox.Show(dt.TableName);
                long id = GetOneNeededID(allTables, dt.TableName); // id для каждой таблицы 

                Fill_ImbaseTable(id, dt);
            }
        }
        private DataTable GetImbaseTables()
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                IDBObjectCollection objCollection = keeper.Session.GetObjectCollection(1069);

                ConditionStructure[] conditions = new ConditionStructure[] { };
                object[] columns = new object[]
                {
                       ObligatoryObjectAttributes.F_OBJECT_ID, "Наименование"
                };

                object[] sortColumns = new object[]
                {
                         ObligatoryObjectAttributes.F_OBJECT_ID, "Наименование",
                };

                SortOrders[] order = new SortOrders[] { SortOrders.ASC, SortOrders.ASC };

                DBRecordSetParams pars = new DBRecordSetParams(conditions, columns, sortColumns, order, 0, null, QueryConsts.All, true, "MyIMBASE_TableSelection");

                DataTable dt = objCollection.Select(pars);

                return dt;
            }
        }
        private long GetOneNeededID(DataTable allTablesImbase, string tableName)
        {
            int[][] massivMassivov = new int[][] { };
            int counter = 0;
            foreach (var item in allTablesImbase.AsEnumerable())
            {
                if (item["Наименование"].ToString() == tableName)
                {
                    MessageBox.Show(item[0].ToString() + "  " + tableName);
                    mas.Add(Convert.ToInt32( item[0]), tableName);
                    counter++;
                }
            }
            if (mas.Count > 1)
            {

            }
            else
            {
                return 
            }
            return Convert.ToInt64(item[0]);
        }

        
        //  cad0038a-306c-11d8-b4e9-00304f19f545' 



        private static DataSet GetTotalSheetsData()
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
        public static DataTable GetDataFromExcel(string sheet)
        {
            string command = "SELECT * FROM " + "[" + sheet + "]";

            DataTable dt = new DataTable();
            dt.TableName = sheet.Replace("$", ""); // имя таблицы совпадает с именем листа

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