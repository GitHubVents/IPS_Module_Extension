using Intermech.Imbase;
using Intermech.Interfaces;
using Intermech.Kernel.Search;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddFeatureContextMenu
{
    class AddionalFunctional
    {
        object cell1;
        object cell2;

        private Dictionary<string, string> ExceleData(int sheetOrder, int countR)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult result = dialog.ShowDialog();
            Dictionary<string, string> listExcel = new Dictionary<string, string>();

            if (result == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(dialog.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист)
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[sheetOrder];
                string key = "";
                string value = "";
                for ( int i = 1; i <= countR; i++)
                {
                    cell1 = "A" + i.ToString();
                    cell2 = "B" + i.ToString();
                    
                    Microsoft.Office.Interop.Excel.Range range1 = ObjWorkSheet.Range[cell1];
                    Microsoft.Office.Interop.Excel.Range range2 = ObjWorkSheet.Range[cell2];
                    
                    key = range1.Value.ToString();
                    value = range2.Value.ToString();

                    try
                    {
                        listExcel.Add(key, value);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                ObjExcel.Quit();
            }
            return listExcel;
        }

        private int GetSheetOrder(string tableName)
        {
            int order = 1;

            switch (tableName)
            {
                case "Профили":
                    order = 1;
                break;
                case "Скотч":
                    order = 2;
                    break;
                case "Трубка ПХВ термоусадочная":
                    order = 3;
                    break;
                case "Универсальный герметик":
                    order = 4;
                    break;
                case "Фиксатор":
                    order = 5;
                    break;
                case "Термопреобразователь":
                    order = 6;
                    break;
                case "Предохранители":
                    order = 7;
                    break;
                case "Поролон":
                    order = 8;
                    break;
                case "Плита":
                    order = 9;
                    break;
                case "Хомут монтажный с площадкой":
                    order = 10;
                    break;
            }
            return order;
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
        private long GetOneNeededID(DataTable dt, string tableName)
        {
            foreach (var item in dt.AsEnumerable())
            {
                if (item["Наименование"].ToString() == tableName)
                { 
                    return Convert.ToInt64(item[0]);
                }
            }
            return 0;
        }
        private void Fill_ImbaseTable(long tableID, Dictionary<string, string> erpListWithValues)
        {
            bool flag = false;
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds = TableLoadHelper.GetTables(keeper.Session, tableID, true);
                DataRow row1;

                DataTable tb2 = ds.Tables[Consts.IMS_DATA];
                Dictionary<string, Guid> columnNames = new Dictionary<string, Guid>();

                Guid colGuid;
                string colName;
                
                colGuid = MetaDataHelper.GetAttributeTypeGuid(10);// для наименования
                colName = MetaDataHelper.GetAttributeTypeName(10);
                columnNames.Add(colName, colGuid);
                
                /*
                colGuid = MetaDataHelper.GetAttributeTypeGuid(1025);//Код ОКП
                colName = MetaDataHelper.GetAttributeTypeName(1025);
                columnNames.Add(colName, colGuid);

                foreach (var item in columnNames)
                {
                    MessageBox.Show(item.Key + " " + item.Value.ToString());
                }
                */


                MessageBox.Show("Columns count   " + tb2.Columns.Count);
                try
                {
                    for (int i = 0; i < tb2.Rows.Count; i++)
                    {
                        for (int j = 0; j < tb2.Columns.Count; i++)
                        {
                            MessageBox.Show(tb2.Rows[i][j].ToString());
                        }
                    }
                }
                catch (Exception)
                {

                }
                

                foreach (var item in erpListWithValues.AsEnumerable())
                {
                    //for (int i = 0; i < tb2.Rows.Count; i++)
                    //{

                        //if (tb2.Rows[i]["" + columnNames["Код ОКП"] + ""].ToString() == item.Key.ToString())
                        //{
                        //    flag = true;
                        //}
                    //}
                    if (flag == false)
                    {
                        row1 = tb2.NewRow();
                        //row1.SetField("" + columnNames["Код ОКП"] + "", item.Key);
                        //row1.SetField("" + columnNames["Наименование"] + "", item.Value);
                        //row1["" + columnNames["Код ОКП"] + ""] = item.Key;
                        row1["" + columnNames["Наименование"] + ""] = item.Value;
                        tb2.AcceptChanges();

                        TableLoadHelper.StoreData(keeper.Session, tableID, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
                    }
                 }
            }
        }


        public void MainMethodForFillingImbaseTable(string tableName, int countR)
        {
            int sheetOrder = GetSheetOrder(tableName);
            Dictionary<string, string> erpListWithValues = ExceleData(sheetOrder, countR);

            DataTable allTables = GetImbaseTables();
            long id = GetOneNeededID(allTables, tableName);

            Fill_ImbaseTable(id, erpListWithValues);
        }
    }
}