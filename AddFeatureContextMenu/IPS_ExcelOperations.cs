using Intermech.Imbase;
using Intermech.Interfaces;
using Intermech.Interfaces.Client;
using Intermech.Kernel.Search;
using Intermech.Navigator;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace AddFeatureContextMenu
{
    public class IPS_ExcelOperations
    {
        long newSelectedID;
        // ds - это все таблицы с Excel
        public void MainMethodForFillingImbaseTable(DataSet ds)
        {
            DataTable allTables = GetImbaseTables();
            foreach (var item in ds.Tables)
            {
                DataTable dt = (DataTable)item;
                List<ImbaseTableList> similarTables = GetOneNeededID(allTables, dt.TableName); // id для каждой таблицы 

                if (similarTables.Count == 1)
                {
                    ImbaseTableList itemDt = similarTables.First();
                    Fill_ImbaseTable( itemDt.Id, dt);
                }
                else if (similarTables.Count == 0)
                {
                    MessageBox.Show("There is no table with such name " + dt.TableName);
                }
                else if(similarTables.Count > 1)
                {
                    MessageBox.Show("There are more then one table with name " + dt.TableName + Environment.NewLine + "Choose table in which data has to be written.");
                    
                    ShowContextForm(1069);
                    
                    Fill_ImbaseTable(newSelectedID, dt);
                }
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
        private List<ImbaseTableList> GetOneNeededID(DataTable allTablesImbase, string tableName)
        {
            int tempID = 0;
            List<ImbaseTableList> list = new List<ImbaseTableList>();
            
            foreach (var item in allTablesImbase.AsEnumerable())
            {
                if (item["Наименование"].ToString() == tableName)
                {
                    tempID = Convert.ToInt32(item[0]);
                    list.Add(new ImbaseTableList(tempID, tableName));
                }
            }
            
            return list;
        }
        private void Fill_ImbaseTable(long newSelectedID, DataTable dt)
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds;
                DataTable tb2;
                DataRow row1;
                try
                {
                    ds = TableLoadHelper.GetTables(keeper.Session, newSelectedID, true);

                    tb2 = ds.Tables[Intermech.Imbase.Consts.IMS_DATA];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Таблицы с tableID = " + newSelectedID.ToString() + " не найдено");
                    throw ex;
                }

                Dictionary<string, Guid> columnNamesAndGuids = new Dictionary<string, Guid>();

                byte columnsCount = Convert.ToByte(dt.Columns.Count); // количество колонок в табллице
                List<string> columnNamesinExcel = new List<string>(columnsCount);
                for (byte i = 0; i < columnsCount; i++)
                {
                    columnNamesinExcel.Add(dt.Columns[i].ColumnName); // создали список наименований колонок в таблице
                    columnNamesAndGuids.Add(dt.Columns[i].ColumnName, MetaDataHelper.GetAttributeByTypeNameGuid(dt.Columns[i].ColumnName)); // имя аттрибута и его Guid
                }
                
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

                TableLoadHelper.StoreData(keeper.Session, newSelectedID, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
            }
        }

        Intermech.Navigator.DBObjectTypes.Descriptor desriptor;
        public void ShowContextForm(int type)
        {
            desriptor = new Intermech.Navigator.DBObjectTypes.Descriptor(type);

            DynamicSelectionEventHandler delegat = new DynamicSelectionEventHandler(ImbaseTblSelected);

            SelectionWindow.DynamicSelectObjects("Таблицы IMBASE", "Выберите таблицу в которую будет произведена запись данных", desriptor, delegat, SelectionOptions.Default);
            
            SelectionWindow.OnSelectionWindowAfterClose += SelectionWindow_OnSelectionWindowAfterClose;
        }

        private void SelectionWindow_OnSelectionWindowAfterClose(object sender, EventArgs e)
        {
            SelectionWindow.CloseWindow(desriptor);
        }

        bool ImbaseTblSelected(long id, DynamicSelectionMode mode)
        {
            newSelectedID = id;
            MessageBox.Show("Hello from new method  ImbaseTblSelected. Was selected  " + id.ToString());

            return true;
        }



        public void GetTags()
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                long tableID = TableLoadHelper.GetTableReference(keeper.Session, newSelectedID);
                MessageBox.Show("tableID  " + tableID.ToString());
                DataTable dt = TableLoadHelper.GetTableRefIDsByTableID(keeper.Session, tableID);
                string columnns = "";
                foreach (var item in dt.Columns)
                {
                    columnns += item.ToString() + Environment.NewLine;
                }
                foreach (var item in dt.Rows)
                {
                    string values = ((DataRow)item)[0].ToString() + "  " + ((DataRow)item)[1].ToString();
                    MessageBox.Show("Rows  " + values);
                }
                MessageBox.Show(columnns);
            }
        }
    }



    //Описывает таблицы IMBASE, если с указанным наименованием существует больше одной
    public class ImbaseTableList
    {
        public ImbaseTableList(long id, string name)
        {
            Id = id;
            Name = name != null ? name : string.Empty;
        }

       public long Id { get; set; }
       public string Name { get; set; }
    }
}