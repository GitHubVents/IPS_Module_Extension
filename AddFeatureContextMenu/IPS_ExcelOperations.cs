using Intermech;
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
        // ds - это все таблицы с Excel
        public void MainMethodForFillingImbaseTable(DataSet ds)
        {
            DataTable allTables = GetImbaseTables();

            foreach (var item in ds.Tables)
            {
                DataTable dt = (DataTable)item;
                MessageBox.Show(dt.TableName);
                List<ImbaseTableList> similarTables = GetOneNeededID(allTables, dt.TableName); // id для каждой таблицы 

                if (similarTables.Count == 1)
                {
                    ImbaseTableList itemDt = similarTables.First();
                    Fill_ImbaseTable( itemDt.Id, dt);
                }
                else
                {
                    MessageBox.Show("There are more then one table with name " + dt.TableName);
                    int idNewType = CreateTemporarType(similarTables);
                    ShowContextForm(idNewType);
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
            int[][] massivMassivov = new int[][] { };
            int tempID = 0;
            List<ImbaseTableList> list = new List<ImbaseTableList>();


            foreach (var item in allTablesImbase.AsEnumerable())
            {
                if (item["Наименование"].ToString() == tableName)
                {
                    MessageBox.Show(item[0].ToString() + "  " + tableName);
                    tempID = Convert.ToInt32(item[0]);
                    list.Add(new ImbaseTableList(tempID, tableName));
                }
            }
            
            return list;
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
                    tb2 = ds.Tables[Intermech.Imbase.Consts.IMS_DATA];

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

        Intermech.Navigator.DBObjectTypes.Descriptor desriptor;
        public void ShowContextForm(int type)
        {
            desriptor = new Intermech.Navigator.DBObjectTypes.Descriptor(type);

            DynamicSelectionEventHandler delegat = new DynamicSelectionEventHandler(ImbaseTblSelected);
            //long[] selectAll = SelectionWindow.SelectObjects("Selection", "d", desriptor, SelectionOptions.DisableMultiselect);

            SelectionWindow.DynamicSelectObjects("Selection", "d", desriptor, delegat, SelectionOptions.Default);


            #region
            /*
            Intermech.Navigator.Controls.ISelectedItemsAnalyzer analizer = new Intermech.Navigator.Controls.SelectedItemsAnalyzer();
            
            var service = this.desriptor.GetService(typeof(Intermech.Navigator.Controls.ISelectedItemsHost));

            ISelectedItemsText selected = new Intermech.Navigator.Controls.SelectedItemsAnalyzer();

            analizer.Analyze(this, (Intermech.Navigator.Controls.ISelectedItemsHost)service);

            SelectionWindow.RegisterAnalyze(analizer);
            */

            // var r = SelectionWindow.CreateForm("Formochka", "Форма для отображения таблиц IMBASE с одинаковыми названиями",
            //   desriptor, typeof(string), delegat, SelectionOptions.DisableMultiselect);

            //SelectionWindow.OnSelectionWindowAfterClose += SelectionWindow_OnSelectionWindowAfterClose;
            #endregion
        }

        private void SelectionWindow_OnSelectionWindowAfterClose(object sender, EventArgs e)
        {
            MessageBox.Show("The end");
        }

        bool ImbaseTblSelected(long id, DynamicSelectionMode mode)
        {
            MessageBox.Show("Hello from new method    " + id.ToString());

            ShowContextForm((int)id);

            return true;
        }

        public int CreateTemporarType(List<ImbaseTableList> id_s)
        {
            int idNewType = 0; // id нового типа

            using (SessionKeeper keeper = new SessionKeeper())
            {
                    /*
                    //создали новый тип
                    IDBObjectTypeCollection col = keeper.Session.GetObjectTypeCollection(1007);
                
                    ObjectTypeProperties properties = ObjectTypeProperties.Empty();
                    properties.ObjectTypeGuid = new Guid();
                    properties.ObjectInstanceName = "TemporaryTable";
                    properties.ObjectTypeName = "SimilarTables2";
                    properties.Versionable = ObjectVersionModes.SingleVersion;
                    properties.AnyAttributes = true;
                    properties.Options = ObjectTypeOptions.LocalObjectType;
                    properties.PublicLCSchema = InheritModes.Private;
                    properties.SchemaID = 1;


                    idNewType = col.Create(properties);
                    */
                //берем коллекцию обьектов нового типа
                IDBObjectCollection coll = keeper.Session.GetObjectCollection(1897);
                IDBObject newObj = null;
                //Intermech.Client.Core.FormDesigner.
                foreach (var item in id_s)
                {
                    newObj = coll.Create(idNewType);
                    
                    //добавляем атрибуты объекта
                    IDBAttribute atrName = newObj.GetAttributeByID(9);
                    IDBAttribute atrDesignation = newObj.GetAttributeByID(10);
                    newObj.ObjectType = 1897;
                    atrDesignation.Value = item.Id.ToString();
                    //atrName.Value = item.Name.ToString();
                    MessageBox.Show(newObj.ID + "  " + newObj.ObjectID);
                    newObj.CommitCreation(true);
                }
            }

            return idNewType;
        }
        private void DeleteTempType(int type)
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                //создали новый тип
                IDBObjectTypeCollection col = keeper.Session.GetObjectTypeCollection(1007);
                IDBObjectType typeObj = keeper.Session.GetObjectType(type);
                int resDel = typeObj.Delete((int)DeleteObjectsJobMode.AbortOnError);
                MessageBox.Show("delete tempType = " + resDel.ToString());
            }
        }




    }



    //Описывает 
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