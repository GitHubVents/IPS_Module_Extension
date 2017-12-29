using System;
using Intermech;
using Intermech.Bars;
using Intermech.Interfaces;
using Intermech.Interfaces.Plugins;
using Intermech.Interfaces.Client;
using Intermech.NavBars;
using System.Runtime.InteropServices;
using Intermech.Runtime.ComInterop.LocalServer;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;
using Intermech.Imbase;
using Vents_PLM.Spigot;

namespace AddFeatureContextMenu
{
    [ComVisible(true), Guid("F8358B23-3E51-453C-8EF8-9ABD22525214"), ProgId("IPS.EditContextMenu")]
    [ClassInterface(ClassInterfaceType.None)]

    public class Operations_with_IPS : IPackage
    {
        int objType = 0;
        DataSet ds = null;
        DataTable tb1 = null;//таблица IMBASE с данными
        Dictionary<string, string> dictionary;//наименования колонок таблицы IMBASE и их значения

        AddionalFunctional exel = new AddionalFunctional();
        UserControl user;
        internal static IServiceProvider _serviceProvider;

        public void Load(IServiceProvider serviceProvider)
        {
            if (ComHost.Configuration.ComSupportActive)
            {
                ComHost.ActivateClassFactory(typeof(Operations_with_IPS));
            }

            _serviceProvider = serviceProvider;

            AddMenuItems();
        }
        public void Unload()
        {
            if (ComHost.Configuration.ComSupportActive)
            {
                ComHost.DeactivateClassFactory(typeof(Operations_with_IPS));
            }
        }
        string IPackage.Name
        {
            get
            {
                return "AddFeatureContextMenu" + DateTime.Now.DayOfWeek.Name() + DateTime.UtcNow.ToString();
            }
        }


        public void AddMenuItems()
        {
            // Получаем ссылку на управление содержимым докинга
            Intermech.Docking.IContentProvider provider = ServicesManager.GetService(typeof(Intermech.Docking.IContentProvider)) as Intermech.Docking.IContentProvider;
            // Подписываемся на событие "Восстановить окно"
            provider.ContentCallback += new Intermech.Docking.GetContentCallback(WellKnownDockWindow.RestoreWindowCallback);


            // Получаем ссылку на сервис именованных значков
            INamedImageList _images = _serviceProvider.GetService(typeof(INamedImageList)) as INamedImageList;

            // Получаем ссылку на службу панелей приложений "Навигатора"
            INavigationBar navBar = ServicesManager.GetService(typeof(INavigationBar)) as INavigationBar;

            // Служба найдена
            if (navBar != null)
            {
                // Ищем панель "Приложения"
                IAppPane pane = navBar.FindPane("appPane") as IAppPane;

                // Панель найдена
                if (pane != null)
                {
                    pane.Add("AAAAAAAAAAAAAAAAAAAAAAAAAAAH",
                        new EventHandler(WellKnownDockWindow.ShowWellKnownDockWindow),
                        _images.ImageIndex(WellKnownDockWindow.WindowImageName));

                    //получаем обьект главного меню
                    BarManager barManager = _serviceProvider.GetService(typeof(BarManager)) as BarManager;
                    //добавляем в него новый Item - AirVentsCAD
                    MenuBarItem airVentsMenuItem = barManager.MenuBar.AddMenuBar("AirVentsCAD");
                    airVentsMenuItem.ForeColor = System.Drawing.Color.MediumAquamarine;
                    //создаем пункты подменю
                    airVentsMenuItem.Items.Add("Spigot", Spigot_EventBuilder);
                    airVentsMenuItem.Items.Add("Roof", Roof_EventBuilder);
                    airVentsMenuItem.Items.Add("Добавить электронный документ", AddDoc_EventBuilder);
                }
            }
        }

        /////////отображение контролов
        private void Spigot_EventBuilder(object sender, EventArgs e)
        {
            user = new SpigotControl();
            user.Dock = DockStyle.Fill;
            
            Form form = new Form();
            form.Size = new System.Drawing.Size(600, 500);
            form.Location = new System.Drawing.Point(500, 300);
            form.Controls.Add(user);
            user.Show();
            form.Show();
        }
        private void Roof_EventBuilder(object sender, EventArgs e)
        {
            // user = new RoofBilder();
            user.Dock = DockStyle.Fill;

            Form form = new Form();
            form.Size = new System.Drawing.Size(600, 500);
            form.Location = new System.Drawing.Point(500, 300);
            form.Controls.Add(user);
            user.Show();
            form.Show();
        }
        private void AddDoc_EventBuilder(object sender, EventArgs e)
        {
            exel.MainMethodForFillingImbaseTable("Скотч", 39);
        }


        private void ReadBlob()
        {
            using (SessionKeeper sk = new SessionKeeper())
            {
                IDBObject iDBObject = sk.Session.GetObject(1962223);
                IDBAttributable iDBAttributable = iDBObject as IDBAttributable;
                IDBAttribute[] iDBAttributeCollection = iDBAttributable.Attributes.GetAttributesByType(FieldTypes.ftFile);//получили аттрибут File
                
                if (iDBAttributeCollection != null)

                foreach (IDBAttribute fileAtr in iDBAttributeCollection)
                {
                    if (fileAtr.ValuesCount < 2)
                        {
                            MemoryStream m = new MemoryStream();
                            BlobProcReader reader = new BlobProcReader(fileAtr, (int)fileAtr.AsInteger, m, null, null);
                            MessageBox.Show("ElementID" + reader.ElementID.ToString() + "DataBlockSize =  " + reader.DataBlockSize.ToString() + "fileAtr.AsInteger  " + fileAtr.AsInteger.ToString() + "  fileAtr.AsDouble  " + fileAtr.AsDouble.ToString());
                            BlobInformation info = reader.BlobInformation;
                            reader.ReadData();
                            byte[] @byte = m.GetBuffer();
                            MessageBox.Show(@byte.Length.ToString());
                            //BlobProcReader reader = new BlobProcReader((long)info.BlobID, (int)AttributableElements.Object, fileAtr.AttributeID, 0, 0, 0);
                            


                            //try
                            //{
                            //    IVaultFileReaderService serv = sk.Session.GetCustomService(typeof( IVaultFileReaderService)) as IVaultFileReaderService;
                            //    IVaultFileReader fileReader = serv.GetVaultFileReader(sk.Session.SessionID);
                            //    if (fileReader != null) MessageBox.Show("file reader exist       RealFileSize:" + info.RealFileSize.ToString() + Environment.NewLine + "PackedFileSize:" + info.PackedFileSize.ToString() + Environment.NewLine + iDBObject.ObjectID.ToString() + Environment.NewLine + fileAtr.AsInteger.ToString());

                            //    BlobInformation item = fileReader.OpenBlob(0, 1962223, 0, 8);
                            //    MessageBox.Show("Blob is open");

                            //    byte[] @byt = fileReader.ReadDataBlock((int)item.PackedFileSize);
                            //    MessageBox.Show(@byt.GetLength(@byt.Rank).ToString());
                            //}
                            //catch { MessageBox.Show("No service found"); }
                            //return;
                        }
                    }
            }
        }
        /// <summary>
        /// Запись файлов в DB IPS
        /// </summary>
        /// <param name="filePath">Путь к файлу, который сохраняем</param>
        /// <param name="objType">Тип файла</param>
        /// <param name="attrName">Наименование</param>
        /// <param name="attrDesignation">Обозначение</param>
        public long WriteBlob(SessionKeeper sk, string filePath, int objType, string attrName, string attrDesignation)
        {
            MessageBox.Show(filePath);
            long id = 0;
                
            int atrFile = MetaDataHelper.GetAttributeTypeID(new Guid(SystemGUIDs.attributeFile));// аттрибут "Файл"
            IDBObjectCollection coll = sk.Session.GetObjectCollection(objType);
            IDBObject newObj = coll.Create(objType);
            IDBAttribute fileAtr = newObj.GetAttributeByID(atrFile);

            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(filePath, FileMode.Open);
                MessageBox.Show(fileStream.Name);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }

            int bytesFromSTREAM = fileStream.ReadByte();
            MessageBox.Show(bytesFromSTREAM.ToString());
                
            try
            // write into DB
            {                    
                BlobInformation blInfo = new BlobInformation(0, 0, DateTime.Now, filePath, ArcMethods.NotPacked, note: "JustJustJust_OneMore");
                BlobProcWriter writer = new BlobProcWriter(fileAtr, (int)AttributableElements.Object, blInfo, fileStream, null, null);
                writer.WriteData();
                try
                {
                    // добавление создаваемого обьекта в Документы/Конструкторские/Электронные модели
                    IDBAttribute atrName = newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeName));//наименование обьекта
                    IDBAttribute atrDesignation = newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeDesignation));//обозначение обьекта
                    atrName.Value = attrName;
                    atrDesignation.Value = attrDesignation;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                newObj.CommitCreation(true);
                id = newObj.ObjectID;
                MessageBox.Show("Data has been written. New objectID is:    " + newObj.ObjectID.ToString());
                return id;
            }
            catch (Exception e) { MessageBox.Show(e.Message); }
          
            return id;
        }


        //private void WriteIntoIMBASE_Roof_Table(string width, string height, string roofType)//ROOF
        //{
        //    using (SessionKeeper keeper = new SessionKeeper())
        //    {
        //        GetIMBASETable((long)ImbaseTableID.Roof, out ds, out dictionary);
        //        DataRow row1 = tb1.NewRow();

        //        if (CheckForSimilarRows(width, height, roofType, tb1))
        //        {
        //            return;
        //        }
        //        else
        //        {
        //            row1[Intermech.Imbase.Consts.F_GUID] = new Guid();
        //            row1[""+dictionary["Ширина"]+""] = width;
        //            row1[""+dictionary["Высота"]+""] = height;
        //            row1[""+dictionary["Тип крыши"]+""] = roofType;

        //            tb1.Rows.Add(row1);
        //            tb1.AcceptChanges();
        //            TableLoadHelper.StoreData(keeper.Session, (long)ImbaseTableID.Roof, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
        //        }
        //    }
        //}

        public void WriteIntoIMBASE_Spigot_Table(DataTable tb1, string filePath, string oboznachenie, string naimenovanie, string width, string height, string flangeType, int elementType)//SPIGOT
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataRow row1 = tb1.NewRow();

                objType = GetObjectType(elementType);

                long id = WriteBlob(keeper, filePath, objType, naimenovanie, oboznachenie);
                MessageBox.Show("WriteBlob   " + id.ToString());
                if (CreateProduct(elementType, id, naimenovanie, oboznachenie))
                {
                    ////////////запись в IMBASE
                    row1[Intermech.Imbase.Consts.F_GUID] = new Guid();
                    row1["" + dictionary["Ширина"] + ""] = width;
                    row1["" + dictionary["Высота"] + ""] = height;
                    row1["" + dictionary["Тип фланца"] + ""] = flangeType;
                    row1["" + dictionary["Element Type"] + ""] = elementType;
                    row1["" + dictionary["Обозначение"] + ""] = naimenovanie;

                    tb1.Rows.Add(row1);
                    tb1.AcceptChanges();
                    TableLoadHelper.StoreData(keeper.Session, (long)ImbaseTableID.Spigot, ds,
                                                keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer))
                                                    as Intermech.Interfaces.Imbase.ITablesIndexer);
                    MessageBox.Show("This is uspeh!");
                }
            }
        }

        private int GetObjectType(int elementType)
        {
            switch (elementType)
            {
                case 0:
                    objType = 1361; // сборка
                    break;
                case 1:
                    objType = 1296; // деталь
                    break;
                case 2:
                    objType = 1324; // чертеж
                    break;
            }
            return objType;
        }



        public DataTable GetIMBASETable(long tableID, out DataSet ds, out Dictionary<string, string> dictionary)
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                ds = TableLoadHelper.GetTables(keeper.Session, tableID, true);
                tb1 = ds.Tables[0];
                DataTable tb2 = ds.Tables[1];
                string colName = "";
                dictionary = new Dictionary<string, string>();
                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    try
                    {
                        colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                        dictionary.Add(colName, tb1.Columns[i].ToString());
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("There is no such attribute");
                        //throw;
                    }
                }
                return tb1;
            }
        }

        //public bool CheckForSimilarRows(string width, string height, string roofType, DataTable table)
        //{
        //    bool itemExist = false;
        //    int flag = 1;
        //    L1:
        //        foreach (var item in table.AsEnumerable())
        //        {
        //            if (item[dictionary["Ширина"]].ToString() == width 
        //                && item[dictionary["Высота"]].ToString() == height
        //                && item[dictionary["Тип крыши"]].ToString() == roofType)
        //            {
        //                MessageBox.Show("Такая деталь существует");

        //                return itemExist = true;
        //            }
        //        }
        //    flag--;
        //    if (itemExist == false && flag == 0)
        //    {
        //        string temp = width;
        //        width = height;
        //        height = temp;
        //        goto L1;
        //    }
        //    return itemExist;
        //}






        //public bool CheckForSimilarRows(string width, string height, string flangeType, string elementType, DataTable table)
        //{
        //    bool itemExist = false;
        //    int flag = 1;
        //    L1:
        //    foreach (var item in table.AsEnumerable())
        //    {
        //        if (item[dictionary["Ширина"]].ToString() == width
        //            && item[dictionary["Высота"]].ToString() == height
        //            && item[dictionary["Тип фланца"]].ToString() == flangeType)
        //        {
        //            MessageBox.Show("Такая деталь существует");
        //            return itemExist = true;
        //        }
        //    }
        //    flag--;
        //    if (itemExist == false && flag == 0)
        //    {
        //        string temp = width;
        //        width = height;
        //        height = temp;
        //        goto L1;
        //    }
        //    return itemExist;
        //}


        public bool CheckForSimilarRows(string filePath, DataTable table, Dictionary<string, string> dictionary)
        {
            bool itemExist = false;

            MessageBox.Show(table.Rows.Count.ToString());
            
            foreach (var item in table.AsEnumerable())
            {
                if (item[dictionary["Обозначение"]].ToString() == filePath)
                {
                    MessageBox.Show("Такая деталь существует");
                    return itemExist = true;
                }
            }
            return itemExist;
        }


        /// <summary>
        /// Создает изделие
        /// </summary>
        /// <param name="elementType">0-сборка, 1-2 деталь</param>
        /// <param name="parentID"></param>
        private bool CreateProduct(int elementType, long parentID, string oboznachenie, string naimenovanie)
        {
            string objType = "";
            switch (elementType)
            {
                case 0:
                    objType = SystemGUIDs.objtypeAssemblyUnit;//изделие сборка
                    break;

                case 1:
                    objType = SystemGUIDs.objtypePart;//изделие деталь
                    break;

                case 2:
                    objType = SystemGUIDs.objtypePart;//изделие деталь
                    break;
            }
            using (SessionKeeper keeper = new SessionKeeper())
            {
                Int32 colType = MetaDataHelper.GetObjectTypeID(new Guid(objType));//тип создаваемого объекта
                IDBObjectCollection coll = keeper.Session.GetObjectCollection(colType);
                IDBObject newObj = coll.Create(colType);

                //берем аттрибуты объекта
                IDBAttribute atrDesignation = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeDesignation));//обозначение
                IDBAttribute atrName = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeName));//наименование
                atrDesignation.Value = oboznachenie;
                atrName.Value = naimenovanie;
                newObj.CommitCreation(true);
                IDBObject newObjChecked =  newObj.CheckOut();

                IDBRelationCollection colRel = keeper.Session.GetRelationCollection(1004);
                if (colRel != null)
                {
                    MessageBox.Show("parentID  " + parentID.ToString() + "          newObj.ID " + newObjChecked.ID.ToString());
                }
                else { MessageBox.Show("Failed create the relColl");}

                IDBRelation relation = colRel.Create(newObjChecked.ObjectID, parentID); // создаем связь между изделием и файлом, 
                if (relation != null)  { MessageBox.Show(relation.CreateDate.DayOfWeek.ToString()); }
                else { MessageBox.Show("Failed create the relation"); }

                newObjChecked.CheckIn();//

                MessageBox.Show("Изделие создано");
                return true;
            }
        }

        public static long GetSessionID()
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                return keeper.Session.ActingUserID;
            }
        }


        //перенести в ServiceConstants
        private enum ImbaseTableID :long
        {
            Spigot = 1746785,
            Roof = 1746878
        }
    }
}