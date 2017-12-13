using System;
using Intermech;
using Intermech.Bars;
using Intermech.Interfaces;
using Intermech.Interfaces.Plugins;
using Intermech.Interfaces.Client;
using Intermech.Navigator.ContextMenu;
using Intermech.Navigator.Interfaces;
using Intermech.NavBars;
using System.Runtime.InteropServices;
using Intermech.Runtime.ComInterop.LocalServer;
using Vents_PLM.Spigot;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;
using Intermech.Imbase;

namespace AddFeatureContextMenu
{

    [ComVisible(true), Guid("F8358B23-3E51-453C-8EF8-9ABD22525214"), ProgId("IPS.EditContextMenu")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IOperatingWithContextMenu))]

    public class EditContextMenu : IPackage, IOperatingWithContextMenu
    {
        internal static IServiceProvider _serviceProvider;
        string filePath = @"D:\Test\Проекты\12 - Вибровставка\12-03-453-357.SLDPRT";

        public void Load(IServiceProvider serviceProvider)
        {
            if (ComHost.Configuration.ComSupportActive)
            {
                ComHost.ActivateClassFactory(typeof(EditContextMenu));
            }

            _serviceProvider = serviceProvider;

            Test2();
        }
        public void Unload()
        {
            if (ComHost.Configuration.ComSupportActive)
            {
                ComHost.DeactivateClassFactory(typeof(EditContextMenu));
            }
        }
        string IPackage.Name
        {
            get
            {
                return "AddFeatureContextMenu" + DateTime.Now.DayOfWeek.Name() + DateTime.UtcNow.ToString();
            }
        }

        public EditContextMenu()
        {
            //Test();
        }

        public static IUserSession session;
        public IUserSession Connect(string loginName, string password)
        {
            string serverURL = "tcp://PDMSRV:8008/IntermechRemoting/Server.rem";

            // Подключаемся к серверу и получаем сессию
            IMServer server = (IMServer)Activator.GetObject(typeof(IMServer), serverURL);
            IUserSession session = server.CreateSession();
            // Получаем смещение для текущего часового пояса и вызываем у сессии функцию авторизации
            DateTime now = DateTime.Now;
            TimeSpan ts = now - now.ToUniversalTime();
            session.Login(loginName, password, Environment.MachineName, ts, 0);
            return session;
        }

        public void Test()
        {
            IFactory fac = ServicesManager.GetService(typeof(IFactory)) as IFactory;
            INamedImageList imgList = ServicesManager.GetService(typeof(INamedImageList)) as INamedImageList;
            MenuTemplateNode createNode = fac.ContextMenuTemplate["Create"];

            if (createNode != null)
            {
                createNode.Nodes.Add(new MenuTemplateNode("Создать", "Создать копию объекта1", imgList.ImageIndex("imgCreate"), 20, Int32.MaxValue));
                createNode.Nodes.Add(new MenuTemplateNode("CreateCopy", "Создать копию объекта2", imgList.ImageIndex("imgSamples.CopyObject"), 20, Int32.MaxValue));

                // Регистрируем провайдер команд контекстного меню для категории "Версии объектов"
                fac.AddCommandsProvider(Intermech.Consts.CategoryObjectVersion, new CommandProvider());
            }
        }
        public void Test2()
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

        private void Spigot_EventBuilder(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("I'm building spigot");

            System.Windows.Forms.UserControl user = new SpigotControl();
            user.Dock = System.Windows.Forms.DockStyle.Fill;

            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            form.Size = new System.Drawing.Size(600, 500);
            form.Location = new System.Drawing.Point(500, 300);
            form.Controls.Add(user);
            user.Show();
            form.Show();
        }
        private void Roof_EventBuilder(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("I'm building roof");
        }

        private void AddDoc_EventBuilder(object sender, EventArgs e)
        {
            //WriteIntoIMBASE_Roof_Table("1500", "640", "2");
            WriteIntoIMBASE_Spigot_Table("453", "357", "20", "1");
            //WriteBlob(filePath, 1296, "Обозначение", "Наименование");
            return;


            Int32 assemblyUnitType = MetaDataHelper.GetObjectTypeID(SystemGUIDs.objtypeAssemblyUnit);//тип создаваемого объекта
            IDBObject newObj;
            System.Windows.Forms.MessageBox.Show(assemblyUnitType.ToString());
            using (SessionKeeper keeper = new SessionKeeper())
            {
                IDBObjectCollection coll = keeper.Session.GetObjectCollection(assemblyUnitType);
                newObj = coll.Create(assemblyUnitType);
                IDBAttribute attribName = newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeDesignation));
                if (attribName != null)
                {
                    attribName.Value = "ILikeIt";
                }
                DataTable dt = newObj.Attributes.GetAttributesDataTable();
                if (dt != null)
                {
                    foreach (var item in dt.Columns)
                    {
                        MessageBox.Show(item.ToString());
                    }
                    MessageBox.Show(dt.Rows.Count.ToString());
                }
                newObj.CommitCreation(true);
            }
            // Получаем ссылку на службу ведомлений
            INotificationService svc = ServicesManager.GetService(typeof(INotificationService)) as INotificationService;

            // Проверим её наличие
            if (svc == null) return;

            // Создаём аргументы события "Создан объект"
            DBObjectsEventArgs oe = new DBObjectsEventArgs(NotificationEventNames.ObjectsCreated, newObj.ObjectID);

            // Уведомляем клиентское приложение о том, что был создан новый объект
            svc.FireEvent(null, oe);

            // Создаём дескриптор нового объекта
            Intermech.Navigator.DBObjects.Descriptor descriptor = new Intermech.Navigator.DBObjects.Descriptor(newObj.ObjectID);

            // Пробуем открыть вновь созданный объект в новом окне "Навигатора"
            Intermech.Navigator.Utils.OpenNewWindow(descriptor, null);
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
        private long WriteBlob(string filePath, int objType, string attrName, string attrDesignation)
        {
            long id = 0;
            using (SessionKeeper sk = new SessionKeeper())
            {
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
                    BlobInformation blInfo = new BlobInformation(0, 0, DateTime.Now, "12-03-453-357.SLDPRT", ArcMethods.NotPacked, note: "JustJustJust_OneMore");
                    BlobProcWriter writer = new BlobProcWriter(fileAtr, (int)AttributableElements.Object, blInfo, fileStream, null, null);
                    writer.WriteData();
                    try
                    {
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
                    MessageBox.Show("Data has been written   " + newObj.ObjectID.ToString());
                    return id;
                }
                catch (Exception e) { MessageBox.Show(e.Message); }
            }
            return id;
        }



        private void WriteIntoIMBASE_Roof_Table(string width, string height, string roofType)//ROOF
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds = TableLoadHelper.GetTables(keeper.Session, 1746878, true);

                DataTable tb1 = ds.Tables[0];//таблица с данными
                DataTable tb2 = ds.Tables[1];
                DataRow row1 = tb1.NewRow();
                string colName = "";
                Dictionary<string, string> dictionary = new Dictionary<string, string>();

                for (int i = 0;i< tb1.Columns.Count; i++)
                {
                    MessageBox.Show(tb1.Columns[i].ToString());
                    switch (i)
                    {
                        case 2:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 3:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 4:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                    }
                }

                if (CheckForSimilarRows(width, height, roofType, tb1, dictionary))
                {
                    return;
                }
                else
                {
                    row1[Intermech.Imbase.Consts.F_GUID] = new Guid();
                    row1[""+dictionary["Ширина"]+""] = width;
                    row1[""+dictionary["Высота"]+""] = height;
                    row1[""+dictionary["Тип крыши"]+""] = roofType;

                    tb1.Rows.Add(row1);
                    tb1.AcceptChanges();
                    TableLoadHelper.StoreData(keeper.Session, 1746878, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
                }                             
            }
        }
        private void WriteIntoIMBASE_Spigot_Table(string width, string height, string flangeType, string elementType)//SPIGOT
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds = TableLoadHelper.GetTables(keeper.Session, 1746785, true);

                DataTable tb1 = ds.Tables[0];//таблица с данными
                DataTable tb2 = ds.Tables[1];
                DataRow row1 = tb1.NewRow();
                string colName = "";
                Dictionary<string, string> dictionary = new Dictionary<string, string>();

                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    switch (i)
                    {
                        case 2:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 3:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 4:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 5:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                    }
                }
                
                if (CheckForSimilarRows(width, height, flangeType, elementType, tb1, dictionary))
                { return;}
                else
                {
                    long id = WriteBlob(filePath, 1296, "Обозначение", "Наименование");
                    //////////////To call Build method
                    if(CreateProduct(1, id))
                    {
                        row1[Intermech.Imbase.Consts.F_GUID] = new Guid();
                        row1["" + dictionary["Ширина"] + ""] = width;
                        row1["" + dictionary["Высота"] + ""] = height;
                        row1["" + dictionary["Тип фланца"] + ""] = flangeType;
                        row1["" + dictionary["Element Type"] + ""] = elementType;

                        tb1.Rows.Add(row1);
                        tb1.AcceptChanges();
                        TableLoadHelper.StoreData(keeper.Session, 1746785, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
                    }
                }
            }
        }


        private bool CheckForSimilarRows(string width, string height, string roofType, DataTable table, Dictionary<string, string> dictionary)
        {
            bool itemExist = false;
            int flag = 1;
            L1:
                foreach (var item in table.AsEnumerable())
                {
                    if (item[dictionary["Ширина"]].ToString() == width 
                        && item[dictionary["Высота"]].ToString() == height
                        && item[dictionary["Тип крыши"]].ToString() == roofType)
                    {
                        MessageBox.Show("Такая деталь существует  " + item[dictionary["Ширина"]].ToString());
                        return itemExist = true;
                    }
                }
            flag--;
            if (itemExist == false && flag == 0)
            {
                MessageBox.Show("else");
                string temp = width;
                width = height;
                height = temp;
                goto L1;
            }
            MessageBox.Show(itemExist.ToString());
            return itemExist;
        }
        private bool CheckForSimilarRows(string width, string height, string flangeType, string elementType, DataTable table, Dictionary<string, string> dictionary)
        {
            bool itemExist = false;
            int flag = 1;
            L1:
            foreach (var item in table.AsEnumerable())
            {
                if (item[dictionary["Ширина"]].ToString() == width
                    && item[dictionary["Высота"]].ToString() == height
                    && item[dictionary["Тип фланца"]].ToString() == flangeType)
                {
                    MessageBox.Show("Ширина  " + item[dictionary["Ширина"]].ToString());
                    return itemExist = true;
                }
            }
            flag--;
            if (itemExist == false && flag == 0)
            {
                MessageBox.Show("else");
                string temp = width;
                width = height;
                height = temp;
                goto L1;
            }
            MessageBox.Show(itemExist.ToString());
            return itemExist;
        }


        private void FillSpigot_Table()
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                DataSet ds = TableLoadHelper.GetTables(keeper.Session, 1746785, true);

                DataTable tb1 = ds.Tables[0];//таблица с данными
                DataTable tb2 = ds.Tables[1];
                DataRow row1 = tb1.NewRow();
                string colName = "";
                string fullDesignation;
                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                   
                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    switch (i)
                    {
                        case 2:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 3:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 4:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 5:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                        case 6:
                            colName = MetaDataHelper.GetAttributeTypeName(new Guid(tb1.Columns[i].ToString()));
                            dictionary.Add(colName, tb1.Columns[i].ToString());
                            break;
                    }
                }

                for (int i = 0; i < tb1.Rows.Count; i++)
                {
                    string value1 = "0", value2 = "0";
                    fullDesignation = tb1.Rows[i][tb1.Columns[6].ToString()].ToString();

                    string[] parsedDesignation = fullDesignation.Split('-');

                    if (parsedDesignation.Length == 3)
                    {
                        if (parsedDesignation[1].Length > 2)
                        {
                            value1 = parsedDesignation[1];
                        }
                        value2 = parsedDesignation[2];
                    }
                    else if (parsedDesignation.Length == 4)
                    {
                        if (parsedDesignation[2].Length > 2)
                        {
                            value1 = parsedDesignation[2];
                        }
                        value2 = parsedDesignation[3];
                    }

                    tb1.Rows[i].SetField("" + dictionary["Ширина"] + "", value1);
                    tb1.Rows[i].SetField("" + dictionary["Высота"] + "", value2);
                    tb1.AcceptChanges();
                    TableLoadHelper.StoreData(keeper.Session, 1746785, ds, keeper.Session.GetCustomService(typeof(Intermech.Interfaces.Imbase.ITablesIndexer)) as Intermech.Interfaces.Imbase.ITablesIndexer);
                }
            }
        }

        /// <summary>
        /// Создает изделие
        /// </summary>
        /// <param name="elementType">0-сборка, 1-2 деталь</param>
        /// <param name="parentID"></param>
        private bool CreateProduct(int elementType, long parentID)
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

                IDBAttribute atrDesignation = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeDesignation));//обозначение
                IDBAttribute atrName = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeName));//наименование
                atrDesignation.Value = "НОВОЕ ИЗДЕЛИЕ1243";
                atrName.Value = "ИзДеЛиЕ+1243";
                newObj.CommitCreation(true);
                IDBObject newObjChecked =  newObj.CheckOut();

                IDBRelationCollection colRel = keeper.Session.GetRelationCollection(1004);
                if (colRel != null)
                {
                    MessageBox.Show("parentID  " + parentID.ToString() + "          newObj.ID " + newObjChecked.ID.ToString());
                } else { MessageBox.Show("Failed create the relColl");}

                IDBRelation relation = colRel.Create(newObjChecked.ObjectID, parentID);
                if (relation != null)  { MessageBox.Show(relation.CreateDate.DayOfWeek.ToString()); }
                else { MessageBox.Show("Failed create the relation"); }

                newObjChecked.CheckIn();

                MessageBox.Show("Изделие создано");
                return true;
            }
        }
    }
}