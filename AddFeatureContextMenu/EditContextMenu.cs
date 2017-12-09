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
using Intermech.Client.Core.FormDesigner.Controls;

using Vents_PLM.Spigot;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;
using static Intermech.Interfaces.BlobProcCustomClass;

namespace AddFeatureContextMenu
{

    [ComVisible(true), Guid("F8358B23-3E51-453C-8EF8-9ABD22525214"), ProgId("IPS.EditContextMenu")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IOperatingWithContextMenu))]

    public class EditContextMenu : IPackage, IOperatingWithContextMenu
    {
        internal static IServiceProvider _serviceProvider;

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
                fac.AddCommandsProvider(Consts.CategoryObjectVersion, new CommandProvider());
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
                    airVentsMenuItem.Items.Add("Добавить елеектроный документ", AddDoc_EventBuilder);
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

           // GetIMBASETable();
           // ReadBlob();

            WriteBlob();
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
        private void WriteBlob()
        {
            using (SessionKeeper sk = new SessionKeeper())
            {
                //int typeO = MetaDataHelper;//тип создаваемого объекта

                int typeA = MetaDataHelper.GetAttributeTypeID(new Guid(SystemGUIDs.attributeFile));
                IDBObjectCollection coll = sk.Session.GetObjectCollection(1296);
                IDBObject newObj = coll.Create(1296);
                IDBAttribute fileAtr = newObj.GetAttributeByID(typeA);

                FileStream fileStream = null;

                string filePath = @"D:\Test\Проекты\12 - Вибровставка\12-20-450.SLDPRT";

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
                    BlobInformation blInfo = new BlobInformation(0, 0, DateTime.Now, "12-20-450.SLDPRT", ArcMethods.NotPacked, note: "JustJustJust_OneMore");
                    BlobProcWriter writer = new BlobProcWriter(fileAtr, (int)AttributableElements.Object, blInfo, fileStream, null, null);
                    MessageBox.Show("Im ready to write data");
                    writer.WriteData();


                    //IDBAttribute atr = newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeF_OBJECT_NAME));
                   // newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeF_OBJECT_NAME)).Value = "12-20-450";


                    newObj.CommitCreation(true);
                    MessageBox.Show("Data has been written" + newObj.ObjectID.ToString());
                }
                catch (Exception e) { MessageBox.Show(e.Message); }
            }
        }

        private void GetIMBASETable()
        {
            // Int32 type = MetaDataHelper.GetObjectTypeID("cad00215306c11d8b4e900304f19f545");//тип создаваемого объекта
            int type = MetaDataHelper.GetObjectTypeID(new Guid(SystemGUIDs.objtypePart));


            IDBObject newObj;
            using (SessionKeeper keeper = new SessionKeeper())
            {
                IDBObjectCollection coll = keeper.Session.GetObjectCollection(1296);//все таблицы IMBASE
                
                if (coll != null)
                {
                    MessageBox.Show("Some IMBASE tables are excist!");
                    newObj = coll.Create(1296);
                    DataTable dt = newObj.Attributes.GetAttributesDataTable();
                    List<IDBAttribute> atrList = newObj.Attributes.ToList();
                    int count = dt.Columns.Count - 1;

                    for(int i = 0; i < count; i++)
                    {
                        MessageBox.Show(dt.Columns[i].ColumnName.ToString());
                    }
                    int r = dt.Rows.Count - 1, c = dt.Columns.Count - 1;
                    foreach (var row in dt.Rows)
                    {
                        for (int k1 = 0; k1 < r; k1++)
                        {
                            for (int k2 = 0; k2 < c; k2++)
                            {
                                MessageBox.Show(dt.Rows[k1][k2].ToString());
                            }
                        }
                    }
                    //IDBAttribute attribName = newObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeDesignation));


                }
            }
        }


        private void CheckObjType()
        {
            using (SessionKeeper sk = new SessionKeeper())
            {
                IDBObjectCollection coll = sk.Session.GetObjectCollection(1296);
                
            }
        }
    }
}