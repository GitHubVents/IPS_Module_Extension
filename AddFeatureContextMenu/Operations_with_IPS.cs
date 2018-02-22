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
using System.Linq;

namespace AddFeatureContextMenu
{
    [ComVisible(true), Guid("F8358B23-3E51-453C-8EF8-9ABD22525214"), ProgId("IPS.EditContextMenu")]
    [ClassInterface(ClassInterfaceType.None)]

    public class Operations_with_IPS : IPackage
    {
        internal static IServiceProvider _serviceProvider;

        private IUserSession keeper = null;
        public IUserSession Session
        {
            get
            {
                if (keeper == null)
                {
                    keeper = new SessionKeeper().Session;
                }
                return keeper;
            }
        }
       
        UserControl userControl;



        delegate void DependsFLAG(ref IDBObject _object,  MyStruct myStruct);
        Operations_with_IPS op = new Operations_with_IPS();
        DependsFLAG dependsFlag;


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
            // provider.ContentCallback += new Intermech.Docking.GetContentCallback(WellKnownDockWindow.RestoreWindowCallback);

            // Получаем ссылку на сервис именованных значков
            INamedImageList _images = Operations_with_IPS._serviceProvider.GetService(typeof(INamedImageList)) as INamedImageList;

            // Получаем ссылку на службу панелей приложений "Навигатора"
            INavigationBar navBar = ServicesManager.GetService(typeof(INavigationBar)) as INavigationBar;

            //получаем обьект главного меню
            BarManager barManager = _serviceProvider.GetService(typeof(BarManager)) as BarManager;
            //добавляем в него новый Item - AirVentsCAD
            MenuBarItem airVentsMenuItem = barManager.MenuBar.AddMenuBar("AirVentsCAD");
            airVentsMenuItem.ForeColor = System.Drawing.Color.MediumAquamarine;

            //создаем пункты подменю
            airVentsMenuItem.Items.Add("Spigot", Spigot_EventBuilder);
            airVentsMenuItem.Items.Add("Roof", Roof_EventBuilder);
            airVentsMenuItem.Items.Add("Добавить записи в таблицу", AddDoc_EventBuilder);

            // Ищем панель "Приложения"
            IAppPane pane = navBar.FindPane("appPane") as IAppPane;
            /*
            // Панель найдена
            if (pane != null)
            {
                pane.Add("AAAAAAAAAAAAAAAAAAAAAAAAAAAH",
                    new EventHandler(WellKnownDockWindow.ShowWellKnownDockWindow),
                    _images.ImageIndex(WellKnownDockWindow.WindowImageName));
            }
            */

        }

        /////////отображение контролов
        private void Spigot_EventBuilder(object sender, EventArgs e)
        {
            userControl = new SpigotControl(Session.UserID);
            userControl.Dock = DockStyle.Fill;

            Form form = new Form();
            form.Size = new System.Drawing.Size(600, 500);
            form.Location = new System.Drawing.Point(500, 300);
            form.Controls.Add(userControl);
            userControl.Show();
            form.Show();
        }
        private void Roof_EventBuilder(object sender, EventArgs e)
        {
            // user = new RoofBilder();
            userControl.Dock = DockStyle.Fill;

            Form form = new Form();
            form.Size = new System.Drawing.Size(600, 500);
            form.Location = new System.Drawing.Point(500, 300);
            form.Controls.Add(userControl);
            userControl.Show();
            form.Show();
        }
        private void AddDoc_EventBuilder(object sender, EventArgs e)
        {
            //Test();
            //ExcelForm exel = new ExcelForm();
            //exel.Show();
        }


        //private void ReadBlob()
        //{
        //    using (SessionKeeper sk = new SessionKeeper())
        //    {
        //        IDBObject iDBObject = sk.Session.GetObject(1962223);
        //        IDBAttributable iDBAttributable = iDBObject as IDBAttributable;
        //        IDBAttribute[] iDBAttributeCollection = iDBAttributable.Attributes.GetAttributesByType(FieldTypes.ftFile);//получили аттрибут File

        //        if (iDBAttributeCollection != null)

        //            foreach (IDBAttribute fileAtr in iDBAttributeCollection)
        //            {
        //                if (fileAtr.ValuesCount < 2)
        //                {
        //                    MemoryStream m = new MemoryStream();
        //                    BlobProcReader reader = new BlobProcReader(fileAtr, (int)fileAtr.AsInteger, m, null, null);
        //                    MessageBox.Show("ElementID" + reader.ElementID.ToString() + "DataBlockSize =  " + reader.DataBlockSize.ToString() + "fileAtr.AsInteger  " + fileAtr.AsInteger.ToString() + "  fileAtr.AsDouble  " + fileAtr.AsDouble.ToString());
        //                    BlobInformation info = reader.BlobInformation;
        //                    reader.ReadData();
        //                    byte[] @byte = m.GetBuffer();
        //                    MessageBox.Show(@byte.Length.ToString());
        //                    //BlobProcReader reader = new BlobProcReader((long)info.BlobID, (int)AttributableElements.Object, fileAtr.AttributeID, 0, 0, 0);



        //                    //try
        //                    //{
        //                    //    IVaultFileReaderService serv = sk.Session.GetCustomService(typeof( IVaultFileReaderService)) as IVaultFileReaderService;
        //                    //    IVaultFileReader fileReader = serv.GetVaultFileReader(sk.Session.SessionID);
        //                    //    if (fileReader != null) MessageBox.Show("file reader exist       RealFileSize:" + info.RealFileSize.ToString() + Environment.NewLine + "PackedFileSize:" + info.PackedFileSize.ToString() + Environment.NewLine + iDBObject.ObjectID.ToString() + Environment.NewLine + fileAtr.AsInteger.ToString());

        //                    //    BlobInformation item = fileReader.OpenBlob(0, 1962223, 0, 8);
        //                    //    MessageBox.Show("Blob is open");

        //                    //    byte[] @byt = fileReader.ReadDataBlock((int)item.PackedFileSize);
        //                    //    MessageBox.Show(@byt.GetLength(@byt.Rank).ToString());
        //                    //}
        //                    //catch { MessageBox.Show("No service found"); }
        //                    //return;
        //                }
        //            }
        //    }
        //}

        public DataTable GetIMBASETable(long tableID, out DataSet ds, out Dictionary<string, string> dictionary)
        {
            using (SessionKeeper keeper = new SessionKeeper())
            {
                ds = TableLoadHelper.GetTables(keeper.Session, tableID, true);
                DataTable tb1 = ds.Tables[0];
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
                        MessageBox.Show("There is no such attribute:  " + colName.ToString() + "  " + tableID.ToString());
                    }
                }
                return tb1;
            }
        }

        public bool CheckForSimilarRows(string fileName, DataTable table, Dictionary<string, string> dictionary)
        {
            bool itemExist = false;

            MessageBox.Show(table.Rows.Count.ToString());

            foreach (var item in table.AsEnumerable())
            {
                if (item[dictionary["Обозначение"]].ToString() == fileName)
                {
                    MessageBox.Show("Такая деталь существует " + fileName);
                    return itemExist = true;
                }
            }
            return itemExist;
        }

        
        public void MakeRelationsBtwnDocs(List<long> children, long parentID)
        {
            IDBRelationCollection coll = Session.GetRelationCollection(1014);

            IDBObject objectParent = Session.GetObject(parentID);
            IDBObject objectParentChecked = objectParent.CheckOut(true); //  asm


            IDBObject newChild;
            IDBObject newChildChecked;

            foreach (var childID in children)
            {
                newChild = Session.GetObject(childID);
                newChildChecked = newChild.CheckOut(true); //  prt

                IDBRelation relation = coll.Create(objectParentChecked.ObjectID, newChildChecked.ObjectID); // создаем связь между изделием и документом
                
                MessageBox.Show(relation.CreateDate.DayOfWeek.ToString());
            }
        }

        //public void Test()
        //{
        //    string pathPart = @"C:\Users\kb81\Desktop\Новая папка\1850-0051.01.012 - Вставка 1.sldprt";
        //    string pathASM = @"C:\Users\kb81\Desktop\Новая папка\1850-0051.01.002-02 - Корпус сонотрода (второй ремонт).SLDASM";


        //    var a = GetIMBASETable((int)ImbaseTableID.Spigot, out ds, out dictionary);

        //    WriteIntoIMBASE_Spigot_Table(a, pathPart, "part", "name", "789", "987", "20", 1); // prt
        //    WriteIntoIMBASE_Spigot_Table(a, pathASM, "asm", "name", "789", "987", "20", 0); //asm


        //   // MakeRelationsBtwnDocs(new List<long> { idpart }, idparent);
        //}
        


        //перенести в ServiceConstants
        private enum ImbaseTableID :long
        {
            Spigot = 1746785,
            Roof = 1746878
        }

        public enum IPSObjectTypes
        {
            //документы
            DocSldAssembly= 1361,
            DocSldPart = 1296,
            DocSldDrw = 1324,
            DocSldDrwAssembly = 1285,

            //изделия
            ProductPart = 1052,
            ProductASM = 1074,
            ProductOther = 1138,
            ProductStandart = 1105,
            Empty = -100,
            //Материал
            Materials = 1128
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="DocObj">Передаем null, инициализируеться в методе</param>
        /// <param name="_object"></param>
        private void CreatDoc(ref IDBObject DocObj, MyStruct _object)
        {
            IDBObjectCollection coll = Session.GetObjectCollection((int)_object.IPsDocType);
            DocObj = coll.Create((int)_object.IPsDocType);

            // добавление создаваемого обьекта в Документы/Конструкторские/Электронные модели
            IDBAttribute atrName = DocObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeName));//наименование документа
            IDBAttribute atrDesignation = DocObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeDesignation));//обозначение документа

            atrName.Value = _object.Name;
            atrDesignation.Value = _object.Designition;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="DocObj">При записи в БД нового обьекта, DocObj это новый обьект созданный предварительным вызовом CreateDoc(); при добавлении файла к существующей сборке, DocObj - обьект сборки"</param>
        /// <param name="filePath"></param>
        private void Blob(ref IDBObject DocObj, MyStruct _object)
        {
            if (DocObj == null)// пишем не в новый документ, а в сборку
            {
                BlobIntoASM(ref DocObj, _object);
            }

            int attrFile = MetaDataHelper.GetAttributeTypeID(new Guid(SystemGUIDs.attributeFile));// атрибут "Файл"
            IDBAttribute fileAtr = DocObj.GetAttributeByID(attrFile);

            using (var ms = new MemoryStream(File.ReadAllBytes(_object.Path)))
            {
                BlobInformation blInfo = new BlobInformation(0, 0, DateTime.Now, _object.Path, ArcMethods.NotPacked, null);
                BlobProcWriter writer = new BlobProcWriter(fileAtr, (int)AttributableElements.Object, blInfo, ms, null, null);
                writer.WriteData();
            }
        }

        private void BlobIntoASM(ref IDBObject DocObj, MyStruct _object)
        {
            /////////////// DocObj - обьект сборки
        }

        /// <summary>
        /// Создает изделие
        /// </summary>
        /// <param name="productType"></param>
        /// <param name="parentID">ID документа на который создаеться изделие</param>
        private void CreateProduct(ref IDBObject DocObj, MyStruct _object)
        {
            IDBObjectCollection coll = Session.GetObjectCollection((int)_object.IPsProductType);
            IDBObject newObj = coll.Create((int)_object.IPsProductType);

            //берем аттрибуты объекта
            IDBAttribute atrDesignation = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeDesignation));//обозначение
            IDBAttribute atrName = newObj.GetAttributeByGuid(new Guid(SystemGUIDs.attributeName));//наименование

            atrDesignation.Value = _object.Designition;
            atrName.Value = _object.Name;
            newObj.CommitCreation(true);

            IDBObject newObjChecked = newObj.CheckOut();

            IDBRelationCollection colRel = Session.GetRelationCollection(1004);
            IDBRelation relation = colRel.Create(newObjChecked.ObjectID, DocObj.ID); // создаем связь между изделием и докумиентом, 
            if (relation == null) { MessageBox.Show("Failed create the relation of type " + colRel.RelationTypeID.ToString()); }

            //newObjChecked.CheckIn();//
        }


        private DependsFLAG DependsOnFLAG(int flag)
        {
            switch (flag)
            {
                case 0:

                    dependsFlag = op.CreatDoc;
                    dependsFlag += op.Blob;
                    dependsFlag += op.CreateProduct;

                    break;
                case 1:

                    break;
                case 2:

                    break;
                case 3:

                    break;
                case 4:

                    break;
                case 5:
                    dependsFlag += op.Blob;
                    dependsFlag += op.CreateProduct;
                    break;
            }
            return dependsFlag;
        }

        //вызывать после построения модели в библиотеке SolidWorksLibrary
        public void InvokeMyDelegats(List<MyStruct> objectsToWrite)
        {

            string mainASMName = objectsToWrite.Where(x => x.RefAsmName.Equals(String.Empty)).Select(x => x.Name).First();

            BuildTreeView(objectsToWrite, mainASMName);
        }

        private void BuildTreeView(List<MyStruct> objectsToWrite, string name)
        {

            //cписок деталей/чертежей/сборок, которые ссылаються на name (имя сборки более высокого уровня)
            List<MyStruct> list = objectsToWrite.Where(
                                                        x =>
                                                        (x.RefAsmName == name)// && x.IPsDocType.Equals(IPSObjectTypes.DocSldAssembly))
                                                      ).ToList();

            if (list.Count == 0)
            {
                op.WriteIntoIPS(objectsToWrite);
            }
            else
            {
                foreach (var item in list.Where(x=>x.RefAsmName.Equals(IPSObjectTypes.DocSldAssembly)))
                {
                    BuildTreeView(objectsToWrite, item.Name);
                }
            }
        }

        private void WriteIntoIPS(List<MyStruct> objectsToWrite)
        {
            IDBObject obj = null;
            foreach (var item in objectsToWrite)
            {
                dependsFlag = DependsOnFLAG(item.Flag);
                dependsFlag(ref obj, item);
            }
        }
    }
}