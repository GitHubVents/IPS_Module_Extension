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
using System.Linq;
using ControlsLibrary.MountingFrame;
using ControlsLibrary.Spigot;

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
        private long sessionID;
        public long SessionID
        {
            get
            {
                if (sessionID == 0)
                {
                    sessionID = Session.SessionID;
                }
                return sessionID;
            }
        }


        UserControl userControl;
        Dictionary<string, IDBObject> createdDocs = new Dictionary<string, IDBObject>();

        delegate void DependsFLAG(ref IDBObject _object,  MyStruct myStruct);
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
            airVentsMenuItem.Items.Add("Mounting frame", Frame_EventBuilder);
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
            form.ShowDialog();
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

        private void Frame_EventBuilder(object sender, EventArgs e)
        {
            userControl = new MountingFrame();
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
            ExcelForm exel = new ExcelForm();
            exel.Show();

            //InvokeMyDelegats(Test());
            //WriteIntoIPS(Test());
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

        
        public void MakeRelationsBtwnDocs(Dictionary<string, IDBObject> children, IDBObject parent)
        {
            IDBRelationCollection coll = Session.GetRelationCollection(1014);
            
            IDBObject objectParentChecked = parent.CheckOut(true); //  asm
            IDBObject newChildChecked;

            foreach (var child in children)
            {

                newChildChecked = child.Value.CheckOut(true); //  prt
                IDBRelation relation = coll.Create(objectParentChecked.ObjectID, newChildChecked.ObjectID); // создаем связь между изделием и документом
            }
        }

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
            DocObj = null;
            IDBObjectCollection coll = Session.GetObjectCollection((int)_object.IPsDocType);
            DocObj = coll.Create((int)_object.IPsDocType);
            
            // добавление создаваемого обьекта в Документы/Конструкторские/Электронные модели
            IDBAttribute atrName = DocObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeName));//наименование документа
            IDBAttribute atrDesignation = DocObj.Attributes.FindByGUID(new Guid(SystemGUIDs.attributeDesignation));//обозначение документа

            atrName.Value = _object.Name;
            atrDesignation.Value = _object.Designition;

            var res = createdDocs.Where(x => x.Key.Equals(_object.Designition)).ToList();
            if(res.Count == 0)
            {
                createdDocs.Add(_object.Designition, DocObj);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="DocObj">При записи в БД нового обьекта, DocObj это новый обьект созданный предварительным вызовом CreateDoc(); при добавлении файла к существующей сборке, DocObj - обьект сборки"</param>
        /// <param name="filePath"></param>
        private void Blob(ref IDBObject DocObj, MyStruct _object)
        {
            if (_object.Flag == 5)// пишем Blob не в новый документ, а в сборку
            {
                DocObj = createdDocs.Where(x => x.Key.Equals(_object.RefAsmName)).Select(y=>y.Value).First();
            }
            
            int attrFile = MetaDataHelper.GetAttributeTypeID(new Guid(SystemGUIDs.attributeFile));// атрибут "Файл"
            IDBAttribute fileAtr = DocObj.GetAttributeByID(attrFile);

            if (fileAtr.Values.Count() >= 1)
            {
                fileAtr.AddValue(_object.Path);
            }
            using (var ms = new MemoryStream(File.ReadAllBytes(_object.Path)))
            {
                BlobInformation blInfo = new BlobInformation(0, 0, DateTime.Now, _object.Path, ArcMethods.NotPacked, null);
                BlobProcWriter writer = new BlobProcWriter(fileAtr, (int)AttributableElements.Object, blInfo, ms, null, null);

                writer.WriteData();
            }
        }


        /// <summary>
        /// Создает изделие
        /// </summary>
        /// <param name="productType"></param>
        /// <param name="parentID">ID документа на который создаеться изделие</param>//prt
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
        }
        


        private DependsFLAG DependsOnFLAG(int flag)
        {
            dependsFlag = null;
            
            switch (flag)
            {
                case 0:
                    dependsFlag += CreatDoc;
                    dependsFlag += Blob;
                    dependsFlag += CreateProduct;
                    break;
                case 1:
                    dependsFlag += Blob;
                    dependsFlag += CreateProduct;
                    break;
                case 2:
                    dependsFlag = CreatDoc;
                    dependsFlag += Blob;//prt
                    break;
                case 3:

                    break;
                case 4:

                    break;
                case 5:
                    dependsFlag += Blob;//asm
                    break;
            }
            return dependsFlag;
        }

        //вызывать после построения модели в библиотеке SolidWorksLibrary
        public void InvokeMyDelegats(List<MyStruct> objectsToWrite)
        {
            string mainASMDesignition = objectsToWrite.Where(x => x.RefAsmName.Equals(String.Empty)).Select(x => x.Designition).First();

            BuildTreeView(objectsToWrite, mainASMDesignition, objectsToWrite);

            foreach (var item in createdDocs)
            {
                item.Value.CommitCreation(true);
            }
        }


        //asmDesignition - имя сборки находящейся на одном уровне с objectsOfTheSameLevel
        private void BuildTreeView(List<MyStruct> allObjects, string asmDesignition, List<MyStruct> objectsOfTheSameLevel)
        {
            MyStruct asm;
            Predicate<MyStruct> getAsm = delegate (MyStruct t) { return t.Designition == asmDesignition; };
            Predicate<MyStruct> getHigherReferencedAsm = delegate (MyStruct t) { return t.RefAsmName == asmDesignition; };

            //cписок деталей/чертежей/сборок нижнего уровня, которые ссылаються на name (имя сборки более высокого уровня)
            List<MyStruct> list = allObjects.FindAll(getHigherReferencedAsm);
            
            // сборки которые есть на нижнем уровне
            List<MyStruct> subList = list?.Where(x => x.IPsDocType.Equals(IPSObjectTypes.DocSldAssembly)).ToList();

            
            if (list.Count == 0) // в list только одна сборка, на которую не ссылаються другие обьекты, сборка нижнего уровня
            {
                asm = allObjects.Find(getAsm);
                WriteIntoIPS(new List<MyStruct> { asm }); // записали сборку
                //связь со сборкой высшего уровня???
            }
            
            if (subList.Count == 0)//если сборок нет, заливаем в IPS
            {
                asm = objectsOfTheSameLevel.Find(getAsm);

                IDBObject obj = null;
                CreatDoc(ref obj, asm); // создали документ сборки

                list.Remove(asm);
                WriteIntoIPS(list); // а теперь обьекты, которые на нее ссылаються

                MakeRelationsBtwnDocs(createdDocs, createdDocs.Where(x=>x.Key.Equals(asm.Designition)).First().Value);
            }
            else
            {
                foreach (var item in subList)
                {
                    BuildTreeView(allObjects, item.Designition, list);
                }
            }
        }

        private void WriteIntoIPS(List<MyStruct> objectsToWrite)
        {
            IDBObject obj = null;
            foreach (var item in objectsToWrite)
            {
                dependsFlag = DependsOnFLAG(item.Flag);
                dependsFlag?.Invoke(ref obj, item);
            }
        }

        private List<MyStruct> Test()
        {
            List<MyStruct> structura = new List<MyStruct>();
            MyStruct newS = new MyStruct(5,  "Design", "12-03-837-914", @"D:\IPS Vault\4\Workspace\Проекты\Blauberg\12 - Вибровставка\12-03-837-914.SLDPRT", "12-20-837-914", IPSObjectTypes.DocSldPart, IPSObjectTypes.ProductPart);
            structura.Add(newS);
           
            newS = new MyStruct(0, "Design","12-03-837",  @"D:\IPS Vault\4\Workspace\Проекты\Blauberg\12 - Вибровставка\12-20-837.SLDPRT", "12-20-837-914", IPSObjectTypes.DocSldPart, IPSObjectTypes.ProductPart);
            structura.Add(newS);

            newS = new MyStruct(0, "Design", "12-20-914", @"D:\IPS Vault\4\Workspace\Проекты\Blauberg\12 - Вибровставка\12-20-914.SLDPRT", "12-20-837-914", IPSObjectTypes.DocSldPart, IPSObjectTypes.ProductPart);
            structura.Add(newS);
         
            newS = new MyStruct(0, "ASM", "12-20-837-914", @"D:\IPS Vault\4\Workspace\Проекты\Blauberg\12 - Вибровставка\12-20-837-914.SLDASM", String.Empty, IPSObjectTypes.DocSldAssembly, IPSObjectTypes.ProductASM);
            structura.Add(newS);

            return structura;
        }
    }
}