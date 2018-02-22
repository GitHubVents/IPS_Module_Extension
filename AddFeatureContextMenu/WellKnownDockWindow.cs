using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

using Intermech;
using Intermech.Docking;
using Intermech.Interfaces;
using Intermech.Interfaces.Client;
using Intermech.Client.Core;
using System.Xml;
using System.IO;
using Intermech.Navigator.Interfaces;
using Intermech.Navigator.Controls;
using Intermech.Navigator.DBObjects;
using System.ComponentModel.Design;

namespace AddFeatureContextMenu
{
    public class WellKnownDockWindow : Intermech.Docking.DockControl, IIODestination, ITreeListColumns
    {
        #region Константы

        /// <summary>
        /// Заголовок окна
        /// </summary>
        public const string WindowText = "ААААААААААААААААААААААААААААААН";

        /// <summary>
        /// Название изображения
        /// </summary>
        public const string WindowImageName = "imgSamples.SampleWindow";

        /// <summary>
        /// Название окна
        /// </summary>
        public const string WindowName = "Каталоги и справочники IMBAS";

        /// <summary>
        /// Guid, характеризующий экземпляры данного класса
        /// </summary>
        private static readonly Guid _persistStateGuid = new Guid("{ED9F24EA-6319-42B7-902C-329A4D1E06EE}");

        #endregion

        #region Поля класса

        /// <summary> 
        /// Контейнер компонентов
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Выполнена ли активация окна
        /// </summary>
        protected bool _activated;

        /// <summary>
        /// Выполнена ли загрузка данных в окно
        /// </summary>
        protected bool _loaded;

        /// <summary>
        /// Сервис службы уведомлений
        /// </summary>
        protected NotificationService _notificationService;

        /// <summary>
        /// Сервис именованных изображений
        /// </summary>
        private INamedImageList _images;

        /// <summary>
        /// Контейнер сервисов для дерева "Навигатора"
        /// </summary>
        protected AdvancedServiceContainer _servicesTree;

        /// <summary>
        /// Контейнер сервисов для менеджера закладок
        /// </summary>
        protected AdvancedServiceContainer _servicesPages;

        /// <summary>
        /// Коллекции команд по умолчанию
        /// </summary>
        private IDefaultCommands4ObjTypes _defaultCommands4ObjTypes;

        /// <summary>
        /// Диспетчер событий
        /// </summary>
        private IIODispatcher _IODispatcher = new IODispatcher();

        /// <summary>
        /// Системный диспетчер событий
        /// </summary>
        private IIODispatcher _systemDispatcher;

        /// <summary>
        /// Сервис службы "горячих клавиш" и связанных с ними команд
        /// </summary>
        private IHotKeysManager _hotKeysManager;

        /// <summary>
        /// Контейнер
        /// </summary>
        private SplitContainer splitContainer;
        private PageViewsManager pageViewsManager;

        /// <summary>
        /// Дерево "Навигатора"
        /// </summary>
        private Intermech.Navigator.Controls.NavigatorTreeView navigatorTreeView;

        #endregion

        #region Конструктор

        /// <summary>
        /// Создать экземпляр окна
        /// </summary>
        public WellKnownDockWindow()
        {
            // Инициализируем сервисы
            InitializeComponent();

            // Перекрываем базовый Guid
            base.Guid = _persistStateGuid;

            // Инициализируем сервисы
            if (!this.DesignMode)
                InitializeServices();

            // Назначаем заголовок
            this.Text = WellKnownDockWindow.WindowText;

            // Ищем индекс изображения для окна
            Int32 idx = (_images != null) ? _images.ImageIndex(WindowImageName) : -1;

            // Если изображение есть, разрешаем закладке его отображать
            this.ShowImageInDocumentTab = (idx >= 0);

            // И назначаем индекс изображения закладке
            this.TabImageIndex = idx;

            // Активация формы ещё не выполнена
            _activated = false;
        }

        #endregion

        #region Деструктор

        /// <summary> 
        /// Освободить занимаемые формой ресурсы
        /// </summary>
        /// <param name="disposing">true, если требуется освободить управляемые ресурсы</param>
        protected override void Dispose(bool disposing)
        {
            // Получаем ссылку на службу именованных окон "Навигатора"
            IWellKnownNavigators wkn = ServicesManager.GetService(typeof(IWellKnownNavigators)) as IWellKnownNavigators;

            // Удалим окно из списка именованных окон "Навигатора"
            if (disposing && wkn != null)
            {
                wkn.Unregister(this);
            }

            // Удаляем компоненты
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            // Базовый деструктор
            base.Dispose(disposing);
        }

        #endregion

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.navigatorTreeView = new Intermech.Navigator.Controls.NavigatorTreeView();
            this.pageViewsManager = new Intermech.Navigator.Controls.PageViewsManager();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.navigatorTreeView)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer
            // 
            this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer.Location = new System.Drawing.Point(0, 0);
            this.splitContainer.Name = "splitContainer";
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.navigatorTreeView);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.pageViewsManager);
            this.splitContainer.Size = new System.Drawing.Size(620, 300);
            this.splitContainer.SplitterDistance = 255;
            this.splitContainer.TabIndex = 0;
            // 
            // navigatorTreeView
            // 
            this.navigatorTreeView.AllowDrop = true;
            this.navigatorTreeView.AllowMultiSelect = false;
            this.navigatorTreeView.AllowUserPinnedColumns = false;
            this.navigatorTreeView.BackgroundImageMode = Infralution.Controls.ImageDrawMode.Tile;
            this.navigatorTreeView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.navigatorTreeView.DisableKeyDownEvents = true;
            this.navigatorTreeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.navigatorTreeView.HeaderStyle.HorzAlignment = System.Drawing.StringAlignment.Near;
            this.navigatorTreeView.ImageList = null;
            this.navigatorTreeView.LineStyle = Infralution.Controls.VirtualTree.LineStyle.Dot;
            this.navigatorTreeView.Location = new System.Drawing.Point(0, 0);
            this.navigatorTreeView.Name = "navigatorTreeView";
            this.navigatorTreeView.RowEvenStyle.WordWrap = false;
            this.navigatorTreeView.RowOddStyle.WordWrap = false;
            this.navigatorTreeView.RowSelectedStyle.BackColor = System.Drawing.SystemColors.Highlight;
            this.navigatorTreeView.RowSelectedStyle.WordWrap = false;
            this.navigatorTreeView.RowSelectedUnfocusedStyle.BackColor = System.Drawing.SystemColors.Highlight;
            this.navigatorTreeView.RowStyle.BorderColor = System.Drawing.SystemColors.Control;
            this.navigatorTreeView.RowStyle.BorderStyle = System.Windows.Forms.Border3DStyle.Adjust;
            this.navigatorTreeView.RowStyle.BorderWidth = 1;
            this.navigatorTreeView.RowStyle.WordWrap = false;
            this.navigatorTreeView.SelectBeforeEdit = true;
            this.navigatorTreeView.SelectionMode = Infralution.Controls.VirtualTree.SelectionMode.FullRow;
            this.navigatorTreeView.ShowRootRow = false;
            this.navigatorTreeView.Size = new System.Drawing.Size(255, 300);
            this.navigatorTreeView.SuppressErrorMessages = true;
            this.navigatorTreeView.TabIndex = 0;
            this.navigatorTreeView.UseThemedHeaders = false;
            this.navigatorTreeView.AfterFocusNode += new System.EventHandler<Intermech.Navigator.Controls.NavigatorTreeNodeEventArgs>(this.navigatorTreeView_AfterFocusNode);
            this.navigatorTreeView.ClearTree += new System.EventHandler(this.navigatorTreeView_ClearTree);
            // 
            // pageViewsManager
            // 
            this.pageViewsManager.ActiveViewPage = null;
            this.pageViewsManager.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pageViewsManager.Location = new System.Drawing.Point(0, 0);
            this.pageViewsManager.Name = "pageViewsManager";
            this.pageViewsManager.Size = new System.Drawing.Size(361, 300);
            this.pageViewsManager.TabIndex = 0;
            // 
            // WellKnownDockWindow
            // 
            this.Controls.Add(this.splitContainer);
            this.Name = "WellKnownDockWindow";
            this.Size = new System.Drawing.Size(620, 300);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.WellKnownDockWindow_Closing);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            this.splitContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.navigatorTreeView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        #region Взаимодействие со службой уведомлений

        /// <summary>
        /// Инициализировать службу уведомлений
        /// </summary>
        /// <returns>Ссылка на службу уведомлений</returns>
        protected virtual NotificationService InitializeNotificationService()
        {
            // Создаём экземпляр службы уведомлений
            NotificationService service = new SwitchedNotificationService();

            // Назначаем ей родительский сервис
            service.Parent = ServicesManager.GetService(typeof(INotificationService)) as NotificationService;

            // Возвращаем результат
            return service;
        }

        /// <summary>
        /// Освободить ресурсы службы уведомлений
        /// </summary>
        /// <param name="notificationService">Служба уведомлений</param>
        protected virtual void DisposeNotificationService(INotificationService notificationService)
        {
            ((IDisposable)notificationService).Dispose();
        }

        /// <summary>
        /// Управление службой уведомлений
        /// </summary>
        /// <param name="notificationService">Управляемая служба уведомлений</param>
        /// <param name="enabled">Разрешить или отключить службу уведомлений</param>
        protected virtual void EnableNotifications(INotificationService notificationService, bool enabled)
        {
            // Приводим к определённому типу
            SwitchedNotificationService switchedService = notificationService as SwitchedNotificationService;

            // Получилось
            if (switchedService != null)
                // Управляем статусом службы уведомлений
                switchedService.Enabled = enabled;
        }

        /// <summary>
        /// Служба уведомлений текущего окна
        /// </summary>
        protected INotificationService NotificationService
        {
            get { return _notificationService; }
        }

        #endregion

        #region Управление ссылками на сервисы

        /// <summary>
        /// Инициализировать сервисы (получить необходимые ссылки)
        /// </summary>
        private void InitializeServices()
        {
            // Получаем ссылки на службы
            _notificationService = InitializeNotificationService();

            // Создаём контейнеры сервисов

            // Общий
            ServiceContainer _services = new ServiceContainer();

            // Для дерева "Навигатора"
            _servicesTree = new AdvancedServiceContainer(_services);

            // Для менеджера закладок
            _servicesPages = new AdvancedServiceContainer(_services);

            // Сервис именованных изображений
            _images = ServicesManager.GetService(typeof(INamedImageList)) as INamedImageList;

            // Системный диспетчер событий
            _systemDispatcher = ServicesManager.GetService(typeof(IIODispatcher)) as IIODispatcher;

            // Получаем сервис для работы с "горячими" клавишами
            _hotKeysManager = ServicesManager.GetService(typeof(IHotKeysManager)) as IHotKeysManager;

            // Коллекции команд по умолчанию для типов объектов
            _defaultCommands4ObjTypes = ServicesManager.GetService(typeof(IDefaultCommands4ObjTypes)) as IDefaultCommands4ObjTypes;

            // Служба уведомлений создана
            if (_notificationService != null)
            {
                // Добавим её в свой контейнер сервисов
                _services.AddService(typeof(INotificationService), _notificationService);
            }

            // Зарегистрируемся как назначение для событий
            _IODispatcher.RegisterDestination(this as IIODestination);

            // Добавляем в контейнер специальные сервисы

            // Сервис менеджера закладок
            _services.AddService(typeof(IViewsManager), this.pageViewsManager);

            // Сервис для корректной работы контекстного меню и закладок
            _services.AddService(typeof(IViewState), new ViewStateService());

            // Сервис для управления колонками в дереве "Навигатора"
            _services.AddService(typeof(ITreeListColumns), this as ITreeListColumns);

            // Сервис для получения списка команд по умолчанию
            _services.AddService(typeof(IDefaultCommands4ObjTypes), _defaultCommands4ObjTypes);

            // Сервис диспетчера событий
            _services.AddService(typeof(IIODispatcher), _IODispatcher);

            // Назначаем контейнер сервисов дереву "Навигатора"
            this.navigatorTreeView.Services = _servicesTree;

            // Назначаем контейнер сервисов менеджеру закладок
            this.pageViewsManager.Services = _servicesPages;
        }

        /// <summary>   
        /// Деинициализировать сервисы
        /// </summary>
        private void DisposeServices()
        {
            // Освобождаем ссылки
            _IODispatcher.UnregisterDestination(this as IIODestination);
            _servicesTree = null;
            _servicesTree = null;
            _images = null;
        }

        #endregion

        #region Активация/деактивация контрола
        

        /// <summary>
        /// Метод вызывается, когда форма активируется
        /// </summary>
        public override void Activated()
        {
            // Базовый метод
            base.Activated();

            // Требуется активация
            if (!_activated)
            {
                // Включаем службу уведомлений
                // IsOpen - активно ли текущее окно
                // UISettings.AutoupdateNonActiveWindows - можно ли обрабатывать уведомления в неактивных формах
                EnableNotifications(NotificationService, IsOpen | UISettings.AutoupdateNonActiveWindows);
            }

            // Данные не загружены
            if (!_loaded && !this.DesignMode)
            {
                // Задаём коллекцию колонок в дереве "Навигатора"
                this.navigatorTreeView.SetColumns(Intermech.Navigator.Utils.CaptionColumnOnly(NodeColumnSortOrder.Ascending));

                // Создаём описание виртуального узла "Объекты"
                //Intermech.Navigator.DBObjectTypes. descriptor = new Intermech.Navigator.DBObjectTypes.ObjectTypesNodeDescriptor();
                //DesktopNodeDescriptor descriptor = new DesktopNodeDescriptor();
                //Intermech.Imbase.ImbaseHelper.SearchImFolderData()

                Intermech.Navigator.DBObjectTypes.Descriptor desriptor = new Intermech.Navigator.DBObjectTypes.Descriptor(1007);
                

                // Загружаем в дерево содержимое виртуального узла "Объекты"
                this.navigatorTreeView.Build(desriptor);
                // Данные загружены
                _loaded = true;
            }

            // Обновим статус элементов управления
            UpdateControls();

            // Активация выполнена
            _activated = true;
        }

        /// <summary>
        /// Контрол деактивирован
        /// </summary>
        public override void Deactivated()
        {
            // Базовый метод
            base.Deactivated();

            // Требуется деактивация
            if (_activated)
            {
                // Отключаем службу уведомлений
                // IsOpen - активно ли текущее окно
                // UISettings.AutoupdateNonActiveWindows - можно ли обрабатывать уведомления в неактивных формах
                EnableNotifications(NotificationService, IsOpen | UISettings.AutoupdateNonActiveWindows);
            }

            // Обновим статус элементов управления
            UpdateControls();

            // Деактивация выполнена
            _activated = false;
        }

        #endregion

        #region Управление статусом элементов управления

        /// <summary>
        /// Обновить статус элементов управления
        /// </summary>
        public virtual void UpdateControls()
        {
            // Установить значения visible, enabled элементам управления на форме

            // ...
        }

        #endregion

        #region Окно закрывается

        /// <summary>
        /// Окно закрывается
        /// </summary>
        /// <param name="sender">Отправитель</param>
        /// <param name="e">Аргументы события</param>
        private void WellKnownDockWindow_Closing(object sender, CancelEventArgs e)
        {
            // Можно проверить, есть ли несохранённые изменения в окне,
            // задать вопрос пользователю, при необходимости сохранить или
            // отменить изменения

            // ...
        }

        #endregion

        #region Сохранение/восстановление статуса окна

        /// <summary>
        /// Возвращает строку состояния окна, которая может быть использована для восстановления окна в
        /// следующем сеансе работы приложения.
        /// </summary>
        /// <returns>Строка состояния окна навигатора.</returns>
        protected override string GetPersistString()
        {
            try
            {
                // Получаем настройки окна в виде XML-документа
                XmlDocument xmlDoc = GetState();

                // Пишем XML в строку
                using (TextWriter textWriter = new StringWriter())
                {
                    XmlWriter xmlWriter = new XmlTextWriter(textWriter);
                    xmlDoc.WriteTo(xmlWriter);
                    xmlWriter.Flush();
                    xmlWriter.Close();

                    // Возвращаем настройки окна в виде строки
                    return textWriter.ToString();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Невозможно сохранить состояние окна!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Возвращает состояние окна в виде XML документа.
        /// </summary>
        /// <returns>Состояние окна в виде XML</returns>
        protected virtual XmlDocument GetState()
        {
            // Создаем документ и корневой узел
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode settingsNode = xmlDoc.CreateElement("SamplesClient.WellKnownDockWindow");

            // Формируем документ
            xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", null, null));
            xmlDoc.AppendChild(settingsNode);

            // Возвращаем документ
            return xmlDoc;
        }

        /// <summary>
        /// Восстановить состояние окна
        /// </summary>
        /// <param name="xmlDoc">Документ XML, в котором хранится состояние окна</param>
        protected virtual void RestoreState(XmlDocument xmlDoc)
        {
            // Восстановить все настройки из документа XML

            // ...
        }

        /// <summary>
        /// Требуется восстановление окна
        /// </summary>
        /// <param name="guid">Guid окна</param>
        /// <param name="persistString">Строка с сохранённым состоянием окна</param>
        /// <returns>Элемент управления с окном</returns>
        public static DockControl RestoreWindowCallback(Guid guid, string persistString)
        {
            // Проверочка
            if (guid != _persistStateGuid) return null;

            try
            {
                // Создаём новый документ
                XmlDocument xmlDoc = new XmlDocument();

                // Загружаем в него строку
                xmlDoc.LoadXml(persistString);

                // Создаём новое окно
                WellKnownDockWindow sampleWindow = new WellKnownDockWindow();
                try
                {
                    // Восстанавливаем ему состояние из XML
                    sampleWindow.RestoreState(xmlDoc);

                    // Возвращаем окно
                    return sampleWindow;
                }
                catch
                {
                    throw;
                }
            }
            catch (Exception e)
            {
                throw new InvalidOperationException("Невозможно восстановить состояние окна!", e);
            }
        }

        #endregion

        #region Открывает окно "Объекты"

        /// <summary>
        /// Открывает окно "Объекты"
        /// </summary>
        /// <param name="sender">Отправитель</param>
        /// <param name="e">Аргументы события</param>
        public static void ShowWellKnownDockWindow(object sender, EventArgs e)
        {
            try
            {
                //Intermech.Imbase.ImbaseHelper.SearchImFolderData()


                // Получаем ссылку на службу именованных окон "Навигатора"
                IWellKnownNavigators wkn = ServicesManager.GetService(typeof(IWellKnownNavigators)) as IWellKnownNavigators;

                //Intermech.Navigator.Interfaces.CanOpenInNewWindow r = new CanOpenInNewWindow();
                

                // Сначала ищем открытое окно "Объекты"
                WellKnownDockWindow sampleWindow = wkn.Get(WellKnownDockWindow.WindowName) as WellKnownDockWindow;

                // Окно не найдено - создаём новое окно
                if (sampleWindow == null)
                {
                    // Новое окно
                    sampleWindow = new WellKnownDockWindow();

                    MessageBox.Show("sampleWindow is null  " + sampleWindow.Name);

                    // Регистрируем окно в списке именованных окон "Навигатора"
                    wkn.Register(WellKnownDockWindow.WindowName, sampleWindow);
                }

                // Добавляем его в текущий докинг "Навигатора"
                sampleWindow.Show(ServicesManager.GetService(typeof(DockManager)) as DockManager);

                // Покажем окно на экране и сделаем его активным окном в докинге
                sampleWindow.Activate();
            }

            catch (Exception x)
            {
                MessageBox.Show(x.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        #endregion

        #region Обработчики событий от дерева "Навигатора"

        /// <summary>
        /// Изменился сфокусированный узел в дереве "Навигатора"
        /// </summary>
        /// <param name="sender">Отправитель</param>
        /// <param name="e">Аргументы события</param>
        private void navigatorTreeView_AfterFocusNode(object sender, Intermech.Navigator.Controls.NavigatorTreeNodeEventArgs e)
        {
            // Если перемещение было удачно - обновляем менеджер закладок
            this.pageViewsManager.UpdateViews(this.navigatorTreeView.SelectedItems);
        }

        /// <summary>
        /// Дерево "Навигатора" очистилось
        /// </summary>
        /// <param name="sender">Отправитель</param>
        /// <param name="e">Аргументы события</param>
        private void navigatorTreeView_ClearTree(object sender, EventArgs e)
        {
            // Закрываем закладки
            pageViewsManager.CloseViews();
        }

        #endregion

        #region Реализация интерфейса IIODestination

        /// <summary>
        /// Список поддерживаемых обработчиком событий
        /// </summary>
        public IOEventTypes SupportedEvents
        {
            get
            {
                return IOEventTypes.evKeyUp;
            }
            set
            {
                // Не требуется
            }
        }

        /// <summary>
        /// Выполнить обработку события
        /// </summary>
        /// <param name="Event">Событие</param>
        /// <returns>true, если обработка выполнена успешно, false, если событие не обработано</returns>
        public bool ProcessEvent(IIOEvent Event)
        {
            // Событие "Нажата клавиша"
            if (Event.EventType == IOEventType.evKeyUp)
            {
                // Выполняем приведение к типу
                KeyEventArgs args = Event.EventData as KeyEventArgs;

                // Нажата "F1"
                if (args != null && args.KeyData == Keys.F1)
                {
                    // Событие от клавиатуры обработано
                    args.Handled = true;

                    // Событие обработано
                    Event.EventFlags |= IOEventFlags.efProcessed;

                    // Покажем окно с сообщением
                    MessageBox.Show("Пример окна \"Объекты\"", "Событие от клавиатуры", MessageBoxButtons.OK);

                    // Событие обработано
                    return true;
                }
            }

            // Передаём событие системному диспетчеру событий
            _systemDispatcher.ProcessEvent(Event);

            // Вернём статус обработки события
            return ((Event.EventFlags & IOEventFlags.efProcessed) > 0);
        }

        #endregion

        #region Реализация интерфейса ITreeListColumns

        /// <summary>
        /// Список видимых колонок (их текущее состояние) в дереве "Навигатора"
        /// </summary>
        public NodeColumnCollection TreeListColumns
        {
            get
            {
                return this.navigatorTreeView.ReflectTreeColumsChanges();
            }
            set
            {
                this.navigatorTreeView.SetColumns(value);
            }
        }

        /// <summary>
        /// Дескриптор корневого узла в дереве окна "Навигатора"
        /// (в примере данное свойство не используется)
        /// </summary>
        public IDescriptor RootDescriptor
        {
            get
            {
                return null;
            }
            set
            {
                ;
            }
        }

        /// <summary>
        /// Контейнер сервисов окна "Навигатора"
        /// (в примере данное свойство не используется)
        /// </summary>
        public IServiceContainer Services
        {
            get { return null; }
        }

        #endregion
    }

}
