using Intermech;
using Intermech.DataFormats;
using Intermech.Interfaces;
using Intermech.Navigator.ContextMenu;
using Intermech.Navigator.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddFeatureContextMenu
{
    internal class CommandProvider : ICommandsProvider
    {
        #region Реализация интерфейса ICommandsProvider

        /// <summary>
        /// Метод вызывается для получения допустимых и подавляемых команд контекстного меню для выделенных 
        /// элементов навигации одной категории и типа. 
        /// Например, если в «Навигаторе» выделены элементы навигации нескольких разных категорий и типов, 
        /// то данная команда будет вызываться для каждой из подгрупп этих элементов, сгруппированных по их 
        /// категориям и типам. 
        /// Это наиболее применяемый метод даного интерфейса. Позволяет перекрывать команды контекстного меню 
        /// для элементов навигации определённых категорий, типов, задавая более высокий приоритет описаниям 
        /// этих команд.
        /// </summary>
		/// <param name="items">
        /// Коллекция выделенных элементов пространства навигации, для которых выполняется 
        /// поиск допустимых команд контекстного меню.
        /// </param>
		/// <param name="viewServices">
        /// Контейнер сервисов (контекст) для выделенных элементов пространства навигации.
        /// </param>
        public CommandsInfo GetMergedCommands(ISelectedItems items, IServiceProvider viewServices)
        {
            // ВНИМАНИЕ! Основное требование к данному методу – нельзя выполнять обращения к базе даных 
            // для того, чтобы проверить, можно ли отображать команду меню или нет!

            // Список добавленных или перекрытых команд контекстного меню
            CommandsInfo commandsInfo = new CommandsInfo();

            // Есть один выделенный элемент
            if (items != null && items.Count == 1)
            {
                // Пробуем получить описание выделенного объекта
                IDBTypedObjectID objID = items.GetItemData(0, typeof(IDBTypedObjectID)) as IDBTypedObjectID;

                // Выделен объект
                if (objID != null)
                {
                    // Получаем идентификатор типа объектов "Рабочий стол"
                    // Разрешать создавать его копию не будем
                    Int32 desktopObjectTypeID = MetaDataHelper.GetObjectTypeID(SystemGUIDs.objtypeWorkspace);

                    // Можем добавить команду "Создать\Копию объекта"
                    if (objID.ObjectType != desktopObjectTypeID)
                    {
                        // Команда "Создать\Копию объекта"
                        commandsInfo.Add("CreateCopyyyyyyy",
                            new CommandInfo(TriggerPriority.ItemCategory,
                            new ClickEventHandler(CommandProvider.CreateObjectCopyyy)));
                    }
                }
            }

            // Вернём список
            return commandsInfo;
        }

        /// <summary>
        /// Метод вызывается для получения допустимых и подавляемых команд контекстного меню для всей группы выделенных 
        /// элементов навигации. Особенности данного метода:
        /// 1. Если команда зарегистрирована на все категории, то метод вызывается один раз и 
        ///    получает в качестве параметра items все выделенные в «Навигаторе» элементы навигации;
        /// 2. Если команда зарегистрирована на конкретную категорию, то метод будет вызван один раз для всех 
        ///    выделенных элементов навигации только в том случае, если все они принадлежат одной категории; 
        ///    для всех выделенных элементов навигации только в том случае, если все они принадлежат указанной категории;
        /// 3. Если команда зарегистрирована на конкретные категорию и тип, то метод будет вызван один раз для всех 
        ///    выделенных элементов навигации только в том случае, если все они принадлежат указанной категории и типу.
        /// </summary>
        /// <param name="items">
        /// Коллекция выделенных элементов пространства навигации, для которых выполняется 
        /// поиск допустимых команд контекстного меню.
        /// </param>
		/// <param name="viewServices">
        /// Контейнер сервисов (контекст) для выделенных элементов пространства навигации.
        /// </param>
        public CommandsInfo GetGroupCommands(ISelectedItems items, IServiceProvider viewServices)
        {
            // ВНИМАНИЕ! Основное требование к данному методу – нельзя выполнять обращения к базе даных 
            // для того, чтобы проверить, можно ли отображать команду меню или нет!

            // Список добавленных или перекрытых команд контекстного меню
            CommandsInfo commandsInfo = new CommandsInfo();

            // Вернём список
            return commandsInfo;
        }

        #endregion

        #region Обработчик команды контекстного меню "Создать\Копию объекта"

        /// <summary>
        /// Пример №6: Выполнить команду контекстного меню "Создать\Копию объекта"
        /// </summary>
        /// <param name="items">Коллекция выделенных элементов пространства навигации, для которых вызвано контекстное меню</param>
        /// <param name="viewServices">Контейнер сервисов (контекст) для выделенных элементов пространства навигации</param>
        /// <param name="additionalInfo">Дополнительные данные для команды контекстного меню</param>
        internal static void CreateObjectCopyyy(ISelectedItems items, IServiceProvider viewServices, object additionalInfo)
        {

        }

        #endregion
    }

}
