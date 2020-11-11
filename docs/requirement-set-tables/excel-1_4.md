| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Возвращает объект привязки по идентификатору.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Возвращает диаграмму по ее имени.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Задает комментарий, связанный с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Возвращает объект Range, сопоставленный с именем.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Указывает, ограничивается ли имя книгой или определенным листом.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Возвращает лист, к которому относится именованный элемент.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Возвращает лист, к которому относится именованный элемент.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (имя: строка, ссылка: \| строка диапазона, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, Примечание?: строка)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Возвращает объект NamedItem, используя его имя.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Получает сводную таблицу по имени.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: \| строка Range)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Возвращает используемый диапазон заданного объекта диапазона.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Получает количество объектов RangeView в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#key)|Ключ, представляющий идентификатор параметра.|
||[value](/javascript/api/excel/excel.setting#value)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (ключ: строка, значение: строка \| Number \| Boolean \| \| массив дат <any> \| Any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Получает запись Setting по ключу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Получает запись Setting по ключу.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Возникает при изменении параметров в документе.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Получает объект Setting, представляющий привязку, которая вызвала событие SettingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Получает таблицу по имени или идентификатору.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Получает количество столбцов в таблице.|
||[getItemOrNullObject (Key: номер \| строки)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Возвращает объект column по имени или идентификатору.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Представляет коллекцию параметров, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием.|
||[псевдоним](/javascript/api/excel/excel.worksheet#names)|Коллекция имен, относящих к текущему листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[NOCOUNT (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Получает объект листа по его имени или ИД.|
