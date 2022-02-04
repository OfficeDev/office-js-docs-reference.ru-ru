| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|Возвращает объект привязки по идентификатору.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|Возвращает диаграмму по ее имени.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|Указывает комментарий, связанный с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|Возвращает объект Range, сопоставленный с именем.|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|Указывает, задано ли имя в книге или в определенной таблице.|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|Возвращает лист, к которому относится именованный элемент.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|Возвращает таблицу, в которую область действия именуемой номенклатуры.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|Получает объект `NamedItem` с его именем.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|Получает сводную таблицу по имени.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersectionornullobject-member(1))|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getusedrangeornullobject-member(1))|Возвращает используемый диапазон заданного объекта диапазона.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|Получает количество объектов `RangeView` в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|Ключ, который представляет ID параметра.|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array \| any)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|Получает запись параметра с помощью ключа.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|Получает запись параметра с помощью ключа.|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|Возникает при смене параметров документа.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|Получает объект `Setting` , представляюющий привязку, которая подняла событие изменения параметров|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|Получает таблицу по имени или ИД.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|Получает количество столбцов в таблице.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|Возвращает объект столбца по имени или идентификатору.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|Представляет коллекцию параметров, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getusedrangeornullobject-member(1))|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием.|
||[имена](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|Коллекция имен, относящих к текущему листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|Получает объект листа по его имени или ИД.|
