| Класс | Поля | Описание |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete__)|Удаляет пользовательскую XML-часть.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getXml__)|Получает полное содержимое пользовательской XML-части.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|Пользовательский ID части XML.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceUri)|Пользовательское пространство имен XML-части URI.|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#setXml_xml_)|Задает полное содержимое пользовательской XML-части.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add_xml_)|Добавляет новую пользовательскую XML-часть в книгу.|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getByNamespace_namespaceUri_)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getCount__)|Получает количество пользовательских частей XML в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getCount__)|Получает количество частей CustomXML в этой коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItem__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID of the PivotTable.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[Набор API: ExcelApi 1.5]|
|[Runtime](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customXmlParts)|Представляет коллекцию пользовательских частей XML, содержащихся в этой книге.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNext_visibleOnly_)|Получает таблицу, которая следует за этим.|
||[getNextOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNextOrNullObject_visibleOnly_)|Получает таблицу, которая следует за этим.|
||[getPrevious (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPrevious_visibleOnly_)|Получает таблицу, предшествующего этому.|
||[getPreviousOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPreviousOrNullObject_visibleOnly_)|Получает таблицу, предшествующего этому.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getFirst_visibleOnly_)|Получает первый лист в коллекции.|
||[getLast (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getLast_visibleOnly_)|Получает последний лист в коллекции.|
