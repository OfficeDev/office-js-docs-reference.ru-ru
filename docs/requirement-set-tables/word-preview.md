| Класс | Поля | Описание |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Происходит при смене данных в области управления контентом.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Происходит при удалении управления контентом.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Возникает при смене выбора в области управления контентом.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|Объект, который поднял событие.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|Тип события.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Удаляет пользовательскую XML-часть.|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Удаляет атрибут с заданным именем из элемента, идентифицированного xpath.|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Удаляет элемент, идентифицированный xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Получает полное XML-содержимое пользовательской части XML.|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Вставляет атрибут с заданным именем и значением в элемент, идентифицированный xpath.|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Вставляет данный XML в родительский элемент, идентифицированный xpath в индексе положения ребенка.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Запрашивает XML-содержимое пользовательской части XML.|
||[id](/javascript/api/word/word.customxmlpart#id)|Получает ID пользовательской части XML.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Получает URI пространства имен пользовательской части XML.|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Задает полное XML-содержимое пользовательской части XML.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Обновляет значение атрибута с заданным именем элемента, идентифицированного xpath.|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Обновляет XML элемента, идентифицированного xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Добавляет в документ новую настраиваемую часть XML.|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (имя: строка)](/javascript/api/word/word.document#deleteBookmark_name_)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.document#getBookmarkRange_name_)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Получает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Получает настраиваемые XML-части в документе.|
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Возникает при добавлении управления контентом.|
||[settings](/javascript/api/word/word.document#settings)|Получает параметры надстройки в документе.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (имя: строка)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Получает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Получает настраиваемые XML-части в документе.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Получает параметры надстройки в документе.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Получает формат inline image.|
|[Список](/javascript/api/word/word.list)|[getLevelFont (уровень: номер)](/javascript/api/word/word.list#getLevelFont_level_)|Получает шрифт пули, номера или изображения на указанном уровне в списке.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Получает кодированное представление строки base64 на указанном уровне в списке.|
||[resetLevelFont (уровень: номер, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Сброс шрифта пули, номера или изображения на указанном уровне в списке.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Задает изображение на указанном уровне в списке.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Получает имена всех закладки или перекрывает диапазон.|
||[insertBookmark (имя: строка)](/javascript/api/word/word.range#insertBookmark_name_)|Вставляет закладки в диапазоне.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#key)|Получает ключ параметра.|
||[value](/javascript/api/word/word.setting#value)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Создает новый параметр или задает существующий параметр.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Удаляет все параметры в этой надстройки.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[items](/javascript/api/word/word.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Объединяет ячейки, ограниченные включительно первой и последней ячейками.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Вставляет управление контентом в строку.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Сливает строку в одну ячейку.|
