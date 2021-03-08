| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Создает новый документ с помощью дополнительного файла base64, закодированного .docx.|
|[Основной текст](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Возвращает весь основной текст (либо его начальную или конечную точку) в виде диапазона.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[lists](/javascript/api/word/word.body#lists)|Возвращает коллекцию объектов списков в основном тексте.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Возвращает родительский текст основного текста.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Возвращает родительский текст основного текста.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Возвращает родительский раздел основного текста.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Возвращает родительский раздел основного текста.|
||[таблицы](/javascript/api/word/word.body#tables)|Возвращает коллекцию объектов таблиц в основном тексте.|
||[type](/javascript/api/word/word.body#type)|Возвращает тип основного текста.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Возвращает или задает имя встроенного стиля основного текста.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Возвращает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Получает диапазоны текстов в области управления контентом с помощью знаков препинания и/или других знаков окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Возвращает коллекцию объектов списков в элементе управления содержимым.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Возвращает родительский текст элемента управления содержимым.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[подтип](/javascript/api/word/word.contentcontrol#subtype)|Возвращает подтип элемента управления содержимым.|
||[таблицы](/javascript/api/word/word.contentcontrol#tables)|Возвращает коллекцию объектов таблиц в элементе управления содержимым.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Возвращает или задает имя встроенного стиля для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTypes (типы: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Получает элементы управления контентом, которые имеют указанные типы и/или подтипы.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Возвращает первый элемент управления содержимым в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Возвращает первый элемент управления содержимым в коллекции.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/word/word.customproperty#key)|Возвращает ключ настраиваемого свойства.|
||[type](/javascript/api/word/word.customproperty#type)|Получает тип значения настраиваемого свойства.|
||[value](/javascript/api/word/word.customproperty#value)|Получает или задает значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Документ](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Получает свойства документа.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Открывает документ.|
||[body](/javascript/api/word/word.documentcreated#body)|Получает объект тела документа.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Получает коллекцию объектов управления контентом в документе.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Получает свойства документа.|
||[сохранено](/javascript/api/word/word.documentcreated#saved)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Получает коллекцию объектов раздела в документе.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Сохраняет документ.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[автор](/javascript/api/word/word.documentproperties#author)|Возвращает или задает автора документа.|
||[категория](/javascript/api/word/word.documentproperties#category)|Возвращает или задает категорию документа.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Возвращает или задает примечания к документу.|
||[company](/javascript/api/word/word.documentproperties#company)|Возвращает или задает компанию документа.|
||[format](/javascript/api/word/word.documentproperties#format)|Возвращает или задает формат документа.|
||[ключевые слова](/javascript/api/word/word.documentproperties#keywords)|Возвращает или задает ключевые слова документа.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Возвращает или задает менеджера документа.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Возвращает имя приложения для документа.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Возвращает дату создания документа.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Возвращает коллекцию настраиваемых свойств документа.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Получает последнего автора документа.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Возвращает дату последней печати документа.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Возвращает время последнего сохранения документа.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Возвращает номер редакции документа.|
||[безопасность](/javascript/api/word/word.documentproperties#security)|Получает параметры безопасности документа.|
||[template](/javascript/api/word/word.documentproperties#template)|Возвращает шаблон документа.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Возвращает или задает тему документа.|
||[заголовок](/javascript/api/word/word.documentproperties#title)|Возвращает или задает название документа.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Возвращает следующий встроенный рисунок.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Возвращает следующий встроенный рисунок.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Возвращает рисунок (либо его начальную или конечную точку) в виде диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Возвращает таблицу, содержащую встроенный рисунок.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Возвращает таблицу, содержащую встроенный рисунок.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Возвращает первый встроенный рисунок в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Возвращает первый встроенный рисунок в коллекции.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Возвращает абзацы, обнаруженные на указанном уровне списка.|
||[getLevelString (уровень: номер)](/javascript/api/word/word.list#getlevelstring-level-)|Получает пулю, номер или изображение на указанном уровне в качестве строки.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[id](/javascript/api/word/word.list#id)|Получает id списка.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Проверяет наличие каждого из 9 уровней в списке.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Возвращает типы всех 9 уровней списка.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Возвращает абзацы в списке.|
||[setLevelAlignment (уровень: номер, выравнивание: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Задает выравнивание пули, номера или изображения на указанном уровне в списке.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Задает формат маркеров на указанном уровне списка.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Задает два отступа на указанном уровне списка.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Задает формат нумерации на указанном уровне списка.|
||[setLevelStartingNumber (уровень: номер, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Задает начальный номер на указанном уровне списка.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Возвращает список по идентификатору.|
||[getByIdOrNullObject (id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Возвращает список по идентификатору.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Возвращает первый список в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Возвращает первый список в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Возвращает объект списка по индексу в коллекции.|
||[items](/javascript/api/word/word.listcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getDescendants (directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Возвращает всех потомков элемента списка.|
||[level](/javascript/api/word/word.listitem#level)|Возвращает или задает уровень элемента в списке.|
||[listString](/javascript/api/word/word.listitem#liststring)|Получает пулю элемента списка, номер или изображение в качестве строки.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Возвращает порядковый номер элемента списка относительно элементов того же уровня.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Позволяет присоединить абзац к существующему списку на указанном уровне.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Перемещает абзац за пределы списка (если он является элементом списка).|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Возвращает следующий абзац.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Возвращает следующий абзац.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Возвращает предыдущий абзац.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Возвращает предыдущий абзац.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Возвращает весь абзац (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Получает диапазоны текста в абзаце, используя знаки препинания и/или другие знаки окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Указывает, что абзац является последним в родительском тексте.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Проверяет, является ли абзац элементом списка.|
||[list](/javascript/api/word/word.paragraph#list)|Возвращает объект List, к которому относится абзац.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Возвращает объект ListItem для абзаца.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Возвращает объект ListItem для абзаца.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Возвращает объект List, к которому относится абзац.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Возвращает родительский текст абзаца.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий абзац.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Возвращает таблицу, содержащую абзац.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Возвращает таблицу, содержащую абзац.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Возвращает уровень таблицы, содержащей абзац.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Разделяет абзац на дочерние диапазоны с помощью разделителей.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Создает список, начинающийся с данного абзаца.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Возвращает или задает имя встроенного стиля абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Возвращает первый абзац в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Возвращает первый абзац в коллекции.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Возвращает последний абзац в коллекции.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Возвращает последний абзац в коллекции.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (диапазон: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Сравнивает расположение данного диапазона с расположением другого диапазона.|
||[expandTo (диапазон: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[expandToOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Возвращает дочерние диапазоны гиперссылок в данном диапазоне.|
||[getNextTextRange (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getNextTextRangeOrNullObject (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Получает текстовые детские диапазоны в диапазоне, используя знаки препинания и/или другие знаки окончания.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[intersectWith (диапазон: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[intersectWithOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Проверяет, является ли длина диапазона нулевой.|
||[lists](/javascript/api/word/word.range#lists)|Возвращает коллекцию объектов списков в диапазоне.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Возвращает родительский текст диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Возвращает таблицу, содержащую диапазон.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Возвращает таблицу, содержащую диапазон.|
||[таблицы](/javascript/api/word/word.range#tables)|Возвращает коллекцию объектов таблиц в диапазоне.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Возвращает или задает имя встроенного стиля диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Возвращает первый диапазон в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Возвращает первый диапазон в коллекции.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Набор API: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Возвращает следующий раздел.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Возвращает следующий раздел.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Возвращает первый раздел в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Возвращает первый раздел в коллекции.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона.|
||[выравнивание](/javascript/api/word/word.table#alignment)|Получает или задает выравнивание таблицы со столбцом страницы.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|
||[clear()](/javascript/api/word/word.table#clear--)|Очищает содержимое таблицы.|
||[delete()](/javascript/api/word/word.table#delete--)|Удаляет всю таблицу.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Удаляет определенные столбцы.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Удаляет определенные строки.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Равномерно распределяет ширину столбцов.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Возвращает стиль указанной границы.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellOrNullObject (rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Возвращает следующую таблицу.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Возвращает следующую таблицу.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Возвращает абзац после таблицы.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Возвращает абзац после таблицы.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Возвращает абзац перед таблицей.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Возвращает абзац перед таблицей.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Возвращает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Возвращает и задает количество строк заголовков.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в таблице.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Вставляет в таблицу элемент управления содержимым.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Возвращает шрифт.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Указывает, однородны ли все строки таблицы.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Возвращает уровень вложенности таблицы.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Возвращает родительский текст таблицы.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Возвращает таблицу, которая содержит данную таблицу.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Возвращает таблицу, которая содержит данную таблицу.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Получает количество строк в таблице.|
||[строки](/javascript/api/word/word.table#rows)|Возвращает все строки таблицы.|
||[таблицы](/javascript/api/word/word.table#tables)|Возвращает дочерние таблицы, вложенные на один уровень ниже.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {
            ignorePunct?: boolean
            ignoreSpace?: boolean
            matchCase?: boolean
            matchPrefix?: boolean
            matchSuffix?: boolean
            matchWholeWord?: boolean
            matchWildcards?: boolean
        })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)| Выбирает таблицу или положение в начале или конце таблицы и перемещает пользовательский интерфейс Word в него.| || [setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)| Задает обивку ячейки в points.| || [shadingColor](/javascript/api/word/word.table#shadingcolor)| Получает и задает затеняющий цвет.| || [стиль](/javascript/api/word/word.table#style)| Получает или задает имя стиля для таблицы.| || [styleBandedColumns |](/javascript/api/word/word.table#stylebandedcolumns) Получает и задает, имеет ли таблица полосатую колонку.| || [styleBandedRows](/javascript/api/word/word.table#stylebandedrows)| Получает и задает, имеет ли таблица полосатую строку.| || [styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)| Получает или задает встроенное имя стиля для таблицы.| || [styleFirstColumn |](/javascript/api/word/word.table#stylefirstcolumn) Получает и задает, есть ли в таблице первый столбец со специальным стилем.| || [styleLastColumn |](/javascript/api/word/word.table#stylelastcolumn) Получает и задает, есть ли в таблице последний столбец со специальным стилем.| || [styleTotalRow](/javascript/api/word/word.table#styletotalrow)| Получает и задает, имеет ли таблица общую (последнюю) строку со специальным стилем.| || [значения |](/javascript/api/word/word.table#values) Получает и задает текстовые значения в таблице в виде массива Javascript 2D.| || [verticalAlignment](/javascript/api/word/word.table#verticalalignment)| Получает и задает вертикальное выравнивание каждой ячейки в таблице.| || [ширина](/javascript/api/word/word.table#width)| Получает и задает ширину таблицы в точках.| | [TableBorder](/javascript/api/word/word.tableborder) | [цвет |](/javascript/api/word/word.tableborder#color) Получает или задает цвет границы таблицы.| || [тип](/javascript/api/word/word.tableborder#type)| Получает или задает тип границы таблицы.| || [ширина](/javascript/api/word/word.tableborder#width)| Получает или задает ширину в точках границы таблицы.| | [TableCell](/javascript/api/word/word.tablecell) | [columnWidth](/javascript/api/word/word.tablecell#columnwidth)| Получает и задает ширину столбца ячейки в точках.| || [deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)| Удаляет столбец, содержащий эту ячейку.| || [deleteRow()](/javascript/api/word/word.tablecell#deleterow--)| Удаляет строку, содержащую эту ячейку.| || [getBorder (borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)| Получает пограничный стиль для указанной границы.| || [getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)| Получает обивку ячейки в points.| || [getNext()](/javascript/api/word/word.tablecell#getnext--)| Получает следующую ячейку.| || [getNextOrNullObject() |](/javascript/api/word/word.tablecell#getnextornullobject--) Получает следующую ячейку.| || [horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)| Получает и задает горизонтальное выравнивание ячейки.| || [insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[]])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)| Добавляет столбцы влево или вправо ячейки, используя столбец ячейки в качестве шаблона.| || [insertRows (insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)| Вставляет строки выше или ниже ячейки, используя строку ячейки в качестве шаблона.| || [тело](/javascript/api/word/word.tablecell#body)| Получает объект тела ячейки.| || [cellIndex](/javascript/api/word/word.tablecell#cellindex)| Получает индекс ячейки в строке.| || [parentRow](/javascript/api/word/word.tablecell#parentrow)| Получает родительский ряд ячейки.| || [parentTable](/javascript/api/word/word.tablecell#parenttable)| Получает родительную таблицу ячейки.| || [rowIndex](/javascript/api/word/word.tablecell#rowindex)| Получает индекс строки ячейки в таблице.| || [ширина](/javascript/api/word/word.tablecell#width)| Получает ширину ячейки в точках.| || [setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)| Задает обивку ячейки в points.| || [shadingColor](/javascript/api/word/word.tablecell#shadingcolor)| Получает или задает затеняющий цвет ячейки.| || [значение](/javascript/api/word/word.tablecell#value)| Получает и задает текст ячейки.| || [verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)| Получает и задает вертикальное выравнивание ячейки.| | [TableCellCollection](/javascript/api/word/word.tablecellcollection) | [getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)| Получает первую ячейку таблицы в этой коллекции.| || [getFirstOrNullObject() |](/javascript/api/word/word.tablecellcollection#getfirstornullobject--) Получает первую ячейку таблицы в этой коллекции.| || [элементы](/javascript/api/word/word.tablecellcollection#items)| Получает загруженные детские элементы в этой коллекции.| | [TableCollection](/javascript/api/word/word.tablecollection) | [getFirst()](/javascript/api/word/word.tablecollection#getfirst--)| Получает первую таблицу в этой коллекции.| || [getFirstOrNullObject() |](/javascript/api/word/word.tablecollection#getfirstornullobject--) Получает первую таблицу в этой коллекции.| || [элементы](/javascript/api/word/word.tablecollection#items)| Получает загруженные детские элементы в этой коллекции.| | [TableRow](/javascript/api/word/word.tablerow) | [clear()](/javascript/api/word/word.tablerow#clear--)| Очищает содержимое строки.| || [delete()](/javascript/api/word/word.tablerow#delete--)| Удаляет всю строку.| || [getBorder (borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)| Получает пограничный стиль ячеек в строке.| || [getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)| Получает обивку ячейки в points.| || [getNext()](/javascript/api/word/word.tablerow#getnext--)| Получает следующую строку.| || [getNextOrNullObject() |](/javascript/api/word/word.tablerow#getnextornullobject--) Получает следующую строку.| || [horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)| Получает и задает горизонтальное выравнивание каждой ячейки в строке.| || [ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)|| || [ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)|| || [insertRows (insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)| Вставляет строки, используя эту строку в качестве шаблона.| || [matchCase](/javascript/api/word/word.tablerow#matchcase)|| || [matchPrefix](/javascript/api/word/word.tablerow#matchprefix)|| || [matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)|| || [matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)|| || [matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)|| || [preferredHeight](/javascript/api/word/word.tablerow#preferredheight)| Получает и задает предпочитаемую высоту строки в точках.| || [cellCount](/javascript/api/word/word.tablerow#cellcount)| Получает число ячеек в строке.| || [ячейки](/javascript/api/word/word.tablerow#cells)| Получает cells.| || [шрифт](/javascript/api/word/word.tablerow#font)| Получает шрифт.| || [isHeader](/javascript/api/word/word.tablerow#isheader)| Проверяет, является ли строка строкой загона.| || [parentTable](/javascript/api/word/word.tablerow#parenttable)| Получает родительская таблица.| || [rowIndex](/javascript/api/word/word.tablerow#rowindex)| Получает индекс строки в родительской таблице.| || [search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)| Выполняет поиск с указанными SearchOptions в области строки.| || [select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)| Выбирает строку и перемещает пользовательский интерфейс Word в него.| || [setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)| Задает обивку ячейки в points.| || [shadingColor](/javascript/api/word/word.tablerow#shadingcolor)| Получает и задает затеняющий цвет.| || [значения |](/javascript/api/word/word.tablerow#values) Получает и задает текстовые значения в строке в виде массива Javascript 2D.| || [verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)| Получает и задает вертикальное выравнивание ячеек в строке.| | [TableRowCollection](/javascript/api/word/word.tablerowcollection) | [getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)| Получает первую строку в этой коллекции.| || [getFirstOrNullObject() |](/javascript/api/word/word.tablerowcollection#getfirstornullobject--) Получает первую строку в этой коллекции.| || [элементы](/javascript/api/word/word.tablerowcollection#items)| Получает загруженные детские элементы в этой коллекции.|
