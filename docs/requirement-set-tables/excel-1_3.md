| Класс | Поля | Описание |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Удаляет привязку.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: \| строка Range, bindingType: Excel. bindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Добавляет привязку к определенному объекту Range.|
||[Аддфромнамедитем (имя: строка, bindingType: Excel. BindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Добавляет новую привязку с учетом именованного элемента в книге.|
||[Аддфромселектион (bindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Лист, содержащий текущую сводную таблицу.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Обновляет сводную таблицу.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Обновляет все сводные таблицы в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[Жетвисиблевиев ()](/javascript/api/excel/excel.range#getvisibleview--)|Представляет видимые строки текущего диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Представляет формулу в формате R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Получает родительский диапазон, сопоставленный с текущим объектом RangeView.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Представляет адреса ячеек RangeView.|
||[Число](/javascript/api/excel/excel.rangeview#columncount)|Число видимых столбцов.|
||[index](/javascript/api/excel/excel.rangeview#index)|Возвращает значение, представляющее индекс RangeView.|
||[Стро](/javascript/api/excel/excel.rangeview#rowcount)|Количество видимых строк.|
||[строки](/javascript/api/excel/excel.rangeview#rows)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном.|
||[text](/javascript/api/excel/excel.rangeview#text)|Текстовые значения указанного диапазона.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Представляет тип данных каждой ячейки.|
||[values](/javascript/api/excel/excel.rangeview#values)|Представляет необработанные значения указанного объекта rangeView.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Получает строку RangeView с помощью индекса.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Указывает, будут ли в столбцах отображаться чередующиеся форматирование, в котором четные столбцы выделяются не так, как в четном, чтобы упростить чтение таблицы.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Указывает, будут ли в строках отображаться полосные форматирования, в результате которой нечетные строки выделяются не так, как в четном, чтобы облегчить чтение таблицы.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Указывает, отображаются ли кнопки фильтра в верхней части каждого заголовка столбца.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Сводные таблицы](/javascript/api/excel/excel.workbook#pivottables)|Представляет коллекцию сводных таблиц, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Сводные таблицы](/javascript/api/excel/excel.worksheet#pivottables)|Коллекция сводных таблиц на листе.|
