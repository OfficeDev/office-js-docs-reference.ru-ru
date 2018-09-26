# <a name="excel-javascript-api-overview"></a>Обзор интерфейса API JavaScript для Excel

Excel JavaScript API можно использовать для создания надстроек для Excel 2016 или более поздней версии. Ниже перечислены объекты Excel высокого уровня, доступные в API. Ссылки на страницы каждый объект содержит описание свойства, события и методы, доступные для объекта. Чтобы узнать больше, перейдите по соответствующим ссылкам в меню.

Для удобства ниже перечислены некоторые из основных объектов Excel. 

- [Workbook](/javascript/api/excel/excel.workbook) — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д. Его также можно использовать для вывода списка связанных ссылок.

- [Worksheet](/javascript/api/excel/excel.worksheet). Представляет лист в книге. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection). Коллекция объектов **Worksheet** в книге.

- [Range](/javascript/api/excel/excel.range). Представляет ячейку, строку, столбец или группу ячеек, содержащую один или несколько смежных блоков ячеек.

- [Table](/javascript/api/excel/excel.table). Представляет коллекцию упорядоченных ячеек, которая упрощает управление данными.
    - [TableCollection](/javascript/api/excel/excel.tablecollection). Коллекция таблиц в книге или на листе.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection). Коллекция всех столбцов в таблице.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection). Коллекция всех строк в таблице.

- [Chart](/javascript/api/excel/excel.chart). Представляет объект диаграммы на листе, который является визуальным представлением базовых данных.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection). Коллекция диаграмм на листе.

- [TableSort](/javascript/api/excel/excel.tablesort). Представляет объект, управляющий операциями сортировки для объектов **Table**.

- [RangeSort](/javascript/api/excel/excel.rangesort). Представляет объект, управляющий операциями сортировки для объектов **Range**.

- [Filter](/javascript/api/excel/excel.filter). Представляет объект, управляющий фильтрацией столбца таблицы.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection). Представляет защиту объекта **Worksheet**.

- [NamedItem](/javascript/api/excel/excel.nameditem). Представляет определенное имя для диапазона ячеек или значения. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection). Коллекция объектов **NamedItem** в книге.

- [Binding](/javascript/api/excel/excel.binding). Абстрактный класс, представляющий привязку к разделу книги.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection). Коллекция объектов **Binding** в книге.

## <a name="excel-javascript-api-open-specifications"></a>Открытые спецификации API JavaScript для Excel

Как мы проектирования и разработки новые интерфейсы API для надстроек Excel, мы будем сделать их доступными для свой отзыв на страницу, где [спецификации Open API](../openspec.md) . Узнайте, что новые функции на конвейере для API JavaScript, Excel и предоставление данных, введенных на нашим спецификациям.

## <a name="excel-javascript-api-reference"></a>Справочник по API JavaScript для Excel

Получить подробные сведения об API JavaScript для Excel обратитесь к [Справочная документация по Excel JavaScript API](/javascript/api/excel).

## <a name="see-also"></a>См. также

- [Общие сведения о надстройках Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Обзор платформы надстроек Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Примеры надстроек на репозиториев Excel](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
