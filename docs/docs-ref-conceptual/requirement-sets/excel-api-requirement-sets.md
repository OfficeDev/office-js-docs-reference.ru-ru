# <a name="excel-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Excel

Наборы обязательных элементов — это именованные группы элементов API. Надстройки Office использовать наборов требований, указанный в манифесте или выполняется проверка среды выполнения для определения поддержки API, которые требуется добавить в приложение Office. Дополнительные сведения см в [различных версиях Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Excel запустите для нескольких версий Office, включая Office 2016 или более поздней версии для Windows, Office для iPad, Office для Mac и Office Online. В следующей таблице перечислены наборы требований Excel, ведущих приложений Office, которые поддерживают каждый набор требований, а также версиях построения или номер для этих приложений.

> [!NOTE]
> Любой API, помеченные как **бета-версии** не готова для конечных пользователей рабочей среды. Мы сделать их доступными для разработчиков попробовать их извлечения в средах разработки и тестирования. Они не предназначены для использования с рабочей/бизнес-важных документов.
> 
> Для наборов требований, которые помечены как **бета-версии**, используйте указанный (или более поздней версии) версии программного обеспечения Office и использовать бета-версию библиотеки на CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Записи не помечена как **бета-версии** обычно доступны и производства библиотеки можно использовать на CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Набор требований  |  Office 365 для Windows\*  |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Бета-версия  | Пожалуйста, [посетите страницу открытых спецификаций Excel JavaScript API](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.8  | Версия 1808 (построение 10730.20102) или более поздней версии | 2.17 или более поздней версии | 16.17 или более поздней версии | Сентябрь 2018 | Ожидается в скором времени |
| ExcelApi1.7  | Версия 1801 (построение 9001.2171) или более поздней версии   | 2,9 или более поздняя версия | 16,9 или более поздняя версия | Апрель 2018 г. | Скоро |
| ExcelApi1.6  | Версия 1704 (сборка 8201.2001) или более поздняя   | Версия 2.2 или более поздняя |Версия 15.36 или более поздняя| Апрель 2017 г. | Скоро|
| ExcelApi1.5  | Версия 1703 (сборка 8067.2070) или более поздняя   | Версия 2.2 или более поздняя |Версия 15.36 или более поздняя| Март 2017 г. | Скоро|
| ExcelApi1.4  | Версия 1701 (сборка 7870.2024) или более поздняя   | Версия 2.2 или более поздняя |Версия 15.36 или более поздняя| Январь 2017 г. | Скоро|
| ExcelApi1.3  | Версия 1608 (сборка 7369.2055) или более поздняя | 1.27 или более поздняя |  15.27 или более поздняя| Сентябрь 2016 г. | Версия 1608 (сборка 7601.6800) или более поздняя|
| ExcelApi1.2  | Версия 1601 (сборка 6741.2088) или более поздняя | 1.21 или более поздняя | 15.22 или более поздняя| Январь 2016 г. ||
| ExcelApi1.1  | Версия 1509 (сборка 4266.1001) или более поздняя | 1.19 или более поздняя | 15.20 или более поздняя| Январь 2016 г. ||

> [!NOTE]
> Номер сборки 2016 Office установлен с помощью MSI — 16.0.4266.1001. Эта версия содержит только наборы требований ExcelApi 1.1.

Дополнительные сведения о версии, номера сборок и Office Online Server можно:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- 
  [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview).

## <a name="whats-new-in-excel-javascript-api-18"></a>Новые возможности Excel 1,8 API JavaScript

Возможности Excel JavaScript API требование set 1,8 включают API-интерфейсы для сводных таблиц, выполнить проверку данных, диаграмм, события для диаграмм, параметры быстродействия и создания книги.

### <a name="pivottable"></a>Сводная таблица

2 звукового файла API-интерфейсов сводной таблицы позволяет надстроек иерархии сводной таблицы. Теперь можно управлять данные и как обобщения. Наш [сводной таблицы в статье](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) имеет несколько новых функциональных возможностей сводной таблицы.

### <a name="data-validation"></a>проверку данных;

Данные проверки дает возможность из какой пользователь вводит на листе. Можно ограничить ячеек, чтобы предварительно заданные ответов наборов или предоставить всплывающих предупреждений о нежелательный входные данные. Дополнительные сведения о [добавлении проверки данных к диапазонам](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) сегодня.

### <a name="charts"></a>диаграммы;

Другой круговой диаграммы API-интерфейсы сводит программного управления элементы диаграммы. Имеется доступ к легенды, осей, линия тренда и область построения.

### <a name="events"></a>События

Добавлены дополнительные [события](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) диаграмм. У вашей надстройки реагировать на пользователей, которым взаимодействует с диаграммой. Также можно [Переключить событий](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) обработки по всей книги.


|Объект| Новые возможности| Описание|Наборы требований|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Метод_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Создает новую книгу скрытых с помощью файла закодированный .xlsx необязательно base64.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Свойство_ > formula1|Получает или задает Formula1, то есть минимальное значение или значение в зависимости от оператора.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Свойство_ > formula2|Получает или задает Formula2, то есть максимальное значение или значение в зависимости от оператора.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Отношения_ > оператор|Оператор, используемый для проверки данных.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > categoryLabelLevel|Возвращает или задает константа перечисления ChartCategoryLabelLevel ссылается на уровне где подписи категорий извлеченные из. Чтение и запись.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > plotVisibleOnly|Значение true, если только видимые ячейки. False, если оба отображаемые и скрытые ячеек на диаграмме. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > seriesNameLevel|Возвращает или задает константа перечисления ChartSeriesNameLevel ссылается на уровне где имена являются, извлеченные из. Чтение и запись.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > showDataLabelsOverMaximum|Представляет, следует ли отображать метки данных, если значение больше, чем максимальное значение на оси значений.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > style|Возвращает или задает стиль диаграммы для диаграммы. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Отношения_ > displayBlanksAs|Возвращает или задает способ, что пустые ячейки будут отображаться на диаграмме. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Отношения_ > plotArea|Представляет plotArea для диаграммы. Только для чтения.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Отношения_ > plotBy|Возвращает или задает способ столбцы или строки используются в качестве рядов данных на диаграмме. ReadWrite.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Свойство_ > chartId|Получает идентификатор диаграммы, который активируется.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Свойство_ > тип|Получает тип события.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в котором активируется диаграммы.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Свойство_ > chartId|Получает идентификатор, который добавляется в лист диаграммы.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Свойство_ > тип|Получает тип события.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в которую добавляется диаграммы.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Отношения_ > источник|Получает источник события.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > isBetweenCategories|Представляет ли ось пересечения оси категорий.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > многоуровневой|Указывает, находится ли оси многоуровневой или нет.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > numberFormat|Представляет код формата для метки делений оси.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > смещение|Представляет расстояние между уровнями меток и расстояние между первый уровень и линии оси. Значение должно быть целое число от 0 до 1000.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > positionAt|Представляет пересечения другой оси в положение указанной оси. Чтобы задать это свойство, следует использовать метод SetPositionAt(double). Только для чтения.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > textOrientation|Представляет ориентации текста тактов подпись оси. Значение должно быть целое число либо от -90 до 90 или 180 для вертикально ориентированного текста.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > Выравнивание|Представляет выравнивание для метки делений указанной оси.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > позиции|Представляет положение указанного оси, пересечения другой оси.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Метод_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Задайте положение указанного оси, где другой оси пересечение в.|1,8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Отношения_ > заливки|Представляет параметры форматирования заливки диаграммы. Только для чтения.|1,8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Метод_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Строковое значение, представляющее формула заголовка оси диаграммы с помощью нотации стиля A1.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Отношения_ > границы|Представляет формат границы, включая цвет, линии и вес. Только для чтения.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Отношения_ > заливки|Представляет параметры форматирования заливки диаграммы. Только для чтения.|1,8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Метод_ > [clear()](/javascript/api/excel/excel.chartborder)|Очистить формат границы элемента диаграммы.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > Автотекст|Логическое значение, указывающее Если подписей данных автоматически создает текст на основе контекста.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > формулы|Строковое значение, представляющее формула нотации стиля A1 подпись данных диаграммы.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > height|Возвращает высоту, в пунктах метки данных диаграммы. Только для чтения. NULL, если подпись данных диаграммы не отображается. Только для чтения.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > слева|Представляет расстояние в пунктах от левого края диаграмму метки данных для левого края области диаграммы. NULL, если подпись данных диаграммы не отображается.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > numberFormat|Строковое значение, представляющее код формата для метки данных.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > text|Строка, представляющая текст метки данных на диаграмме.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > textOrientation|Представляет ориентации текста метки данных диаграммы. Значение должно быть целое число либо от -90 до 90 или 180 для вертикально ориентированного текста.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > в начало|Представляет расстояние в пунктах от верхнего края подпись данных диаграммы в верхней части области диаграммы. NULL, если подпись данных диаграммы не отображается.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > width|Возвращает ширину в точках метки данных диаграммы. Только для чтения. NULL, если подпись данных диаграммы не отображается. Только для чтения.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Отношения_ > формат|Представляет формат метки данных диаграммы. Только для чтения.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Связь_ > horizontalAlignment|Представляет горизонтальное выравнивание для метки данных диаграммы.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Связь_ > verticalAlignment|Представляет вертикальное выравнивание метки данных диаграммы.|1,8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Отношения_ > границы|Представляет формат границы, включая цвет, линии и вес. Только для чтения.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Свойство_ > Автотекст|Представляет ли метки данных автоматически создавать соответствующий текст на основе контекста.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Свойство_ > numberFormat|Представляет код формата для метки данных.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Свойство_ > textOrientation|Представляет ориентации текста метки данных. Значение должно быть целое число, либо от -90 до 90 или 0 – 180 для вертикально ориентированного текста.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Связь_ > horizontalAlignment|Представляет горизонтальное выравнивание для метки данных диаграммы.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Связь_ > verticalAlignment|Представляет вертикальное выравнивание метки данных диаграммы.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Свойство_ > chartId|Получает идентификатор диаграммы, деактивирован.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Свойство_ > тип|Получает тип события.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в котором деактивирован диаграммы.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Свойство_ > chartId|Получает идентификатор диаграммы, удаляется с листа.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Свойство_ > тип|Получает тип события.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в которой удаляется диаграммы.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Отношения_ > источник|Получает источник события.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > height|Представляет высоту legendEntry в условных обозначениях диаграммы. Только для чтения.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > index|Представляет индекс legendEntry в условных обозначениях диаграммы. Только для чтения.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > слева|Представляет слева от диаграммы legendEntry. Только для чтения.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > в начало|Представляет в верхней части диаграммы legendEntry. Только для чтения.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > width|Представляет ширину legendEntry легенды диаграммы. Только для чтения.|1,8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Отношения_ > границы|Представляет формат границы, включая цвет, линии и вес. Только для чтения.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > height|Представляет значения высоты plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > insideHeight|Представляет значение insideHeight plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > insideLeft|Представляет значение insideLeft plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > insideTop|Представляет значение insideTop plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > insideWidth|Представляет значение insideWidth plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > слева|Представляет левой значение plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > в начало|Представляет максимального значения plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Свойство_ > width|Представляет значение ширины plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Отношения_ > формат|Представляет форматирования plotArea диаграммы. Только для чтения.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Отношения_ > позиции|Представляет положение plotArea.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Отношения_ > границы|Представляет атрибуты границы plotArea диаграммы. Только для чтения.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Отношения_ > заливки|Представляет формат заливки объекта, включая сведения о форматировании фона. Только для чтения.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > развертывания|Возвращает или задает значение развертывание круговой диаграммы или кольцевая фрагмента. Возвращает нуль (0), если нет без развертывания (Совет фрагмент — в центре круговой диаграммы). ReadWrite.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > firstSliceAngle|Возвращает или задает угол первого сектора круговой диаграммы или кольцевая диаграмма, в градусов (часовой с вертикального). Применяется только к круговая, объемных круговых и кольцевых диаграммах. Может быть в диапазоне от 0 до 360. ReadWrite|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > invertIfNegative|Значение true, если Microsoft Excel инвертирует шаблон в элементе, если он соответствует отрицательное значение. ReadWrite.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > пересекаются|Указывает расположение строки и столбцы. Может быть в диапазоне от -100 до 100. Применяется только к плоских диаграмм и гистограмм 2-D. ReadWrite.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > secondPlotSize|Возвращает или задает размер дополнительного раздела либо Вторичная круговая диаграмма или панели круговой диаграммы в процентах от размера основной круговой диаграммы. Может быть в диапазоне от 5 до 200. ReadWrite.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > varyByCategories|Значение true, если Microsoft Excel назначает разные цвета или узора к маркерам данных. Диаграмма должен содержать только один ряд. ReadWrite.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Отношения_ > axisGroup|Возвращает или задает группу для указанного ряда. ReadWrite|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Отношения_ > dataLabels|Представляет коллекцию всех dataLabels из серии. Только для чтения.|1,8|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Отношения_ > splitType|Возвращает или задает способ разделить два раздела Вторичная круговая диаграмма или панели круговой диаграммы. ReadWrite.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > backwardPeriod|Представляет число периодов, линия тренда расширяет назад.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > forwardPeriod|Представляет число периодов, линия тренда расширяет вперед.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > showEquation|Значение true, если формулу для линии тренда отображается на диаграмме.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > showRSquared|Значение true, если R-квадрат для линии тренда отображается на диаграмме.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Отношения_ > метки|Представляет метку линии тренда диаграммы. Только для чтения.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > Автотекст|Логическое значение, представляющее Если подписи линии тренда автоматически создает текст на основе контекста.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > формулы|Строковое значение, представляющее формула подпись диаграммы тренда нотации стиля A1.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > height|Возвращает высоту, в пунктах подпись линии тренда диаграммы. Только для чтения. NULL, если подпись диаграммы тренда не отображается. Только для чтения.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > слева|Представляет расстояние в пунктах от левого края диаграммы тренда метки для левого края области диаграммы. NULL, если подпись диаграммы тренда не отображается.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > numberFormat|Строковое значение, представляющее код формата для линии тренда метки.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > text|Строка, представляющая текст подписи линии тренда на диаграмме.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > textOrientation|Представляет ориентации текста диаграммы тренда метки. Значение должно быть целое число либо от -90 до 90 или 180 для вертикально ориентированного текста.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > в начало|Представляет расстояние в пунктах от верхнего края диаграммы тренда метки в верхней части области диаграммы. NULL, если подпись диаграммы тренда не отображается.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Свойство_ > width|Возвращает ширину в пунктах подпись линии тренда диаграммы. Только для чтения. NULL, если подпись диаграммы тренда не отображается. Только для чтения.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Отношения_ > формат|Представляет формат подписи линии тренда диаграммы. Только для чтения.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Связь_ > horizontalAlignment|Представляет горизонтальное выравнивание для диаграммы тренда метки.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Связь_ > verticalAlignment|Представляет вертикальное выравнивание подпись линии тренда диаграммы.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Отношения_ > границы|Представляет формат границы, включая цвет, линии и вес. Только для чтения.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Отношения_ > заливки|Представляет формат заливки метку линия тренда диаграммы. Только для чтения.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Связь_ > font|Представляет атрибуты шрифта (шрифт, размер шрифта, цвета, и т.д.) для элемента label линия тренда диаграммы. Только для чтения.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Свойство_ > fakeFileId|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Свойство_ > fileBase64|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Отношения_ > тип действия|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Свойство_ > формулы| Формула проверки пользовательских данных. Это создает специальные правила ввода, например, препятствующие дубликатов или ограничивать всего в диапазоне ячеек.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Свойство_ > id|Идентификатор DataPivotHierarchy. Только для чтения.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Свойство_ > name|Имя DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Свойство_ > numberFormat|Числовой формат DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Свойство_ > позиции|Положение DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Отношения_ > поля|Возвращает сводные поля, связанного с DataPivotHierarchy. Только для чтения.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Отношения_ > showAs|Определяет, должны ли отображаться данные как конкретных сводки.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Отношения_ > summarizeBy|Определяет, следует ли отображать все элементы DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Метод_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault)|Сброс DataPivotHierarchy значения по умолчанию.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Свойство_ > items|Коллекция объектов dataPivotHierarchy. Только для чтения.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Метод_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Добавляет PivotHierarchy текущей оси.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Получает DataPivotHierarchy по его имени или идентификатора.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection)|Возвращает DataPivotHierarchy по имени. Если DataPivotHierarchy не существует, возвращает значение null, object.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Метод_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Удаляет PivotHierarchy от текущей оси.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Свойство_ > ignoreBlanks|Игнорировать пустые ячейки: проверка данных не будет выполнена на пустые ячейки, по умолчанию используется значение true.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Свойство_ > допустимое|Представляет, если все значения являются допустимыми в соответствии с правилами проверки данных. Только для чтения.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Отношения_ > errorAlert|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Отношения_ > строки|Строки, когда пользователь выбирает ячейки.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Отношения_ > правила|Правила проверки данных, который содержит различные типы условия проверки данных.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Связь_ > type|Введите проверки данных, просматривать [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) для получения дополнительных сведений. Только для чтения.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Метод_ > [clear()](/javascript/api/excel/excel.datavalidation)|Очищает проверку данных из текущего диапазона.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Свойство_ > сообщения|Представляет предупреждающее сообщение об ошибке.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Свойство_ > showAlert|Определяет, следует ли отображать ошибки диалогового окна предупреждения или не в том случае, если пользователь вводит недопустимые данные. Значение по умолчанию: true.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Свойство_ > title|Представляет заголовок диалогового окна оповещения об ошибках.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Отношения_ > Стиль|Проверка данных представляет тип оповещения, можно найти [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) для получения дополнительных сведений.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Свойство_ > сообщения|Представляет сообщение приглашения.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Свойство_ > showPrompt|Определяет необходимость устанавливать в строке, если пользователь выбирает ячейку с помощью проверки данных.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Свойство_ > title|Представляет заголовок сообщения.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > настраиваемых|Условия проверки пользовательских данных.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > даты|Дата условия проверки данных.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > decimal|Условия проверки данных Decimal.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Связь_ > list|Условия проверки данных списка.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > textLength|Условия проверки данных TextLength.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > времени|Условия проверки данных времени.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Отношения_ > wholeNumber|WholeNumber условия проверки данных.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Свойство_ > formula1|Получает или задает Formula1, то есть минимальное значение или значение в зависимости от оператора.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Свойство_ > formula2|Получает или задает Formula2, то есть максимальное значение или значение в зависимости от оператора.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Отношения_ > оператор|Оператор, используемый для проверки данных.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Свойство_ > isEnableEvents {|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Отношения_ > тип действия|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Отношения_ > controlId|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Свойство_ > enableMultipleFilterItems|Определяет, следует ли разрешить несколько элементов фильтра.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Свойство_ > id|Идентификатор FilterPivotHierarchy. Только для чтения.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Свойство_ > name|Имя FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Свойство_ > позиции|Положение FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Отношение_ > fields|Возвращает сводные поля, связанного с FilterPivotHierarchy. Только для чтения.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Метод_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Сброс FilterPivotHierarchy значения по умолчанию.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Свойство_ > items|Коллекция объектов filterPivotHierarchy. Только для чтения.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Метод_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Добавляет PivotHierarchy текущей оси. При наличии других местах на строки, столбца или оси фильтра иерархии, он будет удален из этого расположения.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Получает FilterPivotHierarchy по его имени или идентификатора.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection)|Возвращает FilterPivotHierarchy по имени. Если FilterPivotHierarchy не существует, возвращает значение null, object.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Метод_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Удаляет PivotHierarchy от текущей оси.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Свойство_ > inCellDropDown|Раскрывающийся список в ячейке или не отображается, по умолчанию используется значение true.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Свойство_ > источник|Исходный список для проверки данных|1,8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Свойство_ > fakeFileId|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Отношения_ > тип действия|Передает дополнительные данные на стороне клиента, например, worksheetId для TableSelectionChangedEvent.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Свойство_ > id|Идентификатор сводных полей. Только для чтения.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Свойство_ > name|Имя сводных полей.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Свойство_ > showAllItems|Определяет, следует ли отображать все элементы сводных полей.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Отношения_ > элементов|Возвращает сводные поля, связанного с сводных полей. Только для чтения.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Отношения_ > промежуточные итоги|Промежуточные итоги сводных полей.|1,8|
|[сводных полей](/javascript/api/excel/excel.pivotfield)|_Метод_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Сортировка сводных полей. Если DataPivotHierarchy указан, затем сортировки будет применяться на его основе, если не сводных полей самого зависит сортировки.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Свойство_ > items|Коллекция объектов сводных полей. Только для чтения.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Получает PivotHierarchy по его имени или идентификатора.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection)|Возвращает PivotHierarchy по имени. Если PivotHierarchy не существует, возвращает значение null, object.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Свойство_ > id|Идентификатор PivotHierarchy. Только для чтения.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Свойство_ > name|Имя PivotHierarchy.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Отношение_ > fields|Возвращает сводные поля, связанного с PivotHierarchy. Только для чтения.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Свойство_ > items|Коллекция объектов pivotHierarchy. Только для чтения.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Получает PivotHierarchy по его имени или идентификатора.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection)|Возвращает PivotHierarchy по имени. Если PivotHierarchy не существует, возвращает значение null, object.|1,8|
|[элемент сводной таблицы](/javascript/api/excel/excel.pivotitem)|_Свойство_ > id|Идентификатор элемент сводной таблицы. Только для чтения.|1,8|
|[элемент сводной таблицы](/javascript/api/excel/excel.pivotitem)|_Свойство_ > isExpanded|Определяет, развернут ли для отображения дочерних элементов элемента или если он свернут, и дочерние элементы являются скрытыми.|1,8|
|[элемент сводной таблицы](/javascript/api/excel/excel.pivotitem)|_Свойство_ > name|Имя элемент сводной таблицы.|1,8|
|[элемент сводной таблицы](/javascript/api/excel/excel.pivotitem)|_Свойство_ > visible|Определяет, отображается ли элемент сводной таблицы или нет.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Свойство_ > items|Коллекция объектов элемент сводной таблицы. Только для чтения.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Получает PivotHierarchy по его имени или идентификатора.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection)|Возвращает PivotHierarchy по имени. Если PivotHierarchy не существует, возвращает значение null, object.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Свойство_ > showColumnGrandTotals|Значение true, если сводной таблицы, отчет отображает общие итоги для столбцов.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Свойство_ > showRowGrandTotals|Значение true, если сводной таблицы, отчет отображает общих итогов для строк.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Свойство_ > subtotalLocation|Это свойство показывает SubtotalLocationType всех полей в сводной таблице. Если поля имеют различные состояния, это будет null. Возможные значения: AtTop, AtBottom.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Отношения_ > layoutType|Это свойство показывает PivotLayoutType всех полей в сводной таблице. Если поля имеют различные состояния, это будет null.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Метод_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Возвращает диапазон, где находятся названия столбцов со сводными таблицами.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Метод_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Возвращает диапазон, где находятся значения данных со сводными таблицами.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout.md)|_Метод_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Возвращает диапазон область фильтра сводной таблицы.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Метод_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Возвращает диапазон, который существует со сводными таблицами, за исключением области фильтра.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Метод_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Возвращает диапазон, где находятся названия строк со сводными таблицами.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > columnHierarchies|Иерархии Pivot столбцов сводной таблицы. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > dataHierarchies|Иерархии сводных данных сводной таблицы. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > filterHierarchies|Иерархии Pivot фильтра сводной таблицы. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > иерархий|Иерархии Pivot сводной таблицы. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > макета|PivotLayout, описывающий макет и визуальной структуры со сводными таблицами. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношения_ > rowHierarchies|Иерархии Pivot строк сводной таблицы. Только для чтения.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Метод_ > [delete()](/javascript/api/excel/excel.pivottable)|Удаляет со сводными таблицами.|1,8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Метод_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|Добавление сводной таблицы на основе указанного источника данных и вставить его в верхнюю левую ячейку конечного диапазона.|1,8|
|[range](/javascript/api/excel/excel.range)|_Отношения_ > dataValidation|Возвращает объект проверки данных. Только для чтения.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Свойство_ > id|Идентификатор RowColumnPivotHierarchy. Только для чтения.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Свойство_ > name|Имя RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Свойство_ > позиции|Положение RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Отношение_ > fields|Возвращает сводные поля, связанного с RowColumnPivotHierarchy. Только для чтения.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Метод_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Сброс RowColumnPivotHierarchy значения по умолчанию.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Свойство_ > items|Коллекция объектов rowColumnPivotHierarchy. Только для чтения.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Метод_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Добавляет PivotHierarchy текущей оси. При наличии в другом месте в той строке иерархии столбец,|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Метод_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Возвращает количество иерархий pivot в коллекции.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Метод_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Получает RowColumnPivotHierarchy по его имени или идентификатора.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Метод_ > [getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Возвращает RowColumnPivotHierarchy по имени. Если RowColumnPivotHierarchy не существует, возвращает значение null, object.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Метод_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Удаляет PivotHierarchy от текущей оси.|1,8|
|[среда времени выполнения](/javascript/api/excel/excel.runtime)|_Свойство_ > enableEvents|Переключение события JavaScript в текущей taskpane или контента надстройки.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Отношения_ > baseField|Базовый сводных полей будет создана вычислений ShowAs, если это возможно на основании типа ShowAsCalculation, в противном случае значение null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Отношения_ > baseItem|Базовый элемент для вычисления ShowAs на, если это возможно на основании типа ShowAsCalculation, в противном случае значение null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Отношения_ > вычислений|Расчет ShowAs для сводных полей данных.|1,8|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > autoIndent|Указывает, если текст автоматический отступ, если для выравнивания текста в ячейку в равномерного распределения.|1,8|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > textOrientation|Ориентация текста для стиля.|1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > автоматического|Если автоматический задано значение true, то все остальные значения будет игнорироваться при задании промежуточных итогов.|1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > среднее| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > count| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > countNumbers| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > max| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > мин.| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > продукта| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > standardDeviation| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > standardDeviationP| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > сумм| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > отклонение| |1,8|
|[промежуточные итоги](/javascript/api/excel/excel.subtotals)|_Свойство_ > varianceP| |1,8|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > legacyId|Возвращает числовой идентификатор. Только для чтения.|1,8|
|[workbook](/javascript/api/excel/excel.workbook)|_Свойство_ > только для чтения|Значение true, если книга открыта в режиме только для чтения. Только для чтения.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Свойство_ > id|Возвращает значение, уникальным образом идентифицирующее объект WorkbookCreated. Только для чтения.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Метод_ > [open()](/javascript/api/excel/excel.workbookcreated)|Откройте книгу.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > showGridlines|Получает или задает флаг линии сетки рабочего листа.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > showHeadings|Получает или задает заголовки флаг рабочего листа.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Свойство_ > тип|Получает тип события.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, который вычисляется.|1,8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Новые возможности Excel JavaScript API 1.7

Возможности Excel JavaScript API требование set 1.7 включают API-интерфейсы для диаграмм, события, таблицы, диапазоны, свойства документа, именованные элементы, параметры защиты и стили.

### <a name="customize-charts"></a>Настройка диаграмм

С новой диаграммы API-интерфейсы можно создать дополнительные диаграммы типа, добавьте рядов данных диаграммы, установка заголовка диаграммы, добавьте заголовок оси, добавьте цену, Добавление линии тренда с среднее, линии тренда линейная и многое другое. Ниже приведены некоторые примеры:

* Диаграмма, ось — получения, задания, форматирование и удалить подразделение ось, label и title на диаграмме.
* Рядов диаграммы - добавьте, Установка и удаление рядов в диаграмме.  Изменение маркеров рядов, заказы построения и изменения размера.
* Диаграмма тренда — Добавление, получение и форматирование линии тренда на диаграмме.
* Условных обозначениях диаграммы - формат шрифт легенды на диаграмме.
* Диаграмма точка - цвет точки диаграммы set.
* Диаграмма, заголовок подстроки - получения и задания подстроки заголовок для диаграммы.
* Тип диаграммы — параметр, чтобы создать дополнительные типы диаграмм.

### <a name="events"></a>События

Происходит Excel событий, которые предоставляют API-интерфейсы разнообразные обработчики событий, которые позволяют надстройки для автоматического запуска указанной функции при определенных событий. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. Список событий, которые в настоящее время доступны [событий с помощью интерфейса API JavaScript Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)см.

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Настройка внешнего вида таблиц и диапазоны

С помощью новых интерфейсов API, можно настроить внешний вид таблицы несколькими способами:

* Закрепление областей на отображение столбцов и строк при прокрутке в рабочем листе. К примеру при первой строки в электронной таблице содержит заголовки, может Закрепить эту строку так, чтобы заголовки столбцов, останутся видимыми при прокрутке листа вниз.
* Измените цвет ярлычка листа.
* Добавьте заголовки рабочего листа.


Можно настроить внешний вид диапазонов несколькими способами:

* Задать стиль ячейки для диапазона для обеспечения всем ячейкам в диапазоне должны иметь согласованное форматирование. Стиль ячейки — определенный набор параметров, таких как шрифты и размеры шрифтов, форматы телефонных номеров, ячейки границы и заливка ячеек форматирования. Используйте любой из стилей встроенных ячеек в Excel или создать собственный стиль настраиваемые ячейки.
* Настройка ориентации текста для диапазона.
* Добавление или изменение гиперссылки в диапазоне, который связывает в другое место в книге или на внешний носитель.

### <a name="manage-document-properties"></a>Управление свойствами документа

С помощью API-интерфейсы свойства документа, можно получить доступ к встроенных свойств документа и создание и управление настраиваемых свойств документов для хранения состояния книги и диск рабочих процессов и бизнес-логику.

### <a name="copy-worksheets"></a>Скопируйте листов

С помощью копии листа API-интерфейсы, можно скопировать данные и формат с одного листа для нового листа в той же книге и сокращения объема необходимости передачи данных.

### <a name="handle-ranges-with-ease"></a>Обработка диапазонов с легкостью

С помощью различных диапазона API-интерфейсы, можно выполнить действия, такие как get окружающих региона, получить размер диапазона и многое другое. Эти API-интерфейсы следует сделать намного эффективнее, задач, таких как выполнение различных операций диапазон и назначение адресов.

Кроме того:

* Параметры защиты книг и листов - используйте эти интерфейсы API для защиты данных в лист и структура книги.
* Обновление именованного элемента — используйте этот интерфейс API для обновления к именованному элементу.
* Получение активной ячейки — используйте этот интерфейс API для получения активной ячейки книги.

|Объект| Что нового| Описание|Набор обязательных элементов|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > chartType|Представляет тип диаграммы. Возможные значения: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, и т.д...|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > id|Уникальный идентификатор диаграммы. Только для чтения.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > showAllFieldButtons|Представляет, следует ли отображать все кнопки полей в сводной таблице.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Отношения_ > границы|Представляет границы формат области диаграммы, включая цвет, линии и вес. Только для чтения.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Метод_ > getItem (тип: string, групповой: строка)|Возвращает конкретную ось, указанный тип и группы.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > axisBetweenCategories|Представляет ли ось пересечения оси категорий.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > axisGroup|Представляет группу для указанной оси. Только для чтения. Возможные значения: основной, дополнительный.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > categoryType|Возвращает или задает тип оси категории. Возможные значения: автоматическое обновление, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > пересечение|Представляет указанный ось, пересечения другой оси. Возможные значения: автоматическое, максимальный, минимальный, настраиваемые.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > crossesAt|Представляет указанный ось, где другой оси пересечение в. Только для чтения. Параметр имеет значение этого свойства следует использовать метод SetCrossesAt(double). Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > customDisplayUnit|Представляет значения отображения настраиваемых оси. Только для чтения. Чтобы задать это свойство, используйте метод SetCustomDisplayUnit(double). Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > displayUnit|Представляет деления оси. Возможные значения: None, сотни, тысяч, TenThousands, HundredThousands, миллионов, TenMillions, HundredMillions, миллиардов, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > height|Представляет высота в пунктах оси диаграммы. NULL, если оси не видны. Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > слева|Представляет расстояние в пунктах от левого края оси слева от области диаграммы. NULL, если оси не видны. Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > logBase|Представляет основание логарифма при использовании устранить шкал.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > reversePlotOrder|Представляет ли Нанесение точек данных из последней первого с помощью Microsoft Excel.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > scaleType|Представляет тип шкалы оси значения. Возможные значения: линейная, устранить.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > showDisplayUnitLabel|Указывает, находится ли видимым подпись оси отображения подразделения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > tickLabelSpacing|Представляет число категорий или рядов между меток делений. Может быть значение от 1 до 31999 или пустая строка для автоматической настройки. Возвращаемое значение всегда является номером.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > tickMarkSpacing|Представляет число категорий или рядов между деления.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > в начало|Представляет расстояние в пунктах от верхнего края оси в верхнюю часть области диаграммы. NULL, если оси не видны. Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > тип|Представляет тип оси. Только для чтения. Возможные значения: недопустимый, категории, значение, серии.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > visible|Логическое значение представляет видимости оси.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Свойство_ > width|Представляет ширину в пунктах оси диаграммы. NULL, если оси не видны. Только для чтения.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > baseTimeUnit|Возвращает или задает единицу для оси указанной категории.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > majorTickMark|Представляет тип основные деления для указанной оси.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > majorTimeUnitScale|Возвращает или задает значение масштаба основные единицы для оси категорий при CategoryType задано значение шкала времени.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > minorTickMark|Представляет тип вспомогательные деления для указанной оси.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > minorTimeUnitScale|Возвращает или задает значение масштаба Вспомогательные единицы для оси категорий при CategoryType задано значение шкала времени.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Отношения_ > tickLabelPosition|Представляет положение меток делений на указанной оси.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Метод_ > setCategoryNames(sourceData: Range)|Устанавливает все имена категорий для указанной оси.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Метод_ > setCrossesAt(value: double)|Задайте указанной оси, где другой оси пересечение в.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Метод_ > setCustomDisplayUnit(value: double)|Задает деления оси пользовательское значение.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Свойство_ > color|HTML-код цвета, представляющее цвет границы на диаграмме.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Свойство_ > вес|Представляет Вес границы в точках.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Отношения_ > линии|Представляет тип линии границы.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > позиции|Значение DataLabelPosition, которое представляет положение метки данных. Возможные значения: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > разделителя групп разрядов|Строка, представляющая разделитель, используемый для подписи данных на диаграмме.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showBubbleSize|Логическое значение, которое указывает, отображается ли размер пузырьков с метками данных.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showCategoryName|Логическое значение, которое указывает, отображается ли имя для категории меток данных.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showLegendKey|Логическое значение, которое указывает, отображаются ли условные обозначения для меток данных.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showPercentage|Логическое значение, которое указывает, отображается ли процентное соотношение меток данных.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showSeriesName|Логическое значение, которое указывает, отображается ли имя ряда для меток данных.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Свойство_ > showValue|Логическое значение, которое указывает, отображается ли значение метки данных.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Свойство_ > height|Представляет высоту легенды диаграммы.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Свойство_ > слева|Представляет слева от условных обозначениях диаграммы.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Свойство_ > showShadow|Представляет Если теневая легенде диаграммы.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Свойство_ > в начало|Представляет в верхней части условных обозначениях диаграммы.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Свойство_ > width|Представляет ширину легенды диаграммы.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Отношения_ > legendEntries|Представляет коллекцию legendEntries в легенде. Только для чтения.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Свойство_ > visible|Представляет видимым записи легенды диаграммы.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Свойство_ > items|Коллекция объектов chartLegendEntry. Только для чтения.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Метод_ > getCount()|Возвращает число legendEntry в коллекции.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Метод_ > getItemAt(index: number)|Возвращает legendEntry по указанному индексу.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Свойство_ > hasDataLabel|Представляет ли точка данных имеет datalabel. Неприменимо для предоставления диаграмм.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Свойство_ > markerBackgroundColor|Выберите представление цвета кода HTML цвета фона маркер данных. Пример: #FF0000 обозначает красный.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Свойство_ > markerForegroundColor|Выберите представление цвета кода HTML цвета переднего плана маркер данных. Пример: #FF0000 обозначает красный.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Свойство_ > markerSize|Представляет размер маркера точки данных.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Свойство_ > markerStyle|Представляет стиль маркера точки данных диаграммы. Возможные значения: недопустимый, автоматически, нет, квадрат, ромб, треугольник, X, Star, Dot, тире, обведите, а также, рисунок.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Отношения_ > dataLabel|Возвращает метки данных точки диаграммы. Только для чтения.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Отношения_ > границы|Представляет формат границы точки данных диаграммы, включая цвет, стиль и вес сведения. Только для чтения.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > chartType|Представляет тип диаграммы для ряда. Возможные значения: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, и т.д...|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > doughnutHoleSize|Представляет размер кольцевых отверстий рядов диаграммы.  Допустимо только в doughnutExploded и кольцевых диаграммах.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > отфильтровано|Логическое значение, представляющее Если серии фильтруется или нет. Неприменимо для предоставления диаграмм.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > gapWidth|Представляет ширину разрывов рядов диаграммы.  Допустимо только на диаграмм и гистограмм, а также|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > hasDataLabels|Логическое значение, представляющее серия имеет метки данных или нет.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > markerBackgroundColor|Представляет цвет фона маркеры рядов диаграммы.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > markerForegroundColor|Представляет цвет переднего плана маркеры рядов диаграммы.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > markerSize|Представляет размер маркера рядов диаграммы.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > markerStyle|Представляет стиль маркера рядов диаграммы. Возможные значения: недопустимый, автоматически, нет, квадрат, ромб, треугольник, X, Star, Dot, тире, обведите, а также, рисунок.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > plotOrder|Представляет порядка отображения рядов диаграммы в группе диаграммы.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > showShadow|Логическое значение, представляющее серия имеет теневой или нет.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Свойство_ > плавный|Логическое значение, представляющее если ряд является легко или нет. Только для графиков и точечных диаграмм.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Отношения_ > dataLabels|Представляет коллекцию всех dataLabels из серии. Только для чтения.|ApiSet.InProgressFeatures.ChartingAPI|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Отношения_ > линии тренда|Представляет коллекцию линии тренда в серии. Только для чтения.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Метод_ > delete()|Удаляет рядов диаграммы.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Метод_ > setBubbleSizes(sourceData: Range)|Настройка размеров пузырьковой для ряда диаграммы. Работает только для пузырьковых диаграмм.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Метод_ > setValues(sourceData: Range)|Задайте значения для ряда диаграммы. Точечные диаграммы это означает значения оси Y.|1.7|
|[ChartSeries ряд](/javascript/api/excel/excel.chartseries)|_Метод_ > setXAxisValues(sourceData: Range)|Задайте значения X оси для ряда диаграммы. Работает только для точечных диаграмм.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Метод_ > Добавить (имя: string, индексирование: номер)|Добавление новой серии в коллекцию.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > height|Возвращает высоту, в пунктах заголовок диаграммы. Только для чтения. NULL, если заголовок диаграммы не видны. Только для чтения.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > horizontalAlignment|Представляет горизонтальное выравнивание для заголовка диаграммы. Возможные значения: Center, ЛЕВСИМВ ширине, распределенных, справа.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > слева|Представляет расстояние в пунктах от левого края название диаграммы с левого края области диаграммы. NULL, если заголовок диаграммы не видны.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > позиции|Представляет положение заголовка диаграммы. Возможные значения: вверх, автоматический, вниз, справа, слева.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > showShadow|Представляет логическое значение, которое определяет, имеет ли заголовка диаграммы тени.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > textOrientation|Представляет ориентации текста заголовка диаграммы. Значение должно быть целое число либо от -90 до 90 или 180 для вертикально ориентированного текста.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > в начало|Представляет расстояние в пунктах от верхнего края диаграммы заголовка в верхней части области диаграммы. NULL, если заголовок диаграммы не видны.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > ВертикальноеВыравнивание|Представляет вертикальное выравнивание заголовка диаграммы. Возможные значения: центр, нижней, Top, ширине, распределенных.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Свойство_ > width|Возвращает ширину в пунктах заголовок диаграммы. Только для чтения. NULL, если заголовок диаграммы не видны. Только для чтения.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Метод_ > setFormula(formula: string)|Задает строковое значение, представляющее формула нотации стиля A1 заголовка диаграммы.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Отношения_ > границы|Представляет формат границы название диаграммы, включая цвет, линии и вес. Только для чтения.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > обратной|Представляет число периодов, линия тренда расширяет назад.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > displayEquation|Значение true, если формулу для линии тренда отображается на диаграмме.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > displayRSquared|Значение true, если R-квадрат для линии тренда отображается на диаграмме.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > вперед|Представляет число периодов, линия тренда расширяет вперед.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > отрезка|Представляет значение отрезка линии тренда. Может быть присвоено числовое значение или пустая строка (для автоматического значения). Возвращаемое значение всегда является номером.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > movingAveragePeriod|Представляет собой период диаграммы тренда только для линии тренда с типом MovingAverage.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > name|Представляет имя линии тренда. Может быть присвоено значение string, или можно задать значения автоматического представляет значение null. Возвращаемое значение всегда является строкой|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > polynomialOrder|Представляет приоритет диаграммы тренда только для линии тренда с полиномиальной типа.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Свойство_ > тип|Представляет тип диаграммы тренда. Возможные значения: линейная, экспоненциальное, устранить, MovingAverage, полинома, питания.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Отношения_ > формат|Представляет форматирование диаграммы тренда. Только для чтения.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Метод_ > delete()|Удалите объект линия тренда.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Свойство_ > items|Коллекция объектов chartTrendline. Только для чтения.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Метод_ > add(type: string)|Добавляет новый линия тренда линия тренда коллекцию.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Метод_ > getCount()|Возвращает число линии тренда в коллекции.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Метод_ > getItem(index: number)|Получите объект линия тренда по индексу, который является порядок вставки в массиве элементов.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Отношения_ > строки|Представляет форматирование линий диаграммы. Только для чтения.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Свойство_ > key|Возвращает ключ настраиваемого свойства. Только для чтения. Только для чтения.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Свойство_ > тип|Возвращает значение типа настраиваемого свойства. Только для чтения. Только для чтения. Возможные значения: номер, Boolean, дата, строка, число с плавающей запятой.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Свойство_ > value|Возвращает или задает значение настраиваемого свойства.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Метод_ > delete()|Удаляет настраиваемое свойство.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Свойство_ > items|Коллекция объектов customProperty. Только для чтения.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Метод_ > Добавить (ключ: строковое значение: объектов)|Создает или задает настраиваемое свойство.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Метод_ > deleteAll()|Удаляет все настраиваемые свойства в коллекции.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Метод_ > getCount()|Возвращает количество настраиваемых свойств.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Метод_ > getItem(key: string)|Возвращает объект custom property по ключу, нечувствительному к регистру. Выдает ошибку, если настраиваемое свойство не существует.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Метод_ > getItemOrNullObject(key: string)|Возвращает объект custom property по ключу, нечувствительному к регистру. Возвращает объект null, если настраиваемое свойство не существует.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Свойство_ > items|Коллекция объектов подключение данных. Только для чтения.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Метод_ > refreshAll()|Обновляет все подключения к данным в семействе сайтов.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > author|Получает или задает автора книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > category|Получает или задает категорию рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > comments|Получает или задает комментарии рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > company|Получает или задает компании рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > keywords|Получает или задает ключевые слова из рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > lastAuthor|Получает автор книги. Только для чтения. Только для чтения.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > manager|Получает или задает диспетчер рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > revisionNumber|Получает номер редакции книги. Только для чтения.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > subject|Получает или задает тему рабочей книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Свойство_ > title|Получает или задает заголовок книги.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Связь_ > creationDate|Получает дату создания книги. Только для чтения. Только для чтения.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Отношения_ > настраиваемых|Получает коллекцию настраиваемых свойств книги. Только для чтения. Только для чтения.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Свойство_ > формулы|Получает или задает формулу именованный элемент.  Формула всегда начинается со знака «=».|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Отношения_ > arrayValues|Возвращает объект, содержащий значения и типы именованный элемент. Только для чтения.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Свойство_ > типы|Представляет типы для каждого элемента в массиве именованный элемент только для чтения. Возможные значения: неизвестно, пустой, строка, целое число, двойное, Boolean, ошибка.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Свойство_ > values|Представляет значения каждого элемента в массиве именованный элемент. Только для чтения.|1.7|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > isEntireColumn|Представляет, если текущий диапазон целый столбец. Только для чтения.|1.7|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > isEntireRow|Представляет, если текущий диапазон всю строку. Только для чтения.|1.7|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > numberFormatLocal|Представляет код числового формата Excel для указанного диапазона как строку в языке пользователя.|1.7|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > style|Представляет стиль текущий диапазон. Это возвращает null или строку.|1.7|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getAbsoluteResizedRange (numRows: номер numColumns: номер)|Получает объект Range с одной левый верхний угол как текущий объект Range, но с указанного номера строк и столбцов.|1.7|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getImage()|Отображает диапазон как изображение кодировке base64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getSurroundingRegion()|Возвращает объект Range, который представляет окружающих регион левый верхний угол в этот диапазон. Окружающих область — диапазон, в любом сочетании пустые строки и пустые столбцы, относящиеся к этот диапазон.|1.7|
|[range](/javascript/api/excel/excel.range)|_Метод_ > showCard()|Отображает карточку для активной ячейки, если он имеет значение Форматированный контент.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Свойство_ > textOrientation|Получает или задает ориентацию текста всех ячеек в диапазоне.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Свойство_ > useStandardHeight|Определяет, если высота строки объекта Range равно Стандартная высота листа.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Свойство_ > useStandardWidth|Определяет, если columnwidth объекта Range — это стандартные ширину листа.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Свойство_ > address|Представляет конечный URL-адрес гиперссылки.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Свойство_ > документа.|Представляет документ. целевой объект гиперссылки.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Свойство_ > подсказки|Представляет строку, отображаемую при наведении указателя на гиперссылку.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Свойство_ > textToDisplay|Представляет строку, которая отображается в в верхний левый большинство ячейки в диапазоне.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > addIndent|Указывает, если текст автоматический отступ, если для выравнивания текста в ячейку в равномерного распределения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > autoIndent|Указывает, если текст автоматический отступ, если для выравнивания текста в ячейку в равномерного распределения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > builtIn|Указывает, является ли стиль встроенных стилей. Только для чтения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > formulaHidden|Указывает, если формула будет скрыта при использовании защищенного рабочего листа.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > horizontalAlignment|Представляет горизонтальное выравнивание для стиля. Возможные значения: Общие, слева, центр, справа, заливки, ширине, CenterAcrossSelection, распределенных.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includeAlignment|Указывает, если стиль включает в себя свойства AutoIndent, HorizontalAlignment, ВертикальноеВыравнивание, WrapText, IndentLevel и TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includeBorder|Указывает, если стиль включает в себя свойства границы цвет, ColorIndex (en), линии и вес.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includeFont|Указывает, если стиль включает в себя свойства шрифта фона, полужирный, цвет, ColorIndex (en), FontStyle, курсив, имя, размер, зачеркивание, подстрочный знак, надстрочный знак и подчеркивание.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includeNumber|Указывает, если стиль включает в себя свойства формат числа.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includePatterns|Указывает, если стиль включает в себя цвета ColorIndex (en), InvertIfNegative, шаблон, PatternColor и PatternColorIndex внутренних свойств.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > includeProtection|Указывает, если стиль включает в себя свойства защиты FormulaHidden и Защищаемая ячейка.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > indentLevel|Целое число от 0 до 250 знаков, которое указывает, уровень отступа для стиля.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > locked|Указывает, если объект заблокирован, когда лист защищен.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > name|Имя стиля. Только для чтения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > numberFormat|Код формата числовой формат для стиля.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > numberFormatLocal|Код локализованном формате числовой формат для стиля.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > ориентации|Ориентация текста для стиля.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > направление чтения|Порядок чтения для стиля. Возможные значения: RightToLeft контекстного LeftToRight,.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > shrinkToFit|Указывает, если текст автоматически сжимается для соответствия ширине необходимые столбцы.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > textOrientation|Ориентация текста для стиля.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > ВертикальноеВыравнивание|Представляет вертикальное выравнивание для стиля. Возможные значения: вверху, центр, внизу, ширине, распределенных.|1.7|
|[style](/javascript/api/excel/excel.style)|_Свойство_ > wrapText|Указывает, если Microsoft Excel переносит текст в объект.|1.7|
|[style](/javascript/api/excel/excel.style)|_Отношения_ > границы|Границы коллекцию объектов четыре границы, которые представляют стиль четыре границы. Только для чтения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Отношения_ > заливки|Заливку стиля. Только для чтения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Связь_ > font|Объект Font, представляющий шрифта, стиля. Только для чтения.|1.7|
|[style](/javascript/api/excel/excel.style)|_Метод_ > delete()|Удаляет этот стиль.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Свойство_ > items|Коллекция объектов стиля. Только для чтения.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Метод_ > add(name: string)]|Добавляет в коллекцию новый стиль.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Метод_ > getItem(name: string)|Получает стиль по имени.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > address|Получает адрес, который представляет область измененные таблицы на конкретный лист.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > changeType|Получает тип изменения, представляющий запуска события Changed. Возможные значения: другим пользователям, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > источник|Получает источник события. Возможные значения: локального, удаленного.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > идентификатор таблицы|Получает идентификатор таблицы, в которой данные изменены.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в которой данные изменены.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Свойство_ > address|Получает адреса диапазона, представляющий область в таблице на конкретный лист.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Свойство_ > isInsideTable|Указывает, если выделение находится в таблице, адрес будет использовать невозможно, если IsInsideTable имеет значение false.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Свойство_ > идентификатор таблицы|Получает идентификатор таблицы, в котором изменен выделение.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в котором изменен выделение.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Свойство_ > name|Получает имя книги. Только для чтения.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношения_ > dataConnections|Обновляет все подключения к данным в книге. Только для чтения.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Связь_ > properties|Получает свойства книги. Только для чтения.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношение_ > protection|Возвращает объект Защита книги для книги. Только для чтения.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношения_ > стилей|Представляет коллекцию стилей, связанное с книгой. Только для чтения.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Метод_ > getActiveCell()|Получает активную ячейку в книге.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Свойство_ > protected|Указывает, если книга защищена. Только для чтения. Только для чтения.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Метод_ > protect(password: string)|Обеспечивает защиту книги. Не выполняется, если книга защищена.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Метод_ > unprotect(password: string)|Чтобы на время защиту книги.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > линии сетки|Получает или задает флаг линии сетки рабочего листа.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > заголовков|Получает или задает заголовки флаг рабочего листа.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > showHeadings|Получает или задает заголовки флаг рабочего листа.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > standardHeight|Возвращает высоту по умолчанию все строки в рабочем листе в пунктах. Только для чтения.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > standardWidth|Возвращает или задает ширину standard (по умолчанию) всех столбцов на листе.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Свойство_ > tabColor|Получает или задает цвет ярлычка листа.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Отношения_ > freezePanes|Получает объект, который можно использовать для работы с закрепление областей в рабочем листе только для чтения.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > копии (меню: WorksheetPositionType relativeTo: лист)|Скопируйте листа и поместите его в заданной позиции. Возвращает копии листа.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getRangeByIndexes (startRow: число, startColumn: число, rowCount: число, columnCount: номер)|Метод  возвращает объект диапазона, начинающегося с определенных строки и столбца и занимающего определенное количество строк и столбцов.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа активации.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Свойство_ > источник|Получает источник события. Возможные значения: локального, удаленного.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, который добавляется в книгу.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Свойство_ > address|Получает адреса диапазона, представляющий измененные области определенного рабочего листа.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Свойство_ > changeType|Получает тип изменения, представляющий запуска события Changed. Возможные значения: другим пользователям, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Свойство_ > источник|Получает источник события. Возможные значения: локального, удаленного.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в которой данные изменены.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Свойство_ > worksheetId|Получает идентификатор лист с деактивирован.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Свойство_ > источник|Получает источник события. Возможные значения: локального, удаленного.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, который будет удален из рабочей книги.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > freezeAt (frozenRange: диапазон или строки)|Задает остановленных ячеек в представлении активного листа.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > freezeColumns(count: number)|Закрепление первого столбцов листа на месте.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > freezeRows(count: number)|Закрепление верхней строк листа на месте.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > getLocation()|Получает диапазон с описанием остановленных ячеек в представлении активного листа.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > getLocationOrNullObject()|Получает диапазон с описанием остановленных ячеек в представлении активного листа.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Метод_ > unfreeze()|Удаляет все закрепление областей в рабочем листе.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowEditObjects|Представляет параметр защиты листа, разрешающий редактирование объектов.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowEditScenarios|Представляет параметр защиты листа, разрешающий изменение сценариев.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Отношения_ > selectionMode|Представляет параметр Защита листа Выбор режима.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Свойство_ > address|Получает адреса диапазона, представляющий область определенного рабочего листа.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Свойство_ > тип|Получает тип события. Возможные значения: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Свойство_ > worksheetId|Получает идентификатор листа, в котором изменен выделение.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Новые возможности Excel JavaScript API 1.6 

### <a name="conditional-formatting"></a>Условное форматирование

В этой статье рассматриваются условное форматирование диапазона. Разрешает условное форматирование следующих типов:

* Цветовая шкала
* Гистограмма
* Набор значков
* Пользовательский

Кроме того:

* Возвращает диапазон, который применяется к условное форматирование. 
* Удаление условного форматирования. 
* Предоставляет возможность приоритет и stopifTrue. 
* Получение полной коллекции условного форматирования для определенного диапазона. 
* Полное удаление условного форматирование в указанном диапазоне. 

|Объект| Что нового| Описание|Набор обязательных элементов|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Метод_ > suspendApiCalculationUntilNextSync()|Приостанавливает вычисление до вызова следующего "context.sync()". После этого за пересчет книги и распространение всех зависимостей несет ответственность разработчик.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Отношения_ > формат|Возвращает объект формата, который содержит шрифт, заливку, границы и другие свойства условного форматирования. Только для чтения.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Отношения_ > правила|Представляет объект Rule в этом условном форматировании.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Свойство_ > threeColorScale|Если вы укажете значение true, цветовая шкала будет иметь три точки (минимальная, средняя, максимальная), в противном случае она будет иметь две точки (минимальная, максимальная). Только для чтения.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Отношение_ > criteria|Условия цветовой шкалы. Средняя точка является необязательной при использовании цветовой шкалы с двумя точками.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Свойство_ > formula1|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Свойство_ > formula2|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Свойство_ > operator|Оператор условного форматирования текста. Возможные значения: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Отношения_ > максимальное|Условие цветовой шкалы "максимальная точка".|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Отношения_ > среднее|Условие цветовой шкалы "средняя точка", если используется трехцветная цветовая шкала.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Отношения_ > минимальные|Условие цветовой шкалы "минимальная точка".|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Свойство_ > color|HTML-код цвета цветовой шкалы. Например, #FF0000 обозначает красный.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Свойство_ > формулы|Число, формула или значение null (если указан тип LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Свойство_ > тип|На чем должна основываться условная формула значка. Возможные значения: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Свойство_ > borderColor|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Свойство_ > fillColor|HTML-код, представляющий цвет заливки в формате #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Свойство_ > matchPositiveBorderColor|Указывает, имеет ли отрицательная гистограмма тот же цвет границы, что и положительная.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Свойство_ > matchPositiveFillColor|Указывает, имеет ли отрицательная гистограмма тот же цвет заливки, что и положительная.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Свойство_ > borderColor|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Свойство_ > fillColor|HTML-код, представляющий цвет заливки в формате #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Свойство_ > gradientFill|Указывает, имеет ли гистограмма градиент.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Свойство_ > формулы|Формула, с помощью которой при необходимости оценивается правило гистограммы.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Свойство_ > тип|Тип правила для гистограммы. Возможные значения: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Свойство_ > id|Приоритет условное форматирование в пределах текущего ConditionalFormatCollection. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Свойство_ > приоритет|Приоритет (или индекс) в коллекции условного форматирования, в котором оно в настоящее время существует. Изменение этого параметра также|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Свойство_ > stopIfTrue|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Свойство_ > тип|Тип условного форматирования. Одновременно можно задать только один. Только для чтения. Только для чтения. Возможные значения: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > cellValue|Возвращает свойства условного форматирования по значению ячейки, если используется условное форматирование CellValue. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > cellValueOrNullObject|Возвращает свойства условного форматирования по значению ячейки, если используется условное форматирование CellValue. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > colorScale|Возвращает свойства условного форматирования ColorScale, если используется условное форматирование ColorScale. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > colorScaleOrNullObject|Возвращает свойства условного форматирования ColorScale, если используется условное форматирование ColorScale. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > настраиваемых|Возвращает свойства специального условного форматирования, если используется специальное условное форматирование. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > customOrNullObject|Возвращает свойства специального условного форматирования, если используется специальное условное форматирование. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > dataBar|Возвращает свойства гистограммы, если текущее условное форматирование — гистограмма. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > dataBarOrNullObject|Возвращает свойства гистограммы, если текущее условное форматирование — гистограмма. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > iconSet|Возвращает свойства условного форматирования IconSet, если используется условное форматирование IconSet. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > iconSetOrNullObject|Возвращает свойства условного форматирования IconSet, если используется условное форматирование IconSet. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > предварительно|Возвращает условное форматирование по готовым условиям, например свойства above averagebelow averageunique valuescontains blanknonblankerrornoerror. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > presetOrNullObject|Возвращает условное форматирование по готовым условиям, например свойства above averagebelow averageunique valuescontains blanknonblankerrornoerror. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > textComparison|Возвращает свойства условного форматирования по определенному тексту, если используется текстовое условное форматирование. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > textComparisonOrNullObject|Возвращает свойства условного форматирования по определенному тексту, если используется текстовое условное форматирование. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > сверху вниз|Возвращает свойства условного форматирования TopBottom, если используется условное форматирование TopBottom. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Отношения_ > topBottomOrNullObject|Возвращает свойства условного форматирования TopBottom, если используется условное форматирование TopBottom. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Метод_ > delete()|Удаляет это условное форматирование.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Метод_ > getRange()|Возвращает диапазон, к которому применяется условное форматирование, или объект null, если диапазон является непрерывным. Только для чтения.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Метод_ > getRangeOrNullObject()|Возвращает диапазон, к которому применяется условное форматирование, или объект null, если диапазон является непрерывным. Только для чтения.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Свойство_ > items|Коллекция объектов conditionalFormat. Только для чтения.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Метод_ > add(type: string)|Добавляет новое условное форматирование в коллекцию с наивысшим приоритетом.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Метод_ > clearAll()|Полное удаление условного форматирование в указанном диапазоне.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Метод_ > getCount()|Возвращает количество условных форматов в книге. Только для чтения.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Метод_ > getItem(id: string)|Возвращает условного форматирования с указанным идентификатором.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Метод_ > getItemAt(index: number)|Возвращает условное форматирование по индексу.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Свойство_ > формулы|Формула, с помощью которой при необходимости оценивается правило условного форматирования.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Свойство_ > formulaLocal|Формула, с помощью которой при необходимости оценивается правило условного форматирования на языке пользователя.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Свойство_ > formulaR1C1|Формула, с помощью которой при необходимости оценивается правило условного форматирования в формате R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Свойство_ > формулы|Число или формула в зависимости от типа.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Свойство_ > operator|Значение GreaterThan или GreaterThanOrEqual для каждого типа правила условного форматирования Icon. Возможные значения: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Отношения_ > customIcon|Специальный значок для текущего условия, если он отличается от набора значков по умолчанию, в противном случае возвращается значение null.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Связь_ > type|На чем должна основываться условная формула значка.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Свойство_ > условие|Условие условного форматирования. Возможные значения: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Свойство_ > color|HTML-код, представляющий цвет линии границы в виде #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Свойство_ > id|Представляет идентификатор границы. Только для чтения. Возможные значения: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Свойство_ > sideIndex|Постоянное значение, указывающее определенную сторону границы. Только для чтения. Возможные значения: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Свойство_ > style|Одна из констант стиля линии, определяющая стиль линии границы. Возможные значения: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Свойство_ > count|Количество объектов границы в коллекции. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Свойство_ > items|Коллекция объектов conditionalRangeBorder. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Отношения_ > внизу|Возвращает верхнюю границу. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Отношения_ > слева|Возвращает верхнюю границу. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Отношения_ > вправо|Возвращает верхнюю границу. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Отношения_ > в начало|Возвращает верхнюю границу. Только для чтения.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Метод_ > getItem(index: string)|Возвращает объект границы, используя его имя.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Метод_ > getItemAt(index: number)|Возвращает объект границы, указанный по индексу.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Свойство_ > color|HTML-код, представляющий цвет заливки в формате #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Метод_ > clear()|Удаляет заливку.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Свойство_ > полужирным шрифтом|Указывает, является ли шрифт полужирным.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Свойство_ > color|HTML-код цвета текста. Например, значение #FF0000 обозначает красный цвет.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Свойство_ > курсивом|Указывает, применяется ли курсив.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Свойство_ > зачеркивание|Указывает, зачеркнут ли шрифт.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Свойство_ > подчеркивание|Тип подчеркивания, применяемый для шрифта. Возможные значения: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Метод_ > clear()|Удаляет форматирование шрифтов.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Свойство_ > numberFormat|Представляет код в числовом формате Excel для данного диапазона. Удаляется, если передается значение null.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Отношения_ > границы|Коллекция объектов границы, которые применяются ко всему диапазону условного форматирования. Только для чтения.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Отношения_ > заливки|Возвращает объект заливки, определенный для всего диапазона условного форматирования. Только для чтения.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Связь_ > font|Возвращает объект шрифта, определенный для всего диапазона условного форматирования. Только для чтения.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Свойство_ > operator|Оператор условного форматирования текста. Возможные значения: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Свойство_ > text|Текстовое значение условного форматирования.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Свойство_ > ранга|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Свойство_ > тип|Значения форматирования на основе рейтинга. Возможные значения: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Отношения_ > формат|Возвращает объект формата, который содержит шрифт, заливку, границы и другие свойства условного форматирования. Только для чтения.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Отношения_ > правила|Представляет объект Rule в этом условном форматировании. Только для чтения.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Свойство_ > axisColor|HTML-код, представляющий цвет линии оси в формате #RRGGBB (например, FFA500) или в виде ключевого слова (например, orange).|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Свойство_ > axisFormat|Указывает, как определяется ось для гистограммы Excel. Возможные значения: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Свойство_ > barDirection|Представляет направление, которое должна использовать гистограмма. Возможные значения: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Свойство_ > showDataBarOnly|Значение true скрывает значения ячеек, где применяется гистограмма.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Отношения_ > lowerBoundRule|Правило для нижней границы гистограммы (и как ее вычислить).|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Отношения_ > negativeFormat|Представление всех значений слева от оси в гистограмме Excel. Только для чтения.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Отношения_ > positiveFormat|Представление всех значений справа от оси в гистограмме Excel. Только для чтения.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Отношения_ > upperBoundRule|Правило для верхней границы гистограммы (и как ее вычислить).|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Свойство_ > reverseIconOrder|Значение true меняет порядок значков в наборе значков на обратный. Обратите внимание, что это значение нельзя задать, если используются специальные значки.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Свойство_ > showIconOnly|Значение true скрывает значения и показывает только значки.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Свойство_ > style|Отображает параметр условного форматирования IconSet. Возможные значения: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Отношение_ > criteria|Массив условий и наборов значков для правил и специальных значков для условий. Обратите внимание, что для первого условия можно изменить только специальный значок. Тип, формула и оператор будут игнорироваться.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Отношения_ > формат|Возвращает объект формата, который содержит шрифт, заливку, границы и другие свойства условного форматирования. Только для чтения.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Отношения_ > правила|Правило условного форматирования.|1.6|
|[range](/javascript/api/excel/excel.range)|_Отношения_ > conditionalFormats|Коллекция объектов ConditionalFormats, которые пересекают диапазон. Только для чтения.|1.6|
|[range](/javascript/api/excel/excel.range)|_Метод_ > calculate()|Вычисляет диапазон ячеек на листе.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Отношения_ > формат|Возвращает объект формата, который содержит шрифт, заливку, границы и другие свойства условного форматирования. Только для чтения.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Отношения_ > правила|Правило условного форматирования.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Отношения_ > формат|Возвращает объект формата, который содержит шрифт, заливку, границы и другие свойства условного форматирования. Только для чтения.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Отношения_ > правила|Условия условного форматирования TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношения_ > internalTest|Только для внутреннего использования. Только для чтения.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > calculate(markAllDirty: bool)|Вычисляет все ячейки на листе.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Новые возможности Excel 1,5 API JavaScript

### <a name="custom-xml-part"></a>Пользовательская XML-часть

* Добавление коллекции пользовательских XML-частей к объекту книги.
* Получение пользовательской XML-части по идентификатору
* Получение новой ограниченной коллекции пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.
* Получение строки XML, связанной с частью.
* Предоставление идентификатора и пространства имен части.
* Добавление новой пользовательской XML-части к книге.
* Установка XML-части целиком.
* Удаление пользовательской XML-части.
* Удаление атрибута с указанным именем из элемента, указанного по XPath.
* Запрос содержимого XML по XPath.
* Вставка, обновление и удаление атрибутов.

**Пример реализации:** [здесь](https://github.com/mandren/Excel-CustomXMLPart-Demo) вы найдете пример реализации, в котором показано, как можно использовать XML-части в надстройке.

### <a name="others"></a>Другие
* Метод `range.getSurroundingRegion()` возвращает объект Range, представляющий область вокруг данного диапазона. Это диапазон, ограниченный любым сочетанием пустых строк и столбцов относительно данного диапазона.
* Методы `getNextColumn()` и `getPreviousColumn()`, `getLast() для столбца таблицы.
* Метод `getActiveWorksheet()` для книги.
* Метод `getRange(address: string)` для книги.
* Метод `getBoundingRange(ranges: )` возвращает наименьший объект диапазона, включающий в себя заданные диапазоны. Например, ограничивающий диапазон между диапазонами "B2:C5" и "D10:E15" — "B2:E15".
* С помощью метода `getCount()` можно получать количество элементов в различных коллекциях, таких как именованные элементы, листы, таблицы и т. д. `workbook.worksheets.getCount()`
* Методы `getFirst()` и `getLast()` для различных коллекций, таких как листы, столбцы таблицы, точки диаграммы и представления диапазонов.
* Методы `getNext()` и `getPrevious()` дли коллекций листов и столбцов таблиц.
* Метод `getRangeR1C1()` возвращает объект диапазона, начинающегося с определенных строки и столбца и занимающего определенное количество строк и столбцов.

|Объект| Что нового| Описание|Набор обязательных элементов|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Свойство_ > id|Идентификатор пользовательской XML-части. Только для чтения.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Свойство_ > namespaceUri|URI пространства имен пользовательской XML-части. Только для чтения.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Метод_ > delete()|Удаляет пользовательскую XML-часть.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Метод_ > getXml()|Получает полное содержимое пользовательской XML-части.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Метод_ > setXml(xml: string)|Задает полное содержимое пользовательской XML-части.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Свойство_ > items|Коллекция объектов customXmlPart. Только для чтения.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Метод_ > add(xml: string)|Добавление новой пользовательской XML-части к книге.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Метод_ > getByNamespace(namespaceUri: string)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Метод_ > getCount()|Возвращает количество частей CustomXml в коллекции.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Метод_ > getItem(id: string)|Возвращает пользовательскую XML-часть по идентификатору.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Метод_ > getItemOrNullObject(id: string)|Возвращает пользовательскую XML-часть по идентификатору.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Свойство_ > items|Коллекция объектов customXmlPartScoped. Только для чтения.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Метод_ > getCount()|Возвращает количество частей CustomXML в этой коллекции.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Метод_ > getItem(id: string)|Возвращает пользовательскую XML-часть по идентификатору.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Метод_ > getItemOrNullObject(id: string)|Возвращает пользовательскую XML-часть по идентификатору.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Метод_ > getOnlyItem()|Если коллекция содержит ровно один элемент, этот метод возвращает его.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Метод_ > getOnlyItemOrNullObject()|Если коллекция содержит ровно один элемент, этот метод возвращает его.|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношения_ > customXmlParts|Представляет коллекцию пользовательских XML-частей, содержащихся в этой книге. Только для чтения.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getNext(visibleOnly: bool)|Получает следующий лист. Если следующего листа нет, возникает ошибка.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getNextOrNullObject(visibleOnly: bool)|Получает следующий лист. Если следующего листа нет, метод возвращает объект null.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getPrevious(visibleOnly: bool)|Возвращает предыдущий лист. Если предыдущего листа нет, возникает ошибка.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getPreviousOrNullObject(visibleOnly: bool)|Возвращает предыдущий лист. Если предыдущего листа нет, этот метод возвращает объект null.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Метод_ > getFirst(visibleOnly: bool)|Возвращает первый лист в коллекции.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Метод_ > getLast(visibleOnly: bool)|Возвращает последний лист в коллекции.|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Новые возможности API JavaScript для Excel 1.4
Ниже приведены новые функции в API JavaScript Excel в требование набора 1.4.

### <a name="named-item-add-and-new-properties"></a>Именованный элемент add и новые свойства

Новые свойства:

* `comment`
* `scope` элементы, которые относятся к листу или книги
* `worksheet` возвращает лист, к которому относится именованный элемент.

Новые методы:

* `add(name: string, reference: Range or string, comment: string)`Добавляет новое имя в определенную коллекцию.
* `addFormulaLocal(name: string, formula: string, comment: string)` Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.

### <a name="settings-api-in-in-excel-namespace"></a>Параметры API в пространстве имен Excel

Объект [Setting](/javascript/api/excel/excel.setting) представляет пару "ключ-значение" для параметра, хранящегося в документе. Мы добавили API, связанные с параметрами, в пространство имен Excel. Они не обеспечивают новую функциональность, но позволяют оставаться в пакетном синтаксисе API на основе обещаний и уменьшить зависимость от общих задач, связанных с API для Excel.

API включают `getItem()` для получения параметра с помощью ключа, `add()` для добавления указанной пары параметров "ключ:значение" в книгу.

### <a name="others"></a>Другие

* Задайте имя столбца таблицы (в предыдущей версии разрешено только чтение).
* Добавьте столбец в конец таблицы (в предыдущей версии столбец можно добавить в любом месте, кроме последнего).
* Добавьте в таблицу сразу несколько строк (в предыдущей версии можно добавлять только 1 строку за раз).
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* Получение элемента или пустого объекта: Эта функция позволяет получить объект с помощью ключа. Если объект не существует, для свойства isNullObject возвращаемого объекта будет задано значение true. Это позволяет разработчикам проверить, существует ли объект, не обрабатывая его с помощью исключений. Доступно для листа, именованного элемента, привязки, ряда диаграммы и т. д.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Объект| Что нового| Описание|Набор требований|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > getCount()|Получает количество привязок в коллекции.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > getItemOrNullObject(id: string)|Получает объект привязки по идентификатору. Если объект привязки не существует, возвращает пустой объект.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Метод_ > getCount()|Возвращает количество диаграмм на листе.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Метод_ > getItemOrNullObject(name: string)|Возвращает диаграмму по ее имени. Если одно и то же имя принадлежит нескольким диаграммам, будет возвращена первая из них.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Метод_ > getCount()|Возвращает количество точек диаграммы в ряду.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Метод_ > getCount()|Возвращает количество рядов в коллекции.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Свойство_ > примечание|Представляет примечание, связанное с этим именем.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Свойство_ > область|Указывает, относится ли имя к книге или определенному листу. Только для чтения. Возможные значения: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Отношение_ > worksheet|Возвращает лист, к которому относится именованный элемент. Выдает ошибку, если элемент относится к книге. Только для чтения.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Отношение_ > worksheetOrNullObject|Возвращает лист, к которому относится именованный элемент. Возвращает пустой объект, если элемент относится к книге. Только для чтения.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Метод_ > delete()|Удаляет заданное имя.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Метод_ > getRangeOrNullObject()|Возвращает объект диапазона, связанный с именем. Возвращает пустой объект, если именованный элемент не является диапазоном.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Метод_ > Добавить (имя: ссылка на строку,: диапазон или строкой, комментарий: строка)|Добавляет новое имя в определенную коллекцию.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Метод_ > addFormulaLocal (имя: строка, формулу: строка, комментарий: строка)|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Метод_ > getCount()|Получает количество именованных элементов в коллекции.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Метод_ > getItemOrNullObject(name: string)|Получает объект nameditem по имени. Если объект nameditem не существует, возвращает пустой объект.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Метод_ > getCount()|Получает количество сводных таблиц в коллекции.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Метод_ > getItemOrNullObject(name: string)|Получает сводную таблицу по имени. Если сводная таблица не существует, возвращает пустой объект.|1.4|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getIntersectionOrNullObject (anotherRange: диапазон или строки)|Возвращает объект range, представляющий прямоугольное пересечение заданных диапазонов. Если пересечение не найдено, возвращает пустой объект.|1.4|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getUsedRangeOrNullObject(valuesOnly: bool)|Возвращает используемый диапазон заданного объекта диапазона. Если в диапазоне нет используемых ячеек, эта функция возвращает пустой объект.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Метод_ > getCount()|Получает количество объектов RangeView в коллекции.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Свойство_ > key|Возвращает ключ, представляющий идентификатор setting. Только для чтения.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Свойство_ > значение|Представляет значение, сохраненное для этого параметра.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Метод_ > delete()|Удаляет параметр.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Свойство_ > items|Коллекция объектов setting. Только для чтения.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > Добавить (ключ: строковое значение: (все))|Устанавливает или добавляет указанный параметр в книгу.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > getCount()|Возвращает количество параметров в коллекции.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > getItem(key: string)|Возвращает объект Setting по ключу.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > getItemOrNullObject(key: string)|Возвращает объект Setting по ключу. Если параметр не существует, возвращает пустой объект.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Отношение_ > параметры|Получает объект Setting, представляющий привязку, которая вызвала событие SettingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Метод_ > getCount()]|Получает количество таблиц в коллекции.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Метод_ > getItemOrNullObject (ключ: число или строка)|Получает таблицу по имени или ИД. Если таблица не существует, возвращает пустой объект.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Метод_ > getCount()|Получает количество столбцов в таблице.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Метод_ > getItemOrNullObject (ключ: число или строка)|Возвращает объект столбца по имени или ИД. Если столбец не существует, возвращает пустой объект.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Метод_ > getCount()|Получает количество строк в таблице.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношение_ > settings|Представляет коллекцию параметров, сопоставленных с книгой. Только для чтения.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Отношение_ > имена|Коллекция имен, относящих к текущему листу. Только для чтения.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Метод_ > getUsedRangeOrNullObject(valuesOnly: bool)|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки, которые содержат значение или форматирование. Если весь лист пустой, эта функция возвращает пустой объект.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Метод_ > getCount(visibleOnly: bool)|Получает количество листов в коллекции.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Метод_ > getItemOrNullObject(key: string)|Получает объект листа по его имени или ИД. Если лист не существует, возвращает пустой объект.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Новые возможности API JavaScript для Excel 1.3

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.3, относящийся к API JavaScript для Excel.

|Объект| Новые возможности| Описание|Набор обязательных элементов|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Метод_ > delete()|Удаляет привязку.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > Добавить (диапазона: диапазон или строка, bindingType: string, идентификатор: строка)|Добавляет привязку к определенному объекту Range.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > addFromNamedItem (имя: string, bindingType: string, id: строка)|Добавляет новую привязку с учетом именованного элемента в книге.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > addFromSelection (bindingType: string, id: строка)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Метод_ > getItemOrNull(id: string)|Возвращает объект binding по идентификатору. Если объект binding не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Метод_ > getItemOrNull(name: string)|Возвращает диаграмму по ее имени. Если одно и то же имя принадлежит нескольким диаграммам, будет возвращена первая из них.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Метод_ > getItemOrNull(name: string)|Возвращает объект nameditem по имени. Если объект nameditem не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Свойство_ > name|Имя сводной таблицы.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Отношение_ > worksheet|Лист, содержащий текущую сводную таблицу. Только для чтения.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Метод_ > refresh()|Обновляет сводную таблицу.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Свойство_ > items|Коллекция объектов pivotTable. Только для чтения.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Метод_ > getItem(name: string)|Возвращает сводную таблицу по имени.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Метод_ > getItemOrNull(name: string)|Возвращает сводную таблицу по имени. Если сводная таблица не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getIntersectionOrNull (anotherRange: диапазон или строки)|Возвращает объект range, представляющий прямоугольное пересечение заданных диапазонов. Если пересечение не найдено, возвращает пустой объект.|1.3|
|[range](/javascript/api/excel/excel.range)|_Метод_ > getVisibleView()|Представляет видимые строки текущего диапазона.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > cellAddresses|Представляет адреса ячеек RangeView. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > columnCount|Возвращает количество видимых столбцов. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > formulas|Представляет формулу в формате A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > formulasLocal|Представляет формулу в формате A1 на языке пользователя и в соответствии с его языковым стандартом.  Например, английская формула "=SUM(A1, introduced in 1.5)" превратится в "=СУММ(A1;1,5)" на русском языке.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > formulasR1C1|Представляет формулу в формате R1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > index|Возвращает значение, представляющее индекс RangeView. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > numberFormat|Представляет код в числовом формате Excel для данной ячейки.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > rowCount|Возвращает количество видимых строк. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > text|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > valueTypes|Представляет тип данных каждой ячейки. Только для чтения. Возможные значения: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Свойство_ > values|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейка, которая содержит ошибку, вернет строку ошибки.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Связь_ > rows|Представляет коллекцию объектов rangeView, сопоставленных с диапазоном. Только для чтения.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Метод_ > getRange()|Возвращает родительский диапазон, сопоставленный с текущим объектом RangeView.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Свойство_ > items|Коллекция объектов rangeView. Только для чтения.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Метод_ > getItemAt(index: number)|Возвращает строку RangeView по индексу. Используется нулевой индекс.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Свойство_ > key|Возвращает ключ, представляющий идентификатор setting. Только для чтения.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Метод_ > delete()|Удаляет параметр.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Свойство_ > items|Коллекция объектов setting. Только для чтения.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > getItem(key: string)|Возвращает объект Setting по ключу.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > getItemOrNull(key: string)|Возвращает объект Setting по ключу. Если экземпляр Setting не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Метод_ > задать (ключ: строковое значение: строка)|Устанавливает или добавляет указанный параметр в книгу.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Отношение_ > settingCollection|Получает объект Setting, представляющий привязку, которая вызвала событие SettingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > highlightFirstColumn|Указывает, содержит ли первый столбец специальное форматирование.|1.3|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > highlightLastColumn|Указывает, содержит ли последний столбец специальное форматирование.|1.3|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > showBandedColumns|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|1.3|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > showBandedRows|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|1.3|
|[table](/javascript/api/excel/excel.table)|_Свойство_ > showFilterButton|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Метод_ > getItemOrNull (ключ: число или строка)|Получает таблицу по имени или идентификатору. Если таблица не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Метод_ > getItemOrNull (ключ: число или строка)|Возвращает объект column по имени или идентификатору. Если столбец не существует, у свойства isNull возвращаемого объекта будет значение true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношение_ > pivotTables|Представляет коллекцию сводных таблиц, сопоставленных с книгой. Только для чтения.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношение_ > settings|Представляет коллекцию параметров, сопоставленных с книгой. Только для чтения.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Отношение_ > pivotTables|Коллекция сводных таблиц на листе. Только для чтения.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Новые возможности API JavaScript для Excel 1.2

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.2, относящийся к API JavaScript для Excel.

|Объект| Новые возможности| Описание|Набор обязательных элементов|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Свойство_ > id|Возвращает диаграмму с учетом ее положения в коллекции. Только для чтения.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Отношение_ > worksheet|Лист, содержащий текущую диаграмму. Только для чтения.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Метод_ > getImage (высота: номер, ширина: число, fittingMode: строка)|Отрисовывает диаграмму в виде изображения с кодировкой base64, масштабируя ее в соответствии с указанным размером.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Отношение_ > criteria|Текущий фильтр, заданный для определенного столбца. Только для чтения.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > apply(criteria: FilterCriteria)|Применяет заданные условия фильтра для определенного столбца.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyBottomItemsFilter(count: number)|Применяет к столбцу фильтр по количеству элементов снизу.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyBottomPercentFilter(percent: number)]|Применяет к столбцу фильтр по проценту элементов снизу.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyCellColorFilter(color: string)|Применяет к столбцу фильтр по цвету ячеек.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyCustomFilter (criteria1: string, criteria2: string, номер операции: строка)|Применяет к столбцу фильтр по условиям.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyDynamicFilter(criteria: string)|Применяет к столбцу динамический фильтр.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyFontColorFilter(color: string)|Применяет к столбцу фильтр по цвету шрифта.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyIconFilter(icon: Icon)|Применяет к столбцу фильтр по значку.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyTopItemsFilter(count: number)|Применяет к столбцу фильтр по количеству элементов сверху.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyTopPercentFilter(percent: number)|Применяет к столбцу фильтр по проценту элементов сверху.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > applyValuesFilter (значений: ())|Применяет к столбцу фильтр по значениям.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Метод_ > clear()|Сбрасывает фильтр для определенного столбца.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > color|Строка цвета HTML, которая используется для фильтрации ячеек. Используется с фильтрацией типа "cellColor" и "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > criterion1|Первый критерий фильтрации данных. Используется в качестве оператора при фильтрации типа "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > criterion2|Второй критерий фильтрации данных. Используется в качестве оператора только при фильтрации типа "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > dynamicCriteria|Динамические критерии из набора Excel.DynamicFilterCriteria, которые необходимо применить к этому столбцу. Используется с фильтрацией типа "dynamic". Возможные значения: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > filterOn|Свойство, с помощью которого фильтр определяет, следует ли показывать значения. Возможные значения: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > operator|Оператор, который используется для объединения условий 1 и 2 при "настраиваемой" фильтрации. Возможные значения: And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Свойство_ > values|Набор значений, который используется при фильтрации по значениям.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Отношение_ > icon|Значок, используемый для фильтрации ячеек. Используется с фильтрацией типа "icon".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Свойство_ > date|Дата в формате ISO8601, используемая для фильтрации данных.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Свойство_ > specificity|Точность, с которой производится фильтрация данных на основе даты. Например, если указана дата 2005-04-02, а для свойства specificity задано значение month, после фильтрации останутся все строки, датированные апрелем 2009 г. Возможные значения: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Свойство_ > formulaHidden|Указывает, скрывает ли Excel формулу для ячеек в диапазоне. Значение NULL указывает, что для всего диапазона не задан единый параметр скрытия формулы.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Свойство_ > locked|Указывает, блокирует ли Excel ячейки в объекте. Значение NULL указывает, что для всего диапазона не задан единый параметр блокировки.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Свойство_ > index|Представляет собой индекс значка данного набора.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Свойство_ > set|Представляет собой набор, в который входит значок. Возможные значения: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > columnHidden|Указывает, скрыты ли все столбцы текущего диапазона.|1.2|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > formulasR1C1|Представляет формулу в формате R1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > hidden|Указывает, скрыты ли все ячейки текущего диапазона. Только для чтения.|1.2|
|[range](/javascript/api/excel/excel.range)|_Свойство_ > rowHidden|Указывает, скрыты ли все строки текущего диапазона.|1.2|
|[range](/javascript/api/excel/excel.range)|_Отношение_ > sort|Представляет порядок сортировки текущего диапазона. Только для чтения.|1.2|
|[range](/javascript/api/excel/excel.range)|_Метод_ > merge(across: bool)|Объединяет ячейки диапазона в одну область на листе.|1.2|
|[range](/javascript/api/excel/excel.range)|_Метод_ > unmerge()|Разъединяет ячейки диапазона на отдельные ячейки.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Свойство_ > columnWidth|Возвращает или задает ширину всех столбцов в пределах диапазона. Если столбцы разной ширины, будет возвращено значение NULL.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Свойство_ > rowHeight|Возвращает или задает высоту всех строк в диапазоне. Если строки разной высоты, будет возвращено значение NULL.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Отношение_ > protection|Возвращает объект защиты формата для диапазона. Только для чтения.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Метод_ > autofitColumns()|Изменяет ширину столбцов текущего диапазона на оптимальную с учетом текущих данных в столбцах.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Метод_ > autofitRows()|Изменяет высоту строк текущего диапазона на оптимальную с учетом текущих данных в столбцах.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Свойство_ > address|Представляет видимые строки текущего диапазона.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Метод_ > Применить (полей: SortField matchCase: bool hasHeaders: bool ориентация: метод string: строка)|Выполняет сортировку.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Свойство_ > ascending|Указывает, выполняется ли сортировка по возрастанию.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Свойство_ > color|Представляет цвет, определенный условием, при сортировке по цвету шрифта или ячеек.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Свойство_ > dataOption|Представляет дополнительные параметры сортировки для этого поля. Возможные значения: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Свойство_ > key|Представляет столбец (или строку в зависимости от ориентации сортировки), для которого задано условие. Представляется в виде расстояния от первого столбца (или строки).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Свойство_ > sortOn|Представляет тип сортировки этого условия. Возможные значения: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Отношение_ > icon|Представляет значок, определенный условием, при сортировке по значку ячейки.|1.2|
|[table](/javascript/api/excel/excel.table)|_Отношение_ > sort|Представляет сортировку для таблицы. Только для чтения.|1.2|
|[table](/javascript/api/excel/excel.table)|_Отношение_ > worksheet|Лист, содержащий текущую таблицу. Только для чтения.|1.2|
|[table](/javascript/api/excel/excel.table)|_Метод_ > clearFilters()|Удаляет все фильтры, примененные к таблице.|1.2|
|[table](/javascript/api/excel/excel.table)|_Метод_ > convertToRange()|Преобразовывает таблицу в обычный диапазон ячеек. Все данные сохраняются.|1.2|
|[table](/javascript/api/excel/excel.table)|_Метод_ > reapplyFilters()|Повторно применяет все текущие фильтры к таблице.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Отношение_ > filter|Возвращает фильтр, применяемый к столбцу. Только для чтения.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Свойство_ > matchCase|Указывает, учитывался ли регистр при последней сортировке таблице. Только для чтения.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Свойство_ > method|Указывает метод сортировки китайских символов, который использовался при последней сортировке таблицы. Только для чтения. Возможные значения: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Отношение_ > fields|Указывает текущие условия, которые использовались при последней сортировке таблицы. Только для чтения.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Метод_ > Применить (полей: SortField matchCase: bool, метод: строка)|Выполняет сортировку.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Метод_ > clear()|Удаляет текущие параметры сортировки таблицы. При этом сбрасывается состояние кнопок в заголовках, но порядок сортировки таблицы остается неизменным.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Метод_ > reapply()|Повторно применяет текущие параметры сортировки к таблице.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Отношение_ > functions|Представляет экземпляр приложения Excel, содержащий эту книгу. Только для чтения.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Отношение_ > protection|Возвращает объект защиты листа. Только для чтения.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Свойство_ > protected|Указывает, защищен ли лист. Только для чтения. Только для чтения.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Отношение_ > options|Параметры защиты листа. Только для чтения.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Метод_ > protect(options: WorksheetProtectionOptions)|Защищает лист. Выдает ошибку, если лист защищен.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Метод_ > unprotect()|Снимает защиту с листа.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowAutoFilter|Представляет параметр защиты листа, разрешающий использовать функцию автофильтра.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowDeleteColumns|Представляет параметр защиты листа, разрешающий удалять столбцы.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowDeleteRows|Представляет параметр защиты листа, разрешающий удалять строки.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowFormatCells|Представляет параметр защиты листа, разрешающий форматировать ячейки.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowFormatColumns|Представляет параметр защиты листа, разрешающий форматировать столбцы.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowFormatRows|Представляет параметр защиты листа, разрешающий форматировать строки.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowInsertColumns|Представляет параметр защиты листа, разрешающий вставлять столбцы.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowInsertHyperlinks|Представляет параметр защиты листа, разрешающий вставлять гиперссылки.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowInsertRows|Представляет параметр защиты листа, разрешающий вставлять строки.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowPivotTables|Представляет параметр защиты листа, разрешающий использовать функцию сводных таблиц.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Свойство_ > allowSort|Представляет параметр защиты листа, разрешающий использовать функцию сортировки.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript для Excel 1.1

Excel JavaScript API 1.1 является первой версии API-интерфейса. Для получения дополнительных сведений об API видеть разделы справочника по [Excel JavaScript API](/javascript/api/excel) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
