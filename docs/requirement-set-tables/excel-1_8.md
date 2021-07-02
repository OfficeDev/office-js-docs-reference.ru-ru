| Класс | Поля | Описание |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|Оператор, используемый для проверки данных.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Указывает константу индексации уровня метки категорий диаграммы, ссылаясь на уровень меток исходных категорий.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Указывает, как пустые ячейки заданы на диаграмме.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|True, если отображаются только видимые ячейки.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Возникает при активации диаграммы.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Происходит, когда диаграмма отключена.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Представляет область сюжета для диаграммы.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Указывает константу индексации имен на уровне серии диаграмм, ссылаясь на уровень имен исходных серий.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Указывает, следует ли показывать метки данных, если значение превышает максимальное значение оси значения.|
||[style](/javascript/api/excel/excel.chart#style)|Указывает стиль диаграммы для диаграммы.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Получает ID активированной диаграммы.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Получает ID таблицы, в которой активируется диаграмма.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Получает ID диаграммы, добавляемой в таблицу.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Получает ID таблицы, в которую добавляется диаграмма.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[выравнивание](/javascript/api/excel/excel.chartaxis#alignment)|Указывает выравнивание для указанной метки тик оси.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Указывает, пересекает ли ось значения ось категории между категориями.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multilevel)|Указывает, многоуровневая ли ось.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Указывает код формата для метки тик оси.|
||[смещение](/javascript/api/excel/excel.chartaxis#offset)|Указывает расстояние между уровнями меток и расстоянием между первым уровнем и линией оси.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Указывает указанное положение оси, где пересекается другая ось.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Указывает положение оси, где пересекается другая ось.|
||[setPositionAt (значение: номер)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Задает указанное положение оси, где пересекается другая ось.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Указывает угол, на который ориентирован текст для метки тика оси диаграммы.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Указывает форматирование заполнения диаграммы.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Строковое значение, представляющее формулу заголовка оси диаграммы с использованием нотации стиля A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[граница](/javascript/api/excel/excel.chartaxistitleformat#border)|Указывает пограничный формат заголовка оси диаграммы, который включает цвет, листил и вес.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Указывает форматирование заполнения заголовок оси диаграммы.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Очищает формат границы элемента диаграммы.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Возникает при активации диаграммы.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Возникает при добавлении новой диаграммы в таблицу.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Происходит, когда диаграмма отключена.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Возникает при удалении диаграммы.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autotext)|Указывает, автоматически ли метка данных создает соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Строковое значение, представляющее формулу метки данных диаграммы с использованием нотации стиля A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Представляет горизонтальное выравнивание для метки данных диаграммы.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Представляет расстояние от левого края метки данных диаграммы до левого края области диаграммы (в пунктах). |
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Строковое значение, представляющее код формата для метки данных.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Представляет формат метки данных диаграммы.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Возвращает высоту метки данных диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Возвращает ширину метки данных диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Строка, представляющая текст метки данных на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Представляет угол, на который ориентирован текст для метки данных диаграммы.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Представляет расстояние от верхнего края метки данных диаграммы до верха области диаграммы (в пунктах).|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[граница](/javascript/api/excel/excel.chartdatalabelformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autotext)|Указывает, автоматически ли метки данных создают соответствующий текст на основе контекста.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Указывает горизонтальное выравнивание для метки данных диаграммы.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Указывает код формата для меток данных.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Представляет угол, на который ориентирован текст для меток данных.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Представляет вертикальное выравнивание для метки данных диаграммы.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Получает ID отключаемой диаграммы.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Получает ID таблицы, в которой деактивируется диаграмма.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Получает ID диаграммы, удаляемой из таблицы.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Получает ID таблицы, в которой удаляется диаграмма.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Указывает высоту записи легенды в легенде диаграммы.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Указывает индекс записи легенды в легенде диаграммы.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Указывает левое значение записи легенды диаграммы.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Указывает верхнюю часть записи легенды диаграммы.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Представляет ширину записи легенды на диаграмме Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[граница](/javascript/api/excel/excel.chartlegendformat#border)|Представляет формат границы, включающий цвет, тип линии и толщину.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Указывает значение высоты области участка.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Указывает внутреннее значение высоты области участка.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Указывает внутреннее левое значение области сюжета.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Указывает внутреннее верхнее значение области сюжета.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Указывает внутреннее значение ширины области участка.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Указывает левое значение области сюжета.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Указывает положение области сюжета.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Указывает форматирование области сюжета диаграммы.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Указывает верхнее значение области сюжета.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Указывает значение ширины области участка.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[граница](/javascript/api/excel/excel.chartplotareaformat#border)|Указывает атрибуты границы области диаграммы.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Указывает формат заполнения объекта, который включает сведения о формате фона.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Указывает группу для указанной серии.|
||[взрыв](/javascript/api/excel/excel.chartseries#explosion)|Указывает значение взрыва для среза круговой диаграммы или пончик-диаграммы.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Указывает угол первого среза круговой диаграммы или пончик-диаграммы в градусах (по часовой стрелке от вертикальной).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Верно, Excel выверяет шаблон в элементе, если он соответствует отрицательному номеру.|
||[перекрытие](/javascript/api/excel/excel.chartseries#overlap)|Указывает на расположение строк и столбцов.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Представляет коллекцию всех меток данных в серии.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Указывает размер вторичного раздела диаграммы пирога или диаграммы с круговым пирогом в процентах от размера первичного пирога.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Указывает способ разделения двух разделов диаграммы "пирог-пирог" или диаграммы "планка пирога".|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True, Excel назначит каждому маркеру данных другой цвет или шаблон.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Представляет число периодов, на которые линия тренда расширяется назад.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Представляет число периодов, на которые линия тренда расширяется вперед.|
||[метка](/javascript/api/excel/excel.charttrendline#label)|Представляет метку линии тренда диаграммы.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|Значение true, если формула для линии тренда отображается на диаграмме.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Значение True, если значение r-squared для линии тренда отображается на диаграмме.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Указывает, автоматически ли метка trendline создает соответствующий текст на основе контекста.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Строковая величина, которая представляет формулу метки трендовой линии диаграммы с помощью нотации в стиле A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Представляет горизонтальное выравнивание метки трендовой линии диаграммы.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Представляет расстояние в точках от левого края метки трендовой линии диаграммы до левого края области диаграммы.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Строковое значение, которое представляет код формата для метки trendline.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Формат метки трендовой линии диаграммы.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Возвращает высоту подписи линии тренда диаграммы (в пунктах).|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Возвращает ширину подписи линии тренда диаграммы (в пунктах).|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Строка, представляющая текст подписи линии тренда на диаграмме.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Представляет угол, на который ориентирован текст для метки трендовой линии диаграммы.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Представляет расстояние в точках от верхнего края метки трендовой линии диаграммы до верхней части области диаграммы.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Представляет вертикальное выравнивание метки трендовой линии диаграммы.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[граница](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Указывает пограничный формат, который включает цвет, литейный стил и вес.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Указывает формат заполнения текущей метки трендовой линии диаграммы.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Указывает атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для метки трендовой линии диаграммы.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Формула проверки настраиваемых данных.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Имя DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Числовой формат DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Положение DataPivotHierarchy.|
||[поле](/javascript/api/excel/excel.datapivothierarchy#field)|Возвращает сводные поля, связанные с DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Сбрасывает DataPivotHierarchy до значений по умолчанию.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Указывает, следует ли показывать данные в качестве определенного суммарного вычисления.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Указывает, показаны ли все элементы DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Получает DataPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Получает DataPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Очищает проверку данных из текущего диапазона.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Сообщение об ошибке, когда пользователь вводит недопустимые данные.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Указывает, будет ли проверка данных выполняться на пустых ячейках.|
||[сообщение](/javascript/api/excel/excel.datavalidation#prompt)|Подсказка, когда пользователи выбирают ячейку.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Тип проверки данных см. `Excel.DataValidationType` в подробностях.|
||[допустимо](/javascript/api/excel/excel.datavalidation#valid)|Указывает, являются ли все значения ячеек допустимыми в соответствии с правилами проверки данных.|
||[правило](/javascript/api/excel/excel.datavalidation#rule)|Правило проверки данных, которое содержит различные типы критериев проверки данных.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Представляет сообщение оповещений об ошибке.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Указывает, следует ли показывать диалоговое окно оповещения об ошибке при вводе пользователем недействительных данных.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Тип оповещений о проверке данных см. `Excel.DataValidationAlertStyle` в подробной информации.|
||[заголовок](/javascript/api/excel/excel.datavalidationerroralert#title)|Представляет название диалоговое окно оповещений об ошибке.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Указывает сообщение запроса.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Указывает, отображается ли подсказка, когда пользователь выбирает ячейку с проверкой данных.|
||[заголовок](/javascript/api/excel/excel.datavalidationprompt#title)|Указывает заголовок для запроса.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[настраиваемый](/javascript/api/excel/excel.datavalidationrule#custom)|Условия проверки настраиваемых данных.|
||[дата](/javascript/api/excel/excel.datavalidationrule#date)|Условия проверки данных даты.|
||[десятичной](/javascript/api/excel/excel.datavalidationrule#decimal)|Условия проверки десятичных данных.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Условия проверки данных списка.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Критерии проверки данных длины текста.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Условия проверки данных времени.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Все критерии проверки данных номеров.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Указывает операнд правой руки, когда свойство оператора задано двоичному оператору, такому как GreaterThan (левая операнд — это значение, в который пользователь пытается ввести в ячейку).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|С помощью ternary operators Between and NotBetween указывается верхний операнд.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|Оператор, используемый для проверки данных.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Определяет, следует ли разрешить несколько элементов фильтра.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Имя FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Положение FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Возвращает сводные поля, связанные с FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Сбрасывает FilterPivotHierarchy до значений по умолчанию.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Получает filterPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Получает FilterPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Указывает, следует ли отображать список в выпадаемой ячейке.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Источник списка для проверки данных|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Имя сводного поля.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Возвращает сводные поля, связанные со сводным полем.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Определяет, следует ли отображать все элементы сводного поля.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Сортирует сводное поле.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Промежуточные итоги сводного поля.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Получает количество поворотных полей в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Получает PivotField по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Получает PivotField по имени.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Имя PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Возвращает сводные поля, связанные с PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Получает PivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Получает PivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Определяет, развернут ли элемент для отображения дочерних элементов или же свернут, а дочерние элементы являются скрытыми.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Имя элемента сводной таблицы.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Указывает, отображается ли pivotItem.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Получает число pivotItems в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Получает PivotItem по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Получает PivotItem по имени.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Возвращает диапазон, где находятся названия столбцов сводной таблицы.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Возвращает диапазон, где находятся значения данных сводной таблицы.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Возвращает диапазон области фильтра сводной таблицы.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Возвращает диапазон, в котором существует сводная таблица, за исключением области фильтра.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Возвращает диапазон, где находятся названия строк сводной таблицы.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Это свойство указывает PivotLayoutType всех полей в сводной таблице.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Указывает, показывает ли отчет PivotTable общие итоги для столбцов.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Указывает, показывает ли отчет PivotTable общие итоги для строк.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Это свойство указывает все `SubtotalLocationType` поля на PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Удаляет сводную таблицу.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|Иерархии сводных столбцов сводной таблицы.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|Иерархии сводных данных сводной таблицы.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|Иерархии сводных фильтров сводной таблицы.|
||[иерархии](/javascript/api/excel/excel.pivottable#hierarchies)|Иерархии сводного документа сводной таблицы.|
||[макет](/javascript/api/excel/excel.pivottable#layout)|PivotLayout, описывающий макет и визуальную структуру сводной таблицы.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|Иерархии сводных строк сводной таблицы.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Добавьте pivotTable на основе указанных исходных данных и вставьте его в верхней левой ячейке диапазона назначения.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Возвращает объект проверки данных.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Имя RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Положение RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Возвращает сводные поля, связанные с RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Сбрасывает RowColumnPivotHierarchy до значений по умолчанию.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Добавляет PivotHierarchy к текущей оси.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Получает количество иерархий сводного объекта в коллекции.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Получает RowColumnPivotHierarchy по имени или ID.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Получает RowColumnPivotHierarchy по имени.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Удаляет PivotHierarchy из текущей оси.|
|[Время выполнения](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Добавление событий JavaScript в текущую области задач или надстройку контента.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|PivotField на основе расчета, если применимо `ShowAs` в соответствии с `ShowAsCalculation` типом, еще `null` .|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|Элемент, на основе `ShowAs` расчета, если применимо в соответствии с `ShowAsCalculation` типом, еще `null` .|
||[вычисление](/javascript/api/excel/excel.showasrule#calculation)|`ShowAs`Вычисление, используемого для PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Указывает, будет ли текст автоматически отступным, если выравнивание текста в ячейке задано на равное распределение.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Ориентация текста для стиля.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Если установлено значение , все остальные значения будут `Automatic` `true` игнорироваться при настройке `Subtotals` .|
||[среднее значение](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[продукт](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[отклонение](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Возвращает числимый ID.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, который представляет измененную область таблицы на определенном таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|`true`Возвращается, если книга открыта в режиме только для чтения.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Возникает при расчете таблицы.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Указывает, видны ли линии сетки пользователю.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Указывает, видны ли заголовки пользователю.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Получает ID таблицы, в которой произошел расчет.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Получает диапазон, представляющий измененную область конкретного листа.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Возникает при расчете любого таблицы в книге.|
