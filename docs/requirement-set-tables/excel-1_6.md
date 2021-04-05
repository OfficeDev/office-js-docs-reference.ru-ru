| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Приостанавливать вычисление, пока не `context.sync()` будет вызван следующий.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Указывает объект правила в этом условном формате.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Критерии цветовой шкалы.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Если цветовая шкала будет иметь три точки (минимальная, средней точки, максимум), в противном случае она будет `true` иметь два (минимум, максимум).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Формула, если требуется, для оценки правила условного формата.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Формула, если требуется, для оценки правила условного формата.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Оператор условного формата значения ячейки.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Максимальная точка критерия цветовой шкалы.|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Середина критерия цветовой шкалы, если цветовая шкала — это трехцветная шкала.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Минимальная точка критерия цветовой шкалы.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|Представление цветового кода HTML цвета (например, #FF0000 представляет красный цвет).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Число, формула или `null` (если `type` `lowestValue` есть).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|На чем должна основываться условная формула критерия.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Указывает, имеет ли отрицательная планка данных тот же цвет границы, что и положительная планка данных.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Указывает, имеет ли отрицательная планка данных тот же цвет заполнения, что и положительный.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Указывает, есть ли в панели данных градиент.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Формула, если требуется, для оценки правила панели данных.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Тип правила для панели данных.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Удаляет это условное форматирование.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Возврат диапазона, к которому применено условное форматирование.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Возвращает диапазон, к которому применяется кондитональный формат.|
||[приоритет](/javascript/api/excel/excel.conditionalformat#priority)|Приоритет (или индекс) в условном наборе форматов, в который в настоящее время существует этот условный формат.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является `CellValue` типом.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Возвращает свойства условного формата значения ячейки, если текущий условный формат является `CellValue` типом.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Возвращает свойства условного формата цветовой шкалы, если текущий условный формат является `ColorScale` типом.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Возвращает свойства условного формата цветовой шкалы, если текущий условный формат является `ColorScale` типом.|
||[настраиваемый](/javascript/api/excel/excel.conditionalformat#custom)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Возвращает настраиваемые свойства условного формата, если текущий условный формат является пользовательским типом.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Возвращает свойства панели данных, если текущий условный формат является панели данных.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Возвращает свойства условного формата набора значков, если текущий условный формат является `IconSet` типом.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Возвращает свойства условного формата набора значков, если текущий условный формат является `IconSet` типом.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Приоритет условного формата в текущем `ConditionalFormatCollection` .|
||[предустановка](/javascript/api/excel/excel.conditionalformat#preset)|Возвращает условный формат предварительных критериев.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Возвращает условный формат предварительных критериев.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Возвращает определенные свойства условного формата текста, если текущий условный формат — это текстовый тип.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Возвращает свойства верхнего и нижнего условного формата, если текущий условный формат является `TopBottom` типом.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Возвращает свойства верхнего и нижнего условного формата, если текущий условный формат является `TopBottom` типом.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Тип условного формата.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Если выполняются условия этого условного форматирования, форматы с более низким приоритетом не будут применяться в этой ячейке.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Добавляет новый условный формат в коллекцию с первого и верхнего приоритета.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Полное удаление условного форматирование в указанном диапазоне.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Возвращает количество условных форматов в книге.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Возвращает условное форматирование для указанного идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Возвращает условное форматирование по индексу.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Формула, если требуется, для оценки правила условного формата.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Формула, если требуется, для оценки правила условного формата на языке пользователя.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Формула, если требуется, для оценки правила условного формата в нотации в стиле R1C1.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Пользовательский значок для текущего критерия, если он отличается от набора значков по умолчанию, будет `null` возвращен.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Число или формула в зависимости от типа.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` или `greaterThanOrEqual` для каждого из типов правил для условного формата значка.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|На чем должна основываться условная формула значка.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[критерий](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Критерий условного формата.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|ЦВЕТОВой код HTML, представляющий цвет пограничной строки, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Постоянное значение, указывающее определенную сторону границы.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Одна из констант стиля линии, определяющая стиль линии границы.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Возвращает объект границы по его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Возвращает объект границы по его индексу.|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Получает нижнюю границу.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Количество объектов границы в коллекции.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Получает левую границу.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Получает правую границу.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Получает верхнюю границу.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Удаляет заливку.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|ЦВЕТОВой код HTML, представляющий цвет заполнения, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Указывает, является ли шрифт смелым.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Удаляет форматирование шрифтов.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|Представление цветового кода HTML текстового цвета (например, #FF0000 представляет красный цвет).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Указывает, является ли шрифт italic.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Указывает состояние забастовки шрифта.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Тип подчеркнутого, примененного к шрифту.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Представляет код формата номеров Excel для данного диапазона.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Коллекция пограничных объектов, применимых к общему диапазону условного формата.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Возвращает объект заполнения, определенный в общем диапазоне условного формата.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Возвращает объект шрифта, определенный в общем диапазоне условного формата.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Оператор текстового условного формата.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Текстовое значение условного формата.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|От 1 до 1000 для числовых рейтингов или от 1 до 100 для процентных рейтингов.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Значения формата на основе верхнего или нижнего ранга.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.customconditionalformat#rule)|Указывает объект `Rule` в этом условном формате.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|ЦВЕТОВой код HTML, представляющий цвет линии Axis, в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Представление того, как определяется ось для панели данных Excel.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Указывает, в каком направлении должна основываться графика панели данных.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Правило для нижней границы гистограммы (и как ее вычислить).|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Представление всех значений слева от оси в панели данных Excel.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Представление всех значений справа от оси в панели данных Excel.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Если `true` , скрывает значения из ячеек, где применяется планка данных.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Правило для верхней границы гистограммы (и как ее вычислить).|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Набор критериев и наборов значков для правил и потенциальных пользовательских значков для условных значков.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Если `true` , отменит заказы значка для набора значков.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Если `true` , скрывает значения и показывает только значки.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Если установлено, отображается параметр набора значков для условного формата.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условных форматов, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Правило условного форматирования.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Вычисляет диапазон ячеек на листе.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Эта коллекция `ConditionalFormats` пересекает диапазон.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условного формата, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.textconditionalformat#rule)|Правило условного форматирования.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Возвращает объект формата, инкапсулируя шрифт условного формата, заполнять, границы и другие свойства.|
||[правило](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Критерии условного формата верхнего и нижнего.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Вычисляет все ячейки на листе.|
