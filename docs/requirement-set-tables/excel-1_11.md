| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Предоставляет сведения, основанные на текущих параметрах языковых параметров системы.|
||[деЦималсепаратор](/javascript/api/excel/excel.application#decimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений.|
||[саусандссепаратор](/javascript/api/excel/excel.application#thousandsseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений.|
||[усесистемсепараторс](/javascript/api/excel/excel.application#usesystemseparators)|Указывает, включены ли системные разделители Excel.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[ричконтент](/javascript/api/excel/excel.comment#richcontent)|Получает содержимое форматированного комментария (например, упоминание в комментариях).|
||[определяем](/javascript/api/excel/excel.comment#resolved)|Состояние цепочки комментариев.|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (Целладдресс: \| строка Range, Content: комментричконтент \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Адрес электронной почты объекта, который упоминается в примечании.|
||[id](/javascript/api/excel/excel.commentmention#id)|Идентификатор объекта.|
||[name](/javascript/api/excel/excel.commentmention#name)|Имя объекта, который упоминается в примечании.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Сущности (например, люди), которые упоминаются в комментариях.|
||[определяем](/javascript/api/excel/excel.commentreply#resolved)|Состояние ответа на комментарий.|
||[ричконтент](/javascript/api/excel/excel.commentreply#richcontent)|Содержимое форматированного комментария (например, упоминание в комментариях).|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Добавить (контент: \| строка комментричконтент, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущности (например, люди), упомянутые в комментарии.|
||[ричконтент](/javascript/api/excel/excel.commentrichcontent#richcontent)|Задает расширенное содержимое комментария (например, закомментировать содержимое с упоминанием о том, что первый упомянутый объект имеет атрибут ID 0, а вторая упомянутая сущность имеет атрибут ID 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя языка и региональных параметров в формате languagecode2-Country/regioncode2 (например, "zh-CN" или "en-US").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Определяет формат отображения чисел, соответствующий культуре.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[нумбердеЦималсепаратор](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений.|
||[нумберграупсепаратор](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo (Дестинатионранже: \| строка Range)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Перемещает значения ячеек, форматирование и формулы из текущего диапазона в конечный диапазон, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Аджустиндент (Amount: число)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Настраивает отступ для форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[онровхидденчанжед](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, который выполнил вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[онровхидденчанжед](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменения, которое представляет способ запуска события.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
