| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureInfo)|Предоставляет сведения, основанные на текущих параметрах культуры системы.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalSeparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsSeparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
||[useSystemSeparators](/javascript/api/excel/excel.application#useSystemSeparators)|Указывает, включены ли системные Excel системы.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, люди), указанные в комментариях.|
||[разрешено](/javascript/api/excel/excel.comment#resolved)|Состояние потока комментариев.|
||[richContent](/javascript/api/excel/excel.comment#richContent)|Получает богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updateMentions_contentWithMentions_)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Создает новое примечание с указанным содержимым в определенной ячейке.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Адрес электронной почты объекта, упоминаемого в комментарии.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID объекта.|
||[name](/javascript/api/excel/excel.commentmention#name)|Имя объекта, упомянутого в комментарии.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Сущностям (например, людям), упомянутым в комментариях.|
||[разрешено](/javascript/api/excel/excel.commentreply#resolved)|Состояние ответа на комментарий.|
||[richContent](/javascript/api/excel/excel.commentreply#richContent)|Богатое содержимое комментариев (например, упоминания в комментариях).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updateMentions_contentWithMentions_)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Создает ответ на комментарий для комментария.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущностями (например, людьми), упомянутыми в комментарии.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richContent)|Указывает богатое содержимое комментария (например, комментарий контента с упоминаниями, первая упомянутая сущность имеет атрибут ID 0, а вторая упомянутая сущность имеет атрибут ID 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя культуры в формате languagecode2-country/regioncode2 (например, "zh-cn" или "ru-ru").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberFormat)|Определяет культурный формат отображения номеров.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberDecimalSeparator)|Получает строку, используемую в качестве десятичных сепараторов для числевых значений.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numberGroupSeparator)|Получает строку, используемую для отдельных групп цифр слева от десятичной для числимых значений.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveTo_destinationRange_)|Перемещает значения ячейки, форматирование и формулы из текущего диапазона в диапазон назначения, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustIndent_amount_)|Регулирует отступ форматирования диапазона.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close_closeBehavior_)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save_saveBehavior_)|Сохраняет текущую книгу.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onRowHiddenChanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, завершив вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged)|Происходит, когда скрытое состояние одной или более строк изменилось на определенной таблице.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changeType)|Получает тип изменений, которые представляют, как было вызвано событие.|
||[источник](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменились данные.|
