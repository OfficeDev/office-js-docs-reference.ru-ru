# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> Эта документация является **предварительной версии** [задать требования](/javascript/office/requirement-sets/outlook-api-requirement-sets). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки. Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.

Наборы требований Preview включает в себя все возможности [требование задать 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

- [От](/javascript/api/outlook/office.from) - добавлен новый объект, который предоставляет метод для получения из значения.
- [Организатор](/javascript/api/outlook/office.organizer) - добавлен новый объект, который предоставляет метод для получения значения Организатор.
- [Повторение](/javascript/api/outlook/office.recurrence) - добавлен новый объект, предоставляющий методы для получения и задать шаблон повторения встреч, но только получить шаблон повторения сообщений, которые являются приглашений на собрания.
- [SeriesTime](/javascript/api/outlook/office.seriestime) - добавлен новый объект, который предоставляет методы для получения и задания даты и время встречи в ряду и для получения значения даты и времени приглашений в ряду.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-). Новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением (`allowEvent`). Это значение используется для отмены выполнения события.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - добавлен новый метод, который подключает файла из base64 кодирования в сообщении или встрече.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - добавлен новый метод, который добавляет обработчик событий для события, поддерживаемые.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - изменены для получения из значения в режиме создания.
- [Office.context.mailbox.item.getInitializationContextAsync.](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) - изменены для получения значения картинок в режиме создания.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - добавлено новое свойство, которое возвращает или задает объект, который предоставляет методы для управления повторов элемента встречи. Это свойство можно также использовать для получения шаблон повторения собрания элемента запроса.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - добавлен новый метод, который удаляет обработчик событий.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - добавлено новое свойство, которое получает идентификатор серии вхождения принадлежит.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Добавлен доступ к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - добавлено новое перечисление, указывающее день недели или тип дня.
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - добавлено новое перечисление, указывающее месяца.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - добавлено новое перечисление, указывающее, часовой пояс, применяемые к повторения.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - добавлено новое перечисление, определяющее тип повторения.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - добавлено новое перечисление, указывающее, в течение недели после месяца.
- [Office.EventType](/javascript/api/office/office.eventtype) - изменены для поддержки RecurrenceChanged, RecipientsChanged, AppointmentTimeChanged и OfficeThemeChanged событий с помощью добавления `RecurrencePatternChanged`, `RecipientsChanged`, `AppointmentTimeChanged`, и `OfficeThemeChanged` записи соответственно.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)