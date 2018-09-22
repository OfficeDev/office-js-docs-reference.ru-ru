# <a name="outlook-add-in-api-requirement-set-17"></a>Задайте 1.7 требование API надстройки для Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-17"></a>Новые возможности в 1,7?

Наборы требований 1.7 включает в себя все возможности [требование задать 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Он добавлены следующие возможности.

- Добавлены новые интерфейсы API относительно шаблона повторения встречи и сообщений, которые являются приглашений на собрания.
- Изменение свойств item.from также будут доступны в режиме создания.
- Добавлена поддержка для события RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged.

### <a name="change-log"></a>Журнал изменений

- Добавлена [из](/javascript/api/outlook_1_7/office.from): Добавляет новый объект, который предоставляет метод для получения из значения.
- Добавлена [Организатор](/javascript/api/outlook_1_7/office.organizer): Добавляет новый объект, который предоставляет метод для получения значения Организатор.
- Добавлена [повторения](/javascript/api/outlook_1_7/office.recurrence): Добавляет новый объект, который предоставляет методы для получения и задать шаблон повторения встречи, но только получить шаблон повторения сообщений, которые приглашений на собрания.
- Добавлена [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): Добавляет новый объект, который представляет конфигурацию часового пояса шаблона повторения.
- Добавлена [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): Добавляет новый объект, который предоставляет методы для получения и задания даты и время встречи в ряду и для получения значения даты и времени приглашений в ряду.
- Добавлена [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): Добавляет новый метод, который добавляет обработчик событий для события, поддерживаемые.
- Изменены [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): изменяется для получения из значения в режиме создания.
- Измененные [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - изменяет для получения значения картинок в режиме создания.
- Добавлена [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): Добавляет новое свойство, которое возвращает или задает объект, который предоставляет методы для управления повторов элемента встречи. Это свойство можно также использовать для получения шаблон повторения собрания элемента запроса.
- Добавлена [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): Добавляет новый метод, который удаляет обработчик событий.
- Добавлена [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): Добавляет новое свойство, которое получает идентификатор этой серии вхождения принадлежит.
- Добавлена [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): Добавляет новое перечисление, указывающее день недели или тип дня.
- Добавлена [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): Добавляет новое перечисление, указывающее месяца.
- Добавлена [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): Добавляет новое перечисление, указывающее, часовой пояс, применяемые к повторения.
- Добавлена [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): Добавляет новое перечисление, определяющее тип повторения.
- Добавлена [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): Добавляет новое перечисление, указывающее, в течение недели после месяца.
- Изменены [Office.EventType](/javascript/api/office/office.eventtype): изменяется для поддержки RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged событий с помощью добавления `RecurrenceChanged`, `RecipientsChanged`, и `AppointmentTimeChanged` записи соответственно.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)