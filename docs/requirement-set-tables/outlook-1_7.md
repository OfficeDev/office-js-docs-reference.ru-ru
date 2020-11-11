| Класс | Поля | Описание |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (eventType: Office. EventType \| строка, обработчик: Any, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Добавляет обработчик для поддерживаемого события.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Получает организатора для указанного собрания.|
||[повторения](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Получает или задает шаблон повторения встречи.|
||[removeHandlerAsync (eventType: Office. EventType \| строка, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Удаляет обработчиков для поддерживаемого типа события.|
||[seriesId](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|Получает идентификатор ряда, к которому принадлежит экземпляр.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (eventType: Office. EventType \| строка, обработчик: Any, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Добавляет обработчик для поддерживаемого события.|
||[повторения](/javascript/api/outlook/outlook.appointmentread#recurrence)|Получает шаблон повторения встречи.|
||[removeHandlerAsync (eventType: Office. EventType \| строка, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|Удаляет обработчиков для поддерживаемого типа события.|
||[seriesId](/javascript/api/outlook/outlook.appointmentread#seriesid)|Получает идентификатор ряда, к которому принадлежит экземпляр.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[Async (Options?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|Получает значение, заданное в списке.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (eventType: Office. EventType \| строка, обработчик: Any, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Добавляет обработчик для поддерживаемого события.|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|Получает электронный адрес отправителя сообщения.|
||[removeHandlerAsync (eventType: Office. EventType \| строка, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Удаляет обработчиков для поддерживаемого типа события.|
||[seriesId](/javascript/api/outlook/outlook.messagecompose#seriesid)|Получает идентификатор ряда, к которому принадлежит экземпляр.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (eventType: Office. EventType \| строка, обработчик: Any, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|Добавляет обработчик для поддерживаемого события.|
||[повторения](/javascript/api/outlook/outlook.messageread#recurrence)|Получает шаблон повторения встречи.|
||[removeHandlerAsync (eventType: Office. EventType \| строка, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|Удаляет обработчиков для поддерживаемого типа события.|
||[seriesId](/javascript/api/outlook/outlook.messageread#seriesid)|Получает идентификатор ряда, к которому принадлежит экземпляр.|
|[Organizer](/javascript/api/outlook/outlook.organizer)|[Async (Options?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|Получает значение организатора встречи в виде {@link Office. EmailAddressDetails | Объект EmailAddressDetails}|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[чанжедреЦипиентфиелдс](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Получает значение, указывающее, были ли изменены получатели в поле **"СК"** .|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Получает значение, указывающее, были ли изменены получатели в поле **"копия"** .|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|Получает, если были изменены необязательные участники.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|Получает, если обязательные участники были изменены.|
||[ресурсы](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Возвращает, если ресурсы были изменены.|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|Получает значение, указывающее, были ли изменены получатели в поле " **Кому** ".|
|[Recurrence](/javascript/api/outlook/outlook.recurrence)|[Async (Options?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|Возвращает текущий объект повторения ряда встреч.|
||[рекурренцепропертиес](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|Получает или задает свойства ряда повторяющихся встреч.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|Получает или задает свойства ряда повторяющихся встреч.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|Получает или задает тип ряда повторяющихся встреч.|
||[сериестиме](/javascript/api/outlook/outlook.recurrence#seriestime)|{@Link Office. Сериестиме | Сериестиме} объект позволяет управлять начальной и конечной датами ряда повторяющихся встреч и|
||[setAsync (recurrencePattern: повторения, параметры?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|Задает шаблон повторения для ряда встреч.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[повторения](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|Представляет день месяца.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|Представляет день недели или тип дня, например, выходной день и день недели.|
||[срок](/javascript/api/outlook/outlook.recurrenceproperties#days)|Представляет набор дней для этого повторения.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|Представляет выбранный первый день недели, в противном случае значением по умолчанию является значение в параметрах текущего пользователя.|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Представляет период между экземплярами одних и тех же повторяющихся рядов.|
||[month](/javascript/api/outlook/outlook.recurrenceproperties#month)|Представляет месяц.|
||[викнумбер](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|Представляет число недель в выбранном месяце, например, ' first ' для первой недели месяца.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Представляет имя часового пояса.|
||[корреспондирующей](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Целое значение, представляющее разницу в минутах между местным часовым поясом и временем в формате UTC на дату начала серии собраний.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[ДЛИТ ()](/javascript/api/outlook/outlook.seriestime#getduration--)|Получает значение времени в минутах для обычного экземпляра в серии повторяющихся встреч.|
||[Жетенддате ()](/javascript/api/outlook/outlook.seriestime#getenddate--)|Получает дату окончания расписания повторения в следующем|
||[Жетендтиме ()](/javascript/api/outlook/outlook.seriestime#getendtime--)|Получает время окончания обычной встречи или экземпляра приглашения на собрание для шаблона повторения в любом часовом поясе, который пользователь или|
||[StartDate ()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|Получает дату начала расписания повторения в следующем|
||[Жетстарттиме ()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|Получает время начала обычного экземпляра встречи шаблона повторения в каком-либо часовом поясе, в котором пользователь или надстройка установили|
||[Сетдуратион (минуты: число)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|Задает продолжительность всех встреч в расписании повторения.|
||[Сетенддате (Date: строка)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|Задает дату окончания ряда повторяющихся встреч.|
||[Сетенддате (год: число, месяц: число, день: число)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|Задает дату окончания ряда повторяющихся встреч.|
||[Сетстартдате (Date: строка)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|Задает дату начала ряда повторяющихся встреч.|
||[Сетстартдате (год: число, месяц: число, день: число)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|Задает дату начала ряда повторяющихся встреч.|
||[Сетстарттиме (часы: число, минуты: число)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|Задает время начала всех экземпляров ряда повторяющихся встреч в каком-либо часовом поясе, заданном шаблоном повторения|
||[Сетстарттиме (Time: строка)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|Задает время начала всех экземпляров ряда повторяющихся встреч в каком-либо часовом поясе, заданном шаблоном повторения|
