| Класс | Поля | Описание |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[Close ()](/javascript/api/outlook/outlook.appointmentcompose#close--)|Закрывает текущий элемент, который составляется|
||[notificationMessages](/javascript/api/outlook/outlook.appointmentcompose#notificationmessages)|Получает сообщения уведомления для элемента.|
||[saveAsync (callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-callback--asyncresult-)|Асинхронно сохраняет элемент.|
||[saveAsync (Options: Office. Асинкконтекстоптионс, callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-options--callback--asyncresult-)|Асинхронно сохраняет элемент.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[notificationMessages](/javascript/api/outlook/outlook.appointmentread#notificationmessages)|Получает сообщения уведомления для элемента.|
|[Основной текст](/javascript/api/outlook/outlook.body)|[-Async (coercionType: строка Office. CoercionType \| , Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.body#getasync-coerciontype--options--callback--asyncresult-)|Возвращает текущий текст в указанном формате.|
||[setAsync (Data: String, Options?: Office. Асинкконтекстоптионс & КоерЦионтипеоптионс, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setasync-data--options--callback--asyncresult-)|Заменяет весь текст указанным текстом.|
|[Mailbox](/javascript/api/outlook/outlook.mailbox)|[Конверттоевсид (itemId: строка, Рестверсион: строка MailboxEnums. Рестверсион \| )](/javascript/api/outlook/outlook.mailbox#converttoewsid-itemid--restversion-)|Преобразовывает идентификатор элемента из формата REST в формат EWS.|
||[convertToRestId (itemId: строка, Рестверсион: строка MailboxEnums. Рестверсион \| )](/javascript/api/outlook/outlook.mailbox#converttorestid-itemid--restversion-)|Преобразовывает идентификатор элемента в формате EWS в формат REST.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[Close ()](/javascript/api/outlook/outlook.messagecompose#close--)|Закрывает текущий элемент, который составляется|
||[notificationMessages](/javascript/api/outlook/outlook.messagecompose#notificationmessages)|Получает сообщения уведомления для элемента.|
||[saveAsync (callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#saveasync-callback--asyncresult-)|Асинхронно сохраняет элемент.|
||[saveAsync (Options: Office. Асинкконтекстоптионс, callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#saveasync-options--callback--asyncresult-)|Асинхронно сохраняет элемент.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[notificationMessages](/javascript/api/outlook/outlook.messageread#notificationmessages)|Получает сообщения уведомления для элемента.|
|[NotificationMessageDetails](/javascript/api/outlook/outlook.notificationmessagedetails)|[icon](/javascript/api/outlook/outlook.notificationmessagedetails#icon)|Ссылка на значок, определенный в манифесте в разделе `Resources`.|
||[key](/javascript/api/outlook/outlook.notificationmessagedetails#key)|Идентификатор для сообщения уведомления.|
||[message](/javascript/api/outlook/outlook.notificationmessagedetails#message)|Текст сообщения уведомления.|
||[сохраняемого](/javascript/api/outlook/outlook.notificationmessagedetails#persistent)|Указывает, должно ли сообщение быть постоянным.|
||[type](/javascript/api/outlook/outlook.notificationmessagedetails#type)|Задает `ItemNotificationMessageType` сообщение.|
|[NotificationMessages](/javascript/api/outlook/outlook.notificationmessages)|[addAsync (Key: строка, Жсонмессаже: NotificationMessageDetails, Options?: Office. Асинкконтекстоптионс, обратный вызов?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#addasync-key--jsonmessage--options--callback--asyncresult-)|Добавляет уведомление к элементу.|
||[getAllAsync (Options?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult<NotificationMessageDetails [] >) => void)](/javascript/api/outlook/outlook.notificationmessages#getallasync-options--callback--asyncresult-)|Возвращает все ключи и сообщения для элемента.|
||[removeAsync (Key: String, Options?: Office. Асинкконтекстоптионс, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#removeasync-key--options--callback--asyncresult-)|Удаляет сообщение уведомления для элемента.|
||[replaceAsync (Key: строка, Жсонмессаже: NotificationMessageDetails, Options?: Office. Асинкконтекстоптионс, обратный вызов?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#replaceasync-key--jsonmessage--options--callback--asyncresult-)|Заменяет сообщение уведомления с заданным ключом на другое сообщение.|