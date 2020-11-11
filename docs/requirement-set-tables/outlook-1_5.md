| Класс | Поля | Описание |
|:---|:---|:---|
|[Mailbox](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync (eventType: Office. EventType \| строка, обработчик: (тип: Office. EventType) => void, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|Добавляет обработчик для поддерживаемого события.|
||[getCallbackTokenAsync (Options: Office. Асинкконтекстоптионс & {опускайте?: Boolean}, callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|Получает строку, содержащую маркер, используемый для вызова REST API или веб-служб Exchange (EWS).|
||[Опускайте](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync (eventType: Office. EventType \| строка, Options?: Office. асинкконтекстоптионс, callback?: (AsyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|Удаляет обработчиков для поддерживаемого типа события.|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.|
