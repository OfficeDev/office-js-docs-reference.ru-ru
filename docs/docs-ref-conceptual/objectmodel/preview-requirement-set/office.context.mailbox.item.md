
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).

##### <a name="requirements"></a>Requirements

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restricted|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | Элемент |
| [bcc](#bcc-recipientsjavascriptapioutlookofficerecipients) | Элемент |
| [body](#body-bodyjavascriptapioutlookofficebody) | Элемент |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Элемент |
| [conversationId](#nullable-conversationid-string) | Элемент |
| [dateTimeCreated](#datetimecreated-date) | Элемент |
| [dateTimeModified](#datetimemodified-date) | Элемент |
| [end](#end-datetimejavascriptapioutlookofficetime) | Элемент |
| [from](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | Элемент |
| [internetMessageId](#internetmessageid-string) | Элемент |
| [itemClass](#itemclass-string) | Элемент |
| [itemId](#nullable-itemid-string) | Элемент |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | Элемент |
| [location](#location-stringlocationjavascriptapioutlookofficelocation) | Элемент |
| [normalizedSubject](#normalizedsubject-string) | Элемент |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | Элемент |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Элемент |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | Элемент |
| [recurrence](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | Элемент |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Элемент |
| [sender](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | Элемент |
| [seriesId](#nullable-seriesid-string) | Элемент |
| [start](#start-datetimejavascriptapioutlookofficetime) | Элемент |
| [subject](#subject-stringsubjectjavascriptapioutlookofficesubject) | Элемент |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Элемент |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Метод |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | Метод |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Метод |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Метод |
| [close](#close) | Метод |
| [displayReplyAllForm](#displayreplyallformformdata) | Метод |
| [displayReplyForm](#displayreplyformformdata) | Метод |
| [getEntities](#getentities--entitiesjavascriptapioutlookofficeentities) | Метод |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | Метод |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | Метод |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | Метод |
| [getRegExMatches](#getregexmatches--object) | Метод |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Метод |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Метод |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | Метод |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | Метод |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | Метод |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Метод |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Метод |
| [removeHandlerAsync](#removehandlerasynceventtype-handler-options-callback) | Метод |
| [saveAsync](#saveasyncoptions-callback) | Метод |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Метод |

### <a name="example"></a>Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>Элементы

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Получает массив вложений для элемента. Только в режиме чтения.

> [!NOTE]
> Определенные типы файлов блокируемых в Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются. Для получения дополнительных сведений см [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Тип:

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>bcc :[Recipients](/javascript/api/outlook/office.recipients)

Получает объект, который предоставляет методы для получения или обновления получателей в строке (Скрытая копия) скрытой копии сообщения. Только в режиме создания.

##### <a name="type"></a>Тип:

*   [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a>body :[Body](/javascript/api/outlook/office.body)

Получает объект, предоставляющий методы для работы с основным текстом элемента.

##### <a name="type"></a>Тип:

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>cc: массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получателей](/javascript/api/outlook/office.recipients)

Предоставляет доступ к «копия» (копия) получателей сообщения. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.

##### <a name="compose-mode"></a>Режим создания

`cc` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **копия** сообщения.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.

Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.

Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Получает дату и время создания элемента. Только в режиме чтения.

##### <a name="type"></a>Тип:

*   Date

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Получает дату и время последнего изменения элемента. Только в режиме чтения.

> [!NOTE]
> Этот член не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

##### <a name="type"></a>Тип:

*   Date

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a>end :Date|[Time](/javascript/api/outlook/office.time)

Получает или задает дату и время окончания встречи.

Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `end` возвращает объект `Date`.

##### <a name="compose-mode"></a>Режим создания

Свойство `end` возвращает объект `Time`.

Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### <a name="type"></a>Тип:

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>от:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[из](/javascript/api/outlook/office.from)

Получает адрес электронной почты отправителя сообщения.

Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> `recipientType` Свойства `EmailAddressDetails` объект в `from` — это свойство `undefined`.

##### <a name="read-mode"></a>Режим чтения

`from` Возвращает свойство `EmailAddressDetails` объекта.

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a>Режим создания

`from` Возвращает свойство `From` объект, который предоставляет метод для получения из значения.

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>Требования

|Requirement|||
|---|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|Создание|

#### <a name="internetmessageid-string"></a>internetMessageId :String

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.

Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

|Тип|Описание|Класс элемента|
|---|---|---|
|Элементы встречи|Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|Элементы сообщения|Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

> [!NOTE]
> Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. `itemId` Свойство не совпадать с Идентификатором, используемым API-Интерфейс REST Outlook или идентификатор записи Outlook. Прежде чем вносить API-Интерфейс REST для звонков с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Для получения дополнительных сведений показано [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

Получает тип элемента, который представляет экземпляр.

Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.

##### <a name="type"></a>Тип:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>location :String|[Location](/javascript/api/outlook/office.location)

Получает или задает место встречи.

##### <a name="read-mode"></a>Режим чтения

Свойство `location` возвращает строку, содержащую сведения о месте встречи.

##### <a name="compose-mode"></a>Режим создания

Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.

##### <a name="type"></a>Тип:

*   String | [Location](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)

Получает сообщения уведомления для элемента.

##### <a name="type"></a>Тип:

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Предоставляет доступ к необязательным участникам события. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.

##### <a name="compose-mode"></a>Режим создания

`optionalAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления необязательных участников собрания.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>Организатор:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[организатора](/javascript/api/outlook/office.organizer)

Получает адрес электронной почты организатора указанного собрания.

##### <a name="read-mode"></a>Режим чтения

`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , который представляет организатором собрания.

##### <a name="compose-mode"></a>Режим создания

`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer) , который предоставляет метод для получения значения Организатор.

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [организатора](/javascript/api/outlook/office.organizer)

##### <a name="requirements"></a>Требования

|Requirement|||
|---|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|Создание|

##### <a name="example"></a>Пример

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(значение null) повторения:[повторения](/javascript/api/outlook/office.recurrence)

Получает или задает шаблон повторения встречи. Получает шаблон повторения приглашения на собрание. Читать и создавать режимы для элементов встречи. В режиме чтения к собранию элементы запроса.

`recurrence` При элемента ряд или экземпляра в цикле свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющиеся встречи или собрания запросы. `null`возвращаются для одного встреч и приглашений на собрания из одного встреч. `undefined`возвращается для сообщений, которые не являются приглашений на собрания.

> Примечание: Приглашений на собрание имеют `itemClass` значение IPM. Schedule.Meeting.Request.

> Примечание: Если объект повторения `null`, это означает, что объект является одной встречи или приглашения на собрание из одной встречи и не является частью серии.

##### <a name="type"></a>Тип:

* [Повторение](/javascript/api/outlook/office.recurrence)

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Предоставляет доступ к обязательным участникам события. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.

##### <a name="compose-mode"></a>Режим создания

`requiredAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления обязательных участников собрания.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.

Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

> [!NOTE]
> `recipientType` Свойства `EmailAddressDetails` объект в `sender` — это свойство `undefined`.

##### <a name="type"></a>Тип:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a>(значение null) seriesId: String

Получает идентификатор серии, к которой принадлежит экземпляр.

В OWA и Outlook `seriesId` возвращает идентификатор веб-служб Exchange (EWS) элемента родительского (ряды), к которому принадлежит этот элемент. Однако в iOS и Android `seriesId` возвращает REST идентификатор родительского элемента.

> [!NOTE]
> Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange. `seriesId` Свойство не идентичен идентификаторы Outlook, используемые API-Интерфейс REST Outlook. Прежде чем вносить API-Интерфейс REST для звонков с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Для получения дополнительных сведений показано [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).

`seriesId` Возвращает свойство `null` для элементов, не имеющих родительских элементов, таких как единый встреч, элементы ряда или собрания запрашивает и возвращает `undefined` для других элементов, которые не являются соответствующие запросы.

##### <a name="type"></a>Тип:

* String

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a>start :Date|[Time](/javascript/api/outlook/office.time)

Получает или задает дату и время начала встречи.

Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).

##### <a name="read-mode"></a>Режим чтения

Свойство `start` возвращает объект `Date`.

##### <a name="compose-mode"></a>Режим создания

Свойство `start` возвращает объект `Time`.

Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### <a name="type"></a>Тип:

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>subject :String|[Subject](/javascript/api/outlook/office.subject)

Получает или задает описание, которое отображается в поле темы элемента.

Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### <a name="read-mode"></a>Режим чтения

Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Режим создания

Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Тип:

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>Чтобы: массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получателей](/javascript/api/outlook/office.recipients)

Предоставляет доступ к получателей в строке **Кому** сообщения. Уровень доступа и тип объекта, зависит от режима текущего элемента.

##### <a name="read-mode"></a>Режим чтения

Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.

##### <a name="compose-mode"></a>Режим создания

`to` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **Кому** сообщения.

##### <a name="type"></a>Тип:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Методы

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Добавляет файл в сообщение или встречу в качестве вложения.

Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры
|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`uri`|String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`|String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.isInline`|Boolean|&lt;Необязательный&gt;|Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`AttachmentSizeExceeded`|Вложение превышает максимальный размер.|
|`FileTypeNotSupported`|Расширение вложения не поддерживается.|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="examples"></a>Примеры

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File, attachmentName, [параметры], [обратного вызова])

Добавляет файл из base64 кодирования в сообщение или встречу в виде вложения.

`addFileAttachmentFromBase64Async` Метод загружает файл из кодировки base64 и подключает ее к элементу в форме создания. Этот метод возвращает идентификатор вложения в объекте AsyncResult.value.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### <a name="parameters"></a>Параметры
|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`base64File`|String||Контент, изображения или файла в электронной почте или событие добавляется в кодировке base64.|
|`attachmentName`|String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.isInline`|Boolean|&lt;Необязательный&gt;|Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`AttachmentSizeExceeded`|Вложение превышает максимальный размер.|
|`FileTypeNotSupported`|Расширение вложения не поддерживается.|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="examples"></a>Примеры

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Добавляет обработчик для поддерживаемого события.

В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно вызвать обработчик. |
| `handler` | Function || Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`. |
| `options` | Объект | &lt;необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Object | &lt;необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | function| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.

С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

Если надстройки Office работает в Outlook Web App, `addItemAttachmentAsync` метод могут прикреплять элементов для элементов, отличных от элемента, который вы изменяете; Однако это не поддерживается и не рекомендуется.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`itemId`|String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`|String||Тема вкладываемого элемента. Максимальная длина — 255 символов.|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|Сообщение или встреча содержат слишком много вложений.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

Закрывает текущий создаваемый элемент.

Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.

> [!NOTE]
> В Outlook в Интернете, если элемент является ли он встречей, и он ранее был сохранен с помощью `saveAsync`, то пользователю будет предложено сохранение, удаление или Отмена даже в том случае, если изменений внесено не было с элемента последнего сохранения.

Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restricted|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`formData`|String &#124; Object||Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.|
|`formData.htmlBody`|String|&lt;необязательно&gt;|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;необязательно&gt;|Массив объектов JSON, представляющих собой вложенные файлы или элементы.|
|`formData.attachments.type`|String||Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.|
|`formData.attachments.name`|String||Строка, содержащая имя вложения, длиной до 255 символов.|
|`formData.attachments.url`|String||Используется, только если свойству `type` задано значение `file`. URI расположения файла.|
|`formData.attachments.isInline`|Boolean||Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`formData.attachments.itemId`|String||Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.|
|`callback`|function|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyAllForm`.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`formData`|String &#124; Object||Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.|
|`formData.htmlBody`|String|&lt;необязательно&gt;|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;необязательно&gt;|Массив объектов JSON, представляющих собой вложенные файлы или элементы.|
|`formData.attachments.type`|String||Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.|
|`formData.attachments.name`|String||Строка, содержащая имя вложения, длиной до 255 символов.|
|`formData.attachments.url`|String||Используется, только если свойству `type` задано значение `file`. URI расположения файла.|
|`formData.attachments.isInline`|Boolean||Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.|
|`formData.attachments.itemId`|String||Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.|
|`callback`|function|&lt;Необязательный&gt;|По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="examples"></a>Примеры

Приведенный ниже код передает строку в функцию `displayReplyForm`.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Возвращает сущности, обнаруженные в тело выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Пример

Этот пример ссылается сущностей контакты в тексте текущего элемента.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в тело выбранного элемента.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|Одно из значений перечисления EntityType.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restricted|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если сущности указанного типа отсутствуют в основной текст элемента, метод возвращает пустой массив. В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

|Значение параметра `entityType`|Тип объектов в возвращаемом массиве|Необходимый уровень разрешений|
|---|---|---|
|`Address`|String|**Restricted**|
|`Contact`|Contact|**ReadItem**|
|`EmailAddress`|String|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**Restricted**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|String|**Restricted**|

Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

##### <a name="example"></a>Пример

Следующем примере показано, как получить доступ к массив строк, представляющих почтовых адресов в тексте текущего элемента.

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`name`|String|Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync([options], [callback])

Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).

> [!NOTE]
> Этот метод поддерживается только Outlook 2016 для Windows (больше, чем 16.0.8413.1000 версии Click-to-Run) и Outlook в Интернете для Office 365.

##### <a name="parameters"></a>Параметры
|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Объект|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function|&lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>В случае успешного выполнения инициализации данных предоставляются в `asyncResult.value` свойства в виде строки.<br/>Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="example"></a>Пример

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Requirements

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

<dl class="param-type">

<dt>Тип</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) пункты (допускает значение NULL) {массива. < String >}

Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Описание|
|---|---|---|
|`name`|String|Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">

<dt>Тип</dt>

<dd>Массив. < String ></dd>

</dl>

##### <a name="example"></a>Пример

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к свойству источника, выделение, поступающих из источников, вызовите `asyncResult.value.sourceProperty`, который может быть либо `body` или `subject`.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="returns"></a>Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Пример

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Тип: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.

Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.

##### <a name="requirements"></a>Requirements

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Чтение|

##### <a name="returns"></a>Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

##### <a name="example"></a>Пример

В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync ([параметры] обратного вызова)

Получает свойства выбранной встречи или сообщения в общей папке, календаря или почтового ящика.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Объект|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Общие свойства предоставляются как [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объект в `asyncResult.value` свойство. Этот объект можно использовать для получения общего свойства элемента.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)|Предварительная версия|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.

Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`callback`|function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`. Этот объект можно использовать для получения, задания и удаление настраиваемых свойств из элемента и сохранение изменений для настраиваемого свойства, задайте обратно на сервер.|
|`userContext`|Объект|&lt;необязательно&gt;|Разработчики могут предоставлять любого объекта, которые следует получить доступ к в функции обратного вызова. Этот объект можно получить доступ с `asyncResult.asyncContext` в функции обратного вызова.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание или чтение|

##### <a name="example"></a>Пример

Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Удаляет вложение из сообщения или встречи.

Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`attachmentId`|String||Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function|&lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.|

##### <a name="errors"></a>Ошибки

|Код ошибки|Описание|
|------------|-------------|
|`InvalidAttachmentId`|Идентификатор вложения не существует.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

Указанный ниже код удаляет вложение с идентификатором "0".

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a>removeHandlerAsync (тип события, обработчик, [параметры], [обратного вызова])

Удаляет обработчик событий для события, поддерживаемые.

В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>Параметры

| Имя | Тип | Атрибуты | Описание |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Событие, которое должно вызвать обработчик. |
| `handler` | Function || Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`. |
| `options` | Объект | &lt;необязательно&gt; | Объектный литерал, содержащий одно или несколько из указанных ниже свойств. |
| `options.asyncContext` | Object | &lt;необязательно&gt; | Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова. |
| `callback` | function| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение |

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Асинхронно сохраняет элемент.

При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.

> [!NOTE]
> Если надстройка вызывает `saveAsync` элемент в режиме создания для получения `itemId` для использования с помощью веб-служб Exchange или интерфейса API REST, необходимо учитывать, что когда Outlook находится в режиме кэширования, он может занять некоторое время до элемента фактически синхронизируется с сервера. Пока элемент синхронизирован с помощью `itemId` возвращает ошибку.

Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.

> [!NOTE]
> Следующие клиенты имеют по-разному для `saveAsync` для встреч в режиме создания:
>
> - Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания. Вызов `saveAsync` собрания в Mac Outlook возвращает ошибку.
> - Outlook в Интернете всегда отправляет приглашение или обновления при `saveAsync` вызван на встречи в режиме создания.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`options`|Объект|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.|
|`callback`|function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>В случае успешного выполнения, идентификатор элемента представлен в `asyncResult.value` свойство.|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="examples"></a>Примеры

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.

##### <a name="parameters"></a>Параметры

|Имя|Тип|Атрибуты|Описание|
|---|---|---|---|
|`data`|String||Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.|
|`options`|Object|&lt;необязательно&gt;|Объектный литерал, содержащий одно или несколько из указанных ниже свойств.|
|`options.asyncContext`|Object|&lt;необязательно&gt;|В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;необязательный&gt;|Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.<br/><br/>Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.<br/><br/>Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.|
|`callback`|функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Требования

|Requirement|Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Создание|

##### <a name="example"></a>Пример

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```