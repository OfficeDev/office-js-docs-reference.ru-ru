# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Outlook надстройки API подмножество API JavaScript для Office включает объекты, методы, свойства, и надстройки событий, можно использовать в Outlook.

> [!NOTE]
> Эта документация является **предварительной версии** [задать требования](/javascript/office/requirement-sets/outlook-api-requirement-sets). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки. Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.

Наборы требований Preview включает в себя все возможности [требование задать 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - добавлен новый объект, представляющий свойства элемента встречи или сообщения в общей папке, календаря или почтового ящика.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-). Новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением (`allowEvent`). Это значение используется для отмены выполнения события.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - добавлен новый метод, который подключает файла из base64 кодирования в сообщении или встрече.
- [Office.context.mailbox.item.getInitializationContextAsync.](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - добавлен новый метод, который возвращает объект, который представляет sharedProperties встречи или сообщения.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Добавлен доступ к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - добавлено новое перечисление флаг бит, указывающее делегированных разрешений.
- [Office.EventType](/javascript/api/office/office.eventtype) - изменены для поддержки событий OfficeThemeChanged посредством добавления `OfficeThemeChanged` запись.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)