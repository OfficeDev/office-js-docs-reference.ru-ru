# <a name="outlook-add-in-api-requirement-set-16"></a>Задайте 1.6 требование API надстройки для Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-16"></a>Новые возможности в 1.6?

Наборы требований 1.6 включает в себя все возможности [требование задать 1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Он добавлены следующие возможности.

- Добавлены новые интерфейсы API для контекстной надстройки для получения объекта или регулярное выражение match, что пользователь выбрал для активации надстройки.
- Добавлен новый интерфейс API для открытия форме создания сообщения.
- Добавлена возможность надстройки определить тип учетной записи из почтового ящика пользователя.

### <a name="change-log"></a>Журнал изменений

- Добавлена [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): Добавление новой функции, которая возвращает сущности, обнаруженные в выделенной соответствие пользователь выбрал параметр. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлена [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Добавляет новые функции, которая возвращает строковые значения в выделенной совпадение, соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлена [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Добавляет новую функцию, который открывает новую форму сообщение.
- Добавлена [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Добавляет новый элемент в профиль пользователя, который указывает тип учетной записи пользователя.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)