# <a name="outlook-add-in-api-requirement-set-14"></a>Набор обязательных элементов API для надстройки Outlook 1.4

Outlook надстройки API подмножество API JavaScript для Office включает объекты, методы, свойства, и надстройки событий, можно использовать в Outlook.

> [!NOTE]
> В этой документации — [Задайте требование](/javascript/office/requirement-sets/outlook-api-requirement-sets) отличный от новейшие наборы требований.

## <a name="whats-new-in-14"></a>Новые возможности в версии 1.4

Набор обязательных элементов 1.4 включает все возможности [набора обязательных элементов версии 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-). Отображает диалоговое окно в ведущем приложении Office.
- Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-). Доставляет сообщение из диалогового окна родительской странице.
- Добавлены объекта [диалогового окна](/javascript/api/office/office.dialog) : объект, который возвращается при [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) вызывается метод.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)