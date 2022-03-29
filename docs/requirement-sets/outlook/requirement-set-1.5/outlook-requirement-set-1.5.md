---
title: Набор обязательных элементов API для надстройки Outlook 1.5
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.5.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: de9c7b8f388e969ef00c6a7e9fb9e589a9a713a9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487549"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Набор обязательных элементов API для надстройки Outlook 1.5

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../outlook-api-requirement-sets.md).

## <a name="whats-new-in-15"></a>Новые возможности в версии 1.5

Набор требований 1.5 включает все функции набора [требований 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [закрепляемых областей задач](/office/dev/add-ins/outlook/pinnable-taskpane).
- Добавлена поддержка вызовов [REST API](/office/dev/add-ins/outlook/use-rest-api).
- Добавлена возможность отметить вложение как встроенное.
- Добавлена возможность закрыть область задач или диалоговое окно.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods). Добавляет обработчик для поддерживаемого события.
- Добавлен [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): удаляет обработчики событий для поддерживаемого типа событий.
- Добавлено свойство [Office.EventType](office.md#eventtype-string). Указывает событие, связанное с обработчиком, и включает поддержку события ItemChanged.
- Добавлен метод [Office.context.mailbox.restUrl](office.context.mailbox.md#properties). Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.
- Изменен метод [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods). Добавлен новый вариант этого метода с новой подписью (`getCallbackTokenAsync([options], callback)`). Исходная версия по-прежнему доступна и осталась без изменений.
- Добавлен метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true#office-office-ui-closecontainer-member(1)).
- Изменен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods). Новое значение в словаре `options` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Наборы обязательных элементов и поддерживаемые клиенты](../outlook-api-requirement-sets.md)
