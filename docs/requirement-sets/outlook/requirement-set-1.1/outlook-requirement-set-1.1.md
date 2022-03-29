---
title: Набор требований к API надстройки Outlook 1.1
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.1.
ms.date: 12/17/2019
ms.localizationpriority: medium
ms.openlocfilehash: 71c921cfe60bf0a50fb5a2a69342e46966ad3ec9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487570"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Набор требований к API надстройки Outlook 1.1

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки. Outlook API JavaScript 1.1 (почтовый ящик 1.1) является первой версией API.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../outlook-api-requirement-sets.md).

## <a name="whats-new-in-11"></a>Новые возможности в версии 1.1

Набор требований 1.1 включает все наборы общих [API](../../common/office-add-in-requirement-sets.md), поддерживаемые в Outlook. В нем надстройки получили возможность доступа к тексту сообщений и встреч, а также возможность изменения текущего элемента.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true). Предоставляет методы для добавления и изменения содержимого элемента в надстройке Outlook.
- Добавлен объект [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true). Предоставляет методы, позволяющие получить и задать место проведения собрания в надстройке Outlook.
- Добавлен объект [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true). Предоставляет методы, позволяющие получить и задать получателей для встречи или сообщения в надстройке Outlook.
- Добавлен объект [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true). Предоставляет методы, позволяющие получить и задать тему встречи или сообщения в надстройке Outlook.
- Добавлен объект [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true). Предоставляет методы, позволяющие получить и задать время начала и окончания собрания в надстройке Outlook.
- Добавлен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods). Добавляет файл в сообщение или встречу в качестве вложения.
- Добавлен метод [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods). Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.
- Добавлен метод [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods). Удаляет вложение из сообщения или встречи.
- Добавлено свойство [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties). Получает объект, предоставляющий методы для работы с текстом элемента.
- [Добавлена Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) строки сообщения.
- Добавлено свойство [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true). Указывает тип получателя для встречи.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Наборы обязательных элементов и поддерживаемые клиенты](../outlook-api-requirement-sets.md)
