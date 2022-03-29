---
title: Outlook API надстройки 1.10
description: Набор требований 1.10 для Outlook API надстройки.
ms.date: 11/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a62c5af38b230af14b18a9bbe1ce1db6a57a971
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487436"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook API надстройки 1.10

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../outlook-api-requirement-sets.md).

## <a name="whats-new-in-110"></a>Что нового в 1.10?

Набор требований 1.10 включает все функции набора [требований 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для [активации на](/office/dev/add-ins/outlook/autolaunch) основе событий и функций подписи почты.
- Добавлена поддержка [объекта OfficeRuntime.служба хранилища](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) с функцией активации на основе событий.
- Добавлена возможность включить настраиваемые действия в сообщение уведомления.

### <a name="change-log"></a>Журнал изменений

- [Добавлена точка расширения LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent). Добавлен новый поддерживаемый тип ExtensionPoint. Он настраивает функции активации на основе событий.
- Добавлен [элемент манифеста LaunchEvents](/javascript/api/manifest/launchevents): добавляет элемент манифеста для поддержки настройки функции активации на основе событий.
- [Измененный элемент манифеста runtimes](/javascript/api/manifest/runtimes): добавляет Outlook поддержку. Он ссылается на ФАЙЛЫ HTML и JavaScript, необходимые для функций активации на основе событий.
- Добавлен [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#outlook-office-body-setsignatureasync-member(1)): добавляет новую функцию в `Body` объект. Он добавляет или заменяет подпись в корпусе элемента в режиме Compose.
- Добавлен [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods): добавляется новая функция, которая отключает подпись клиента для отправки почтового ящика в режиме Compose.
- Добавлен [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)): добавляется новая функция, которая получает тип композитного сообщения в режиме Compose.
- [Добавлена Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods): добавляется новая функция, которая проверяет, включена ли подпись клиента на элементе в режиме Compose.
- [Добавлены Office. MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype?view=outlook-js-1.10&preserve-view=true): добавляет новый список. Он представляет тип настраиваемого действия в уведомлении.
- [Добавлены Office. MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true): добавляет новый список, доступный в режиме Compose.
- [Добавлены Office. MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.10&preserve-view=true): добавляет новый тип в `ItemNotificationMessageType` список. Оно представляет сообщение уведомления с пользовательским действием.
- [Добавлены Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction?view=outlook-js-1.10&preserve-view=true): добавляет новый объект, чтобы можно было определить настраиваемые действия для уведомления`InsightMessage`.
- [Добавлены Office. NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.10&preserve-view=true#outlook-office-notificationmessagedetails-actions-member): `InsightMessage` добавляет новое свойство, которое позволяет добавлять уведомление с помощью настраиваемой меры.
- Изменение [OfficeRuntime.служба хранилища](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true): добавляет Outlook поддержку, но только с функцией активации на основе событий.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Наборы обязательных элементов и поддерживаемые клиенты](../outlook-api-requirement-sets.md)
