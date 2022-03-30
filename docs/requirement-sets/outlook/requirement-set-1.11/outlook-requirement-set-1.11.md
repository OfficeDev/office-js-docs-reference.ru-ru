---
title: Outlook API надстройки 1.11
description: Набор требований 1.11 для Outlook API надстройки.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 457657953384f39693a68c5b4ebc23ab606c8648
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503101"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook API надстройки 1.11

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

## <a name="whats-new-in-111"></a>Что нового в 1.11?

Набор требований 1.11 включает все функции набора [требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые события [для активации на основе событий](/office/dev/add-ins/outlook/autolaunch#supported-events).
- Добавлены API SessionData.

### <a name="change-log"></a>Журнал изменений

- [Добавлена Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties): добавляет новое свойство для управления данными сеанса элемента в режиме Compose.
- [Добавлены Office. SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true). Добавляет новый объект, который представляет данные сеанса элемента составить.
- Добавлены новые события [для активации на основе событий](/office/dev/add-ins/outlook/autolaunch#supported-events): добавлена поддержка следующих событий.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- [Добавлены Office. AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true): добавляет объект, поддерживаючий событие`OnAppointmentTimeChanged`.
- [Добавлены Office. AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true): добавляет объект, поддерживаючий события `OnAppointmentAttachmentsChanged` и события`OnMessageAttachmentsChanged`.
- [Добавлены Office. InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true): добавляет объект, который поддерживает `OnInfoBarDismissClicked` событие.
- [Добавлены Office. RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true): добавляет объект, который поддерживает события `OnAppointmentAttendeesChanged` и события`OnMessageRecipientsChanged`.
- [Добавлены Office. RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true): добавляет объект, поддерживающий `OnAppointmentRecurrenceChanged` событие.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [Наборы обязательных элементов и поддерживаемые клиенты](../outlook-api-requirement-sets.md)
