---
title: Элемент PhoneSettings в файле манифеста
description: Элемент PhoneSettings указывает исходные параметры расположения и управления, которые применяются при применении почтовой надстройки на телефоне.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 09b53de93c4f75694c24c8649cfcdead5b51910d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487472"
---
# <a name="phonesettings-element"></a>Элемент PhoneSettings

Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.

> [!IMPORTANT]
> Элемент `PhoneSettings` доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и Outlook 2013 Windows. Чтобы поддерживать Outlook Android и iOS, см. в приложении [Надстройки для Outlook Mobile](/office/dev/add-ins/outlook/outlook-mobile-addins).

**Тип надстройки:** почтовая

## <a name="syntax"></a>Синтаксис

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Содержится в

[Form](form.md)

