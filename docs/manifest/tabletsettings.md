---
title: Элемент TabletSettingst в файле манифеста
description: Элемент TabletSettings указывает параметры управления, которые применяются при добавлении почты на планшете.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3b2af4a220e01017c268125c3db3edc61cf7ec82
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486865"
---
# <a name="tabletsettings-element"></a>Элемент TabletSettingst

Задает параметры управления, которые применяются при использовании вашей почтовой надстройки на планшете.

> [!IMPORTANT]
> Элемент `TabletSettings` доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и Outlook 2013 Windows. Чтобы поддерживать Outlook Android и iOS, см. в приложении [Надстройки для Outlook Mobile](/office/dev/add-ins/outlook/outlook-mobile-addins).

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
