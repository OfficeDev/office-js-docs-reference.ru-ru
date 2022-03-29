---
title: Элемент DesktopSettings в файле манифеста
description: Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 9393871e56f686b710ffd0031f93e776f362a89d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487377"
---
# <a name="desktopsettings-element"></a>Элемент DesktopSettings

Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.

> [!IMPORTANT]
> Элемент `DesktopSettings` доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и Outlook 2013 Windows.

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
