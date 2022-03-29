---
title: Элемент DefaultSettings в файле манифеста
description: Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: f86419bf6a3c18e3ae62091c53b1e8f82c706fb1
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487065"
---
# <a name="defaultsettings-element"></a>Элемент DefaultSettings

Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.

**Тип надстройки:** контентные надстройки и надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Замечания

Расположение источника и другие параметры в **элементе DefaultSettings** применяются только к надстройке контента и области задач. Для надстройок почты в элементе [FormSettings](formsettings.md) указывается расположение исходных файлов по умолчанию и другие параметры по умолчанию.
