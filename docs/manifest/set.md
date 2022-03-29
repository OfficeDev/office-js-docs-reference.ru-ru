---
title: Элемент Set в файле манифеста
description: Элемент Set указывает Office API JavaScript, задаваемого Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f5e8273e07e1abf959def7eac262264bc922f45
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487249"
---
# <a name="set-element"></a>Элемент Set

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте (то есть элемент требования  к бабушке и дедушке является прямым ребенком [OfficeApp](officeapp.md)), элемент Set указывает набор [](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) требований из API JavaScript Office, необходимый вашей Office надстройке для активации Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Как правнук элемента VersionOverrides

Указывает набор требований из API Office JavaScript, который должен поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [версияOverrides](versionoverrides.md) вступила в силу.[](/office/dev/add-ins/develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- То же, что и элемент ["Требования к бабушке и дедушке](requirements.md) ".

**Связанные с этими наборами требований**:

- То же, что и элемент ["Требования к бабушке и дедушке](requirements.md) ".

## <a name="syntax"></a>Синтаксис

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Содержится в

[Sets](sets.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя [набора требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в элементе родительских [наборов](sets.md) .|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в Office версии и платформы, на которых может быть установлена [надстройка](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in).

