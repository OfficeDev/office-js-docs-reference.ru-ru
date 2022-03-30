---
title: Элемент Sets в файле манифеста
description: Элемент Sets указывает минимальный набор API Office JavaScript, который требуется Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 841b867ef4f40c63e114cf9b1dd5ce7ae245690c
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64501771"
---
# <a name="sets-element"></a>Элемент Sets

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте (то есть элемент родительских требований является прямым ребенком [OfficeApp](officeapp.md)), элемент **Sets** указывает минимальный подмножество требований к API javaScript Office (наборы требований [), которые](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets) необходимы вашей Office надстройки для активации Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>Как внук элемента VersionOverrides

Указывает минимальный набор Office API JavaScript [(](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets)наборы требований), которые должны поддерживаться версией и платформой Office (например, Windows, Mac, web, iOS или iPad) для того, чтобы [версияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- То же самое, что и элемент [родительских требований](requirements.md) .

**Связанные с этими наборами требований**:

- То же самое, что и элемент [родительских требований](requirements.md) .

## <a name="syntax"></a>Синтаксис

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Содержится в

[Requirements](requirements.md)

## <a name="can-contain"></a>Может содержать

[Set](set.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Указывает значение атрибута **MinVersion** по умолчанию для [всех элементов](set.md) набора детей. Значение по умолчанию: "1.1".|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в Office версии и платформы, на которых может быть установлена [надстройка](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in).

