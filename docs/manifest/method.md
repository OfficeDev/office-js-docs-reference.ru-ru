---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из API Office JavaScript, который требуется Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9db4f9b2063d3bda4505abb8fc426059a1081f23
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487368"
---
# <a name="method-element"></a>Элемент Method

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте (то есть элемент требования  к бабушке и дедушке является прямым ребенком [OfficeApp](officeapp.md)), элемент Method указывает отдельный метод из API JavaScript Office, который необходим вашей Office надстройке для активации Office.

**Тип надстройки:** контентные надстройки и надстройки области задач.

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Как правнук элемента VersionOverrides

Указывает отдельный метод из API Office JavaScript, который должен поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [ВерсияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- То же, что и элемент ["Требования к бабушке и дедушке](requirements.md) ".

**Связанные с этими наборами требований**:

- То же, что и элемент ["Требования к бабушке и дедушке](requirements.md) ".

## <a name="syntax"></a>Синтаксис

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Содержится в

[Методы](methods.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы указать метод `getSelectedDataAsync` , необходимо указать `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Примечания

Элементы **Methods** и **Method** не поддерживаются почтовыми надстройки при их использования в базовом манифесте. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в Office [API JavaScript](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).
