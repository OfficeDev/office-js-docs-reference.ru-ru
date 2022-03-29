---
title: Элемент Script в файле манифеста
description: Элемент Script определяет параметры скрипта, которые настраиваемая функция использует в Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: af63a1c5be397b5a7923715d5739cead7cd15017
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487457"
---
# <a name="script-element"></a>Элемент Script

Определяет параметры сценариев, используемых пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|Элементы  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.|

## <a name="example"></a>Пример

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
