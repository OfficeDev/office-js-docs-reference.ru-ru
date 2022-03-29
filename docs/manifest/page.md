---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры HTML-страниц, которые настраиваемая функция использует в Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6cd253762c12cd1fffcc38cfacc6f6915600722c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487476"
---
# <a name="page-element"></a>Элемент Page

Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md) 

## <a name="attributes"></a>Атрибуты

Нет

## <a name="child-elements"></a>Дочерние элементы

|  Элемент  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Да  | Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями. |

## <a name="example"></a>Пример

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
