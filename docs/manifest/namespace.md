---
title: Элемент Namespace в файле манифеста
description: Элемент Namespace определяет пространство имен, которое настраиваемая функция использует в Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf05f7c950e2be03baf056284bdcb0bd6a5b5d0e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487642"
---
# <a name="namespace-element"></a>Элемент Namespace

Определяет пространство имен, используемых пользовательской функцией в Excel.

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Нет  | Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md). Может быть не более 32 символов. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<Namespace resid="namespace" />
```
