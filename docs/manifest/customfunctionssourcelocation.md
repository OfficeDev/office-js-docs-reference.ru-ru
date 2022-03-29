---
title: Элемент SourceLocation для настраиваемой функции в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: ae44e513e94fd941e3f3a3593edf06f8add89a56
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487384"
---
# <a name="sourcelocation-element-custom-functions"></a>Элемент SourceLocation (настраиваемые функции)

Определяет расположение ресурса, необходимого элементам **Script** или **Page**, используемым пользовательскими функциями в Excel.

> [!IMPORTANT]
> Эта статья относится только к **SourceLocation** , которая является ребенком элементов **Page** или **Script** . Сведения о элементе **SourceLocation базового** манифеста см. в [sourceLocation](sourcelocation.md).

**Тип надстройки:** Настраиваемая функция

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Да      | Имя ресурса URL-адреса, определенного в разделе **Ресурсы** в манифесте. Может быть не более 32 символов. |

## <a name="child-elements"></a>Дочерние элементы

Нет

## <a name="example"></a>Пример

```xml
<SourceLocation resid="pageURL"/>
```
