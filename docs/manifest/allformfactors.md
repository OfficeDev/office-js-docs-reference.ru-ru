---
title: Элемент AllFormFactors в файле манифеста
description: Указывает параметры всех форм-факторов для надстройки.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 25d9b815cf7db440abce2c8859c9ebb15ea5df1a
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487400"
---
# <a name="allformfactors-element"></a>Элемент AllFormFactors

Указывает параметры всех форм-факторов для надстройки. В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**. Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

> [!NOTE]
> Этот элемент поддерживается только в Excel Windows, на Mac и в Интернете. Он не поддерживается в других Office приложениях или на iOS или Android.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Да |  Определяет, где предоставляются функции надстройки. |

## <a name="allformfactors-example"></a>Пример использования AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
