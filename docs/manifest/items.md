---
title: Элемент Items в файле манифеста
description: Указывает элементы в меню.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 15702e2cdbeb6d8df6d714d59206b6e03143b6a4
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503395"
---
# <a name="items-element"></a>Элемент Items

Указывает элементы в меню.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="syntax"></a>Синтаксис

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>Содержится в

[Элемент управления меню типа](control-menu.md)

## <a name="must-contain"></a>Должен содержать

[Элемент](item.md)

## <a name="examples"></a>Примеры

Примеры см. в [пункте Управление типом Меню](control-menu.md).