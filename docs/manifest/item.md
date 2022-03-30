---
title: Элемент Элемента в файле манифеста
description: Указывает элемент в меню.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b717bcb3041a0164716b50116f387f0b581f2ea
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502632"
---
# <a name="item-element"></a>Элемент Item

Указывает элемент в меню.

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

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)     | Да |  Текст для кнопки. |
|  [Supertip](supertip.md)  | Да |  Суперподсказка для кнопки.    |
|  [Icon](icon.md)      | Да |  Изображение для кнопки.         |
|  [Action](action.md)    | Да |  Указание действия, которое предстоит выполнить. Элемент Item может быть только  одним ребенком **action**.  |
|  [Enabled](enabled.md)    | Нет |  Указывает, включен ли контроль при запуске надстройки.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли кнопка отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Если используется, он должен быть первым *элементом* ребенка. |

### <a name="label"></a>Label

Указывает текст для кнопки с помощью его только атрибута **resid**, который может быть не более 32 символов и должен быть задан к значению атрибута **id** элемента **String** в ребенке **ShortStrings** элемента [Resources](resources.md) .

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

## <a name="examples"></a>Примеры

Примеры см. в [пункте Управление типом Меню](control-menu.md).