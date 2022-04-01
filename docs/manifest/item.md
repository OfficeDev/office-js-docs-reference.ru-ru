---
title: Элемент Элемента в файле манифеста
description: Указывает элемент в меню.
ms.date: 03/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: a2ea0d1dd8d0b01f91340c61ebf2e21abf598b84
ms.sourcegitcommit: 07f76f08c7952f0777a7aa163fa734f88f6dee9b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/01/2022
ms.locfileid: "64604249"
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
|  [Label](#label)     | Да |  Текст элемента меню. |
|  [Supertip](supertip.md)  | Да |  Супертип элемента меню.    |
|  [Icon](icon.md)      | Нет |  Изображение элемента меню.         |
|  [Action](action.md)    | Да |  Указание действия, которое предстоит выполнить. Элемент Item может быть только  одним ребенком **action**.  |
|  [Enabled](enabled.md)    | Нет |  Указывает, включен ли элемент меню при запуске надстройки.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должен ли элемент меню отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Если используется, он должен быть первым *элементом* ребенка. |

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