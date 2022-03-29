---
title: Групповой элемент в файле манифеста
description: Определяет группу элементов управления пользовательским интерфейсом на вкладке.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2dcd853075e770b8a816a2f898a619364c09addf
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486976"
---
# <a name="group-element"></a>Элемент Group

Определяет группу элементов управления пользовательским интерфейсом на вкладке. На настраиваемой вкладке надстройка может создавать несколько групп. Надстройка может создать не более одной специальной вкладки.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) 

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Да  | Уникальный идентификатор группы.|

### <a name="id-attribute"></a>Атрибут id

Обязательный аргумент. Уникальный идентификатор для группы. Это строка длиной до 125 символов. Это должно быть уникальным для всех элементов Group в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)      | Да |  Метка для группы.  |
|  [Icon](icon.md)      | Да |  Изображение для группы. Не поддерживается Outlook надстройки. |
|  [Control](#control)    | Нет |  Представляет объект Control. Может быть ноль или больше.  |
|  [OfficeControl](#officecontrol)  | Нет | Представляет один из встроенных элементов Office элементов управления. Может быть ноль или больше. Не поддерживается Outlook надстройки.|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Нет |  Указывает, должна ли группа отображаться в сочетаниях приложений и платформ, поддерживаюх настраиваемые контекстные вкладки. Не поддерживается Outlook надстройки. |

### <a name="label"></a>Label

Обязательный элемент. Метка группы. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources](resources.md) .

### <a name="icon"></a>Icon

Обязательный элемент. Если вкладка содержит большое количество групп и окно программы повторно, указанное изображение может отображаться вместо этого.

> [!NOTE]
> Этот элемент не поддерживается в Outlook надстройки.

### <a name="control"></a>Элемент управления

Необязательный, но если нет, то должен быть хотя бы один **OfficeControl**. Сведения о типах поддерживаемых элементов управления см. в [элементе Control](control.md) . Порядок управления и  **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon**.

```xml
<Group id="Contoso.CustomTab1.group1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button1">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

Необязательный, но если нет, должен быть хотя бы один **контроль**. Включаем один или несколько встроенных элементов Office в группу с `<OfficeControl>` элементами. Атрибут `id` указывает ID встроенного Office управления. Чтобы найти ID элементов управления, см. в рублях [Find the IDs of controls and control groups](/office/dev/add-ins/design/built-in-button-integration#find-the-ids-of-controls-and-control-groups). Порядок управления и  **OfficeControl** в манифесте взаимозаменяем, и они могут быть взаимозаменяемыми, если существует несколько элементов, но все они должны быть ниже элемента **Icon**.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.3](../requirement-sets/common/add-in-commands-requirement-sets.md)

> [!NOTE]
> Этот элемент не поддерживается в Outlook надстройки.

```xml
<Group id="Contoso.CustomTab2.group2">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

Необязательный (boolean). Указывает, будет ли **группа** скрыта в сочетаниях приложений и платформ, поддерживаюх API, который устанавливает настраиваемую контекстную вкладку на ленту во время запуска. Значение по умолчанию, если не присутствует, является `false`. Если используется, **OverriddenByRibbonApi** должен быть первым *ребенком* **группы**. Дополнительные сведения см. в [переопределенномByRibbonApi](overriddenbyribbonapi.md).

> [!NOTE]
> Этот элемент не поддерживается в Outlook надстройки.

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.CustomTab3">
    <Group id="Contoso.CustomTab3.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
