---
title: Элемент управления типом MobileButton в файле манифеста
description: Определяет кнопку на мобильном устройстве, которая выполняет действие или запускает области задач.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 402a15465487811c0c4c6d8aa02c45b7840d711f
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502646"
---
# <a name="control-element-of-type-mobilebutton"></a>Элемент управления типа MobileButton

Определяет кнопку, которая выполняет действие или запускает области задач, которая отображается только на мобильных платформах.

> [!NOTE]
> В этой статье предполагается знакомство с базовой справочной [статьей Control,](control.md) которая содержит важные сведения о атрибутах элемента.

Кнопка мобильного устройства выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления кнопками `id` должен иметь уникальное значение атрибута среди всех элементов **управления** в манифесте.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Label](#label)     | Да |  Текст для кнопки. |
|  [Icon](icon.md)      | Да |  Изображение для кнопки.         |
|  [Action](action.md)    | Да |  Указание действия, которое предстоит выполнить. **Элементу Control** может быть только один ребенок **действия**. |

### <a name="label"></a>Label

Указывает текст для кнопки с помощью его только атрибута **resid**, который может быть не более 32 символов и должен быть задан к значению атрибута **id** элемента **String** в ребенке **ShortStrings** элемента [Resources](resources.md) .

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="examples"></a>Примеры

В следующем примере кнопка выполняет функцию.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

В следующем примере кнопка отображает области задач.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
