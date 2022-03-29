---
title: Элемент OfficeMenu в файле манифеста
description: Элемент OfficeMenu определяет коллекцию элементов управления, которые будут добавлены в Office контекста.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: fb496953286ef7b4e359b5f9271b5f28bce763ec
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487056"
---
# <a name="officemenu-element"></a>Элемент OfficeMenu

Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

## <a name="attributes"></a>Атрибуты

| Атрибут            | Обязательный | Описание                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Да      | Тип определяемого элемента OfficeMenu.|

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Управление типом Меню](control-menu.md)    | Да |  Коллекция из одного или нескольких объектов Control.  |

## <a name="xsitype"></a>xsi:type

Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.

- `ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.
- `ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel.

## <a name="example"></a>Пример

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.myMenu">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />
          </Action>
        </Item>
      </Items>
    </Control>
</OfficeMenu>
```
