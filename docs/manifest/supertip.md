---
title: Элемент Supertip в файле манифеста
description: Элемент Supertip определяет богатый инструментарий (как название, так и описание).
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d4e59970af7c41f5303ec9cf94ef70e1373d1a59
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64501777"
---
# <a name="supertip"></a>Supertip

Определяет форматированную подсказку (элементы Title и Description). Он используется как с помощью [элементов управления кнопками,](control-button.md) так и [элементов управления меню](control-menu.md).

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0
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
| [Title](#title) | Да | Текст подсказки. |
| [Description](#description) | Да | Описание подсказки.<br>**Примечание**. (Outlook) поддерживаются только Windows и mac-клиенты. |

### <a name="title"></a>Title

Обязательный. Текст суперподсказки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources](resources.md) .

### <a name="description"></a>Описание

Обязательный. Описание суперподсказки. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **String** в **элементе LongStrings** в [элементе Resources](resources.md) .

> [!NOTE]
> Для Outlook только клиенты Windows и Mac поддерживают элемент **Description**.

## <a name="example"></a>Пример

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
