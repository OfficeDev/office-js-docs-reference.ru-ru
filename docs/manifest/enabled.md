---
title: Включен элемент в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: c8167bde639d5d25c4a73a7112dd6efd8ded2d28
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487492"
---
# <a name="enabled-element"></a>Элемент Включен

Указывает, включено ли управление [кнопкой](control-button.md) или [меню](control-menu.md) при запуске надстройки. Элемент **Включен** — это детский элемент [Управления](control.md). Если он опущен, по умолчанию .`true`

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [RibbonApi 1.0](../requirement-sets/common/ribbon-api-requirement-sets.md)

Этот элемент действителен только в Excel, PowerPoint и Word; то есть, `Name` когда атрибутом элемента [Host](host.md) являются "Книга", "Презентация" или "Документ".

Родительский контроль также может быть включен программным образом и отключен. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](/office/dev/add-ins/design/disable-add-in-commands).

## <a name="example"></a>Пример

```xml
<Enabled>false</Enabled>
```
