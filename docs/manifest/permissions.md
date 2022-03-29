---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions указывает уровень доступа к API для Office надстройки.
ms.date: 06/26/2020
ms.localizationpriority: medium
ms.openlocfilehash: 48369c71b94c286d357868bc1f0620cb09b58878
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487049"
---
# <a name="permissions-element"></a>Элемент Permissions

Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

Для надстроек области задач и контентных надстроек:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см. в материалах [Requesting permissions for API use in content and task pane add-in and](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) [Understanding Outlook надстройки](/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions).
