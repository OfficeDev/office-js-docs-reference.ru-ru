---
title: Элемент Scopes в файле манифеста
description: Элемент Scopes содержит разрешения, необходимые для подключения надстройки к внешнему ресурсу.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3b67977421a6ed0da205a051d0ecaf480c96fd5c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487025"
---
# <a name="scopes-element"></a>Элемент Scopes

Содержит разрешения, необходимые надстройки для внешнего ресурса, например Microsoft Graph. Когда microsoft Graph является ресурсом, AppSource использует элемент Scopes для создания диалоговое окно согласия. Когда пользователи устанавливают надстройку из Магазина, им предлагается предоставить ей указанные разрешения на доступ к данным Microsoft Graph.

**Тип надстройки:** Области задач, почты, контента

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Контент 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [IdentityAPI 1.3](../requirement-sets/common/identity-api-requirement-sets.md)

**Области —** это детский элемент элемента [WebApplicationInfo](webapplicationinfo.md) в манифесте.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Scope**                |  Да     |   Имя разрешения; например, Files.Read.All или профиль. |

## <a name="example"></a>Пример

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
