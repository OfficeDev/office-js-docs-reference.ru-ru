---
title: Элемент WebApplicationInfo в файле манифеста
description: Справочная документация элемента WebApplicationInfo для Office файлов манифеста надстройок (XML).
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: db4ce3e94420af385c818d98eb6c4930676d075d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486956"
---
# <a name="webapplicationinfo-element"></a>Элемент WebApplicationInfo

Поддерживает единый вход в надстройках Office. Этот элемент содержит сведения для надстройки в качестве следующего:

- Ресурс OAuth *2.0,* на который Office клиентского приложения могут потребоваться разрешения.
- *Клиент* OAuth 2.0, которому могут потребоваться разрешения для Microsoft Graph.

**Тип надстройки:** Области задач, почты, контента

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Контент 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [IdentityAPI 1.3](../requirement-sets/common/identity-api-requirement-sets.md)

> [!NOTE]
> Единый API входов в настоящее время поддерживается для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API удостоверений](../requirement-sets/common/identity-api-requirement-sets.md). Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Microsoft 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить современную проверку подлинности для клиента](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

**WebApplicationInfo** — дочерний элемент элемента [VersionOverrides](versionoverrides.md) в манифесте.  

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Id**    |  Да   |  **Идентификатор** связанной с надстройкой службы, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  **Resource**  |  Да   |  Указывает **URI идентификатора** надстройки, зарегистрированный в конечной точке Azure Active Directory 2.0.|
|  [Scopes](scopes.md)                |  Да  |  Указывает разрешения, необходимые надстройки для ресурса, например Microsoft Graph.  |

## <a name="webapplicationinfo-example"></a>Пример WebApplicationInfo

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
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
