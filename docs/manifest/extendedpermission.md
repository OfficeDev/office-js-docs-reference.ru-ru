---
title: Элемент ExtendedPermission в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 12020d15cf5ec235541aa27e3b0a2bc0db771d6e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487061"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` элемент

Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции. Элемент `ExtendedPermission` является детским элементом [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.9. См [клиенты и платформы](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.9](/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9)

## <a name="available-extended-permissions"></a>Доступные расширенные разрешения

Ниже приводится доступное значение.

|Доступное значение|Описание|Hosts|
|---|---|---|
|`AppendOnSend`|Объявляет, что надстройка использует Office[. API Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-appendonsendasync-member(1)).|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` пример

Ниже приводится пример элемента `ExtendedPermission` .

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>Содержится в

[ExtendedPermissions](extendedpermissions.md)
