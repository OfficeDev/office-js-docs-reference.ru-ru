---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, Outlook надстройка доступна в общих папках и сценариях общих почтовых ящиков.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf0066639b6d40bb2a460955b14b79fae6f45e8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487241"
---
# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Определяет, доступна ли надстройка Outlook в общих почтовых ящиках (в настоящее время в предварительном просмотре) и общих папках (т. е. в сценариях делегирования доступа). Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). По умолчанию для него установлено значение *false*.

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.8. См [клиенты и платформы](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.8](/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8)

Ниже приводится пример элемента **SupportsSharedFolders** .

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
  </VersionOverrides>
</VersionOverrides>
...
```
