---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация элемента VersionOverrides для Office файлов манифеста надстройок (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 56186059e73ab1bd56c1bfd22fb7e03bb7d524eb
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487140"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте. Его детская разметка может переопределять часть разметки в базовом манифесте (или в родительской **версии VersionOverrides**). **VersionOverrides** — это детский элемент корневого элемента [OfficeApp](officeapp.md) в манифесте или родительского **элемента VersionOverrides** . Этот элемент поддерживается в схеме манифеста v1.1 и более поздней версии, но определяется в отдельных схемах VersionOverrides.

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен VersionOverrides схемы. Допустимые значения различаются `<VersionOverrides>` в зависимости от **значения xsi:type** этого элемента и **значения xsi:type** родительского `<OfficeApp>` элемента. Ниже [см. значения Пространства](#namespace-values) имен.|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространства имен

Ниже перечислены необходимое значение атрибута **xmlns** в зависимости от **значения xsi:type** корневого `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны быть `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.
- **ContentApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны быть `http://schemas.microsoft.com/office/contentappversionoverrides`.
- **MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому значение **xmlns** `<VersionOverrides>` зависит от значения **xsi:type** этого элемента:
  - Когда **xsi:type** , `VersionOverridesV1_0`**xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides`.
  - Когда **xsi:type** , `VersionOverridesV1_1`**xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> В настоящее время Outlook 2016 поддерживает схему VersionOverrides v1.1 `VersionOverridesV1_1` и тип.

## <a name="variant-schemas"></a>Схемы вариантов

Для каждого из возможных значений **xmlns** существует отдельная схема, поэтому каждая из них имеет отдельную справочную страницу.

- [Область задач VersionOverrides 1.0](versionoverrides-1-0-taskpane.md)
- [Содержимое VersionOverrides 1.0](versionoverrides-1-0-content.md)
- [Почта VersionOverrides 1.0](versionoverrides-1-0-mail.md)
- [Почта VersionOverrides 1.1](versionoverrides-1-1-mail.md)
