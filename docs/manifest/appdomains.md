---
title: Элемент AppDomains в файле манифеста
description: Перечисляет все `SourceLocation` домены в дополнение к домену, указанному в элементе, Office надстройка будет использовать и должна доверяться Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: 6bf1785cf11e31648d9bc69e101cd5a5cf3ecb9f
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487690"
---
# <a name="appdomains-element"></a>Элемент AppDomains

Перечислены все домены, `SourceLocation` в дополнение к домену, указанному в элементе, который будет Office надстройка и которому следует доверять Office. Это позволяет страницам в доменах звонить Office.js API из IFrames в надстройки и имеет другие эффекты. Для каждого дополнительного домена укажите элемент **AppDomain**.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Существуют ограничения по значению элемента **AppDomain** . Дополнительные сведения см. в [приложении AppDomain](appdomain.md).

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

[AppDomain](appdomain.md)

## <a name="remarks"></a>Примечания

По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md). Этот элемент не может быть пустым.
