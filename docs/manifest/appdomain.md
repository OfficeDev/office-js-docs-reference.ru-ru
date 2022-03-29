---
title: Элемент AppDomain в файле манифеста
description: Указывает дополнительные домены, которые используются вашей надстройки и должны доверяться Office.
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: ec0dec56e75d25a2ba79eed9932aa0158be45b7e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487288"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Указывает дополнительный домен, Office должен доверять, в дополнение к домену, указанному в [элементе SourceLocation](sourcelocation.md). Указание домена имеет такие эффекты:

- Это позволяет открывать страницы, маршруты или другие ресурсы домена непосредственно в корневой области задач надстройки на настольных Office платформах. (Указание домена в **AppDomain** не требуется для Office в Интернете или для открытия ресурса в IFrame, равно как и для открытия ресурса в диалоговом окрашиваемом [API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins) диалогов.)
- Он позволяет страницам в домене Office.js API-вызовы из IFrames в надстройки.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain.com</AppDomain>`).
> 2. Если для домена имеется явный порт, включай его (например,).`<AppDomain>https://myappdomain.com:9999</AppDomain>`
> 3. Если нужно доверять поддомену, включай его (например,).`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` Subdomain и `mysubdomain.mydomain.com` являются `mydomain.com` различными доменами. Если обоим нужно доверять, то оба должны быть в отдельных **элементах AppDomain** .
> 4. Перечисление того же домена, что и домен, указанный в [элементе SourceLocation](sourcelocation.md) , не имеет эффекта и может вводить в заблуждение. В частности, при разработке `localhost`не требуется создавать элемент **AppDomain** для `localhost`.
> 5. Не включайте в домен какие-либо сегменты URL-адреса. Например, не включайте полный URL-адрес страницы.
> 6. Не *помещай* закрываю черту "/", на значение.

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Замечания

Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).
