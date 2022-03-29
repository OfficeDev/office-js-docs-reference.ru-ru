---
title: Элемент Version в файле манифеста
description: Элемент Version указывает Office надстройки.
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486957"
---
# <a name="version-element"></a>Элемент Version

Указывает версию надстройки Office. Номер версии может быть 1, 2, 3 или 4 частей (например, n, n.n, n.n.n или n.n.n.n.).

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Каждая часть номера версии может быть не более 5 цифр.
