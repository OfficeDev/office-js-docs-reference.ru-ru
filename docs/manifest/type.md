---
title: Тип элемента в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка com надстройка или XLL.
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38063472e403e118baa80ee854f49a1093c1dd3d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487157"
---
# <a name="type-element"></a>Элемент Type

Указывает, является ли эквивалентная надстройка com надстройка или XLL.

**Тип надстройки:** Области задач, настраиваемая функция

## <a name="syntax"></a>Синтаксис

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>Содержится в

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>Значения типа надстройки

Необходимо указать одно из следующих значений элемента `Type` .

- COM. Указывает, что эквивалентная надстройка — это надстройка COM.
- XLL. Указывает, что эквивалентная надстройка является Excel XLL.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)