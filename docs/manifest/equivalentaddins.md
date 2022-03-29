---
title: Элемент EquivalentAddins в файле манифеста
description: Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 406fb43b74afa31b6a325f7e3d3648d89093aadc
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487277"
---
# <a name="equivalentaddins-element"></a>Элемент EquivalentAddins

Указывает обратную совместимость с эквивалентной надстройки COM, XLL или обоих.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>Содержится в

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>Должен содержать

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)