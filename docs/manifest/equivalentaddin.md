---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 154e69367b00a49c896d9fb108d621ba212c4580
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486988"
---
# <a name="equivalentaddin-element"></a>Элемент EquivalentAddin

Указывает обратную совместимость для эквивалентной надстройки COM или XLL.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Тип надстройки:** Области задач, почты, настраиваемой функции

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="syntax"></a>Синтаксис

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>Содержится в

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>Должен содержать

[Тип](type.md)

## <a name="can-contain"></a>Может содержать

[ProgId](progid.md)
 [FileName](filename.md)

## <a name="remarks"></a>Примечания

Чтобы указать надстройки COM в качестве эквивалентной надстройки, укажите как элементы, так `ProgId` и `Type` элементы. Чтобы указать XLL в качестве эквивалентной надстройки, укажите как элементы, так `FileName` и `Type` элементы.

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Убедитесь, что надстройка Office совместима с существующей надстройкой COM](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)