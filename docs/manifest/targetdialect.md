---
title: Элемент TargetDialect в файле манифеста
description: Элемент TargetDialect определяет региональный язык, поддерживаемый этим словарем, представленный в качестве строки имени культуры.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: dfa8b63a6916da45aecbb64398e841a9e04671aa
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486864"
---
# <a name="targetdialect-element"></a>Элемент TargetDialect

Задает поддерживаемый этим словарем региональный язык в виде строки с названием языка и региональных параметров.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<TargetDialect>
   string 
</TargetDialect>
```

## <a name="contained-in"></a>Содержится в

[TargetDialects](targetdialects.md)

## <a name="remarks"></a>Замечания

Укажите значение в формате языковых тегов BCP 47, например `en-US`.

## <a name="see-also"></a>См. также

- [Создание надстройки области задач словаря](/office/dev/add-ins/word/dictionary-task-pane-add-ins)
