---
title: Excel API JavaScript установлено 1.4
description: Сведения о наборе требований ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 555e75fe4d7a210e22a74081b29869a2dc17b0cb
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503472"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Новые возможности API JavaScript для Excel 1.4

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.4, относящийся к API JavaScript для Excel.

## <a name="named-item-add-and-new-properties"></a>Именованный элемент add и новые свойства

Новые свойства:

* `comment`
* `scope` — таблицы или элементы с объемной областью книги.
* `worksheet` - Возвращает таблицу, в которой область действия именуемого элемента.

Новые методы:

* `add(name: string, reference: Range or string, comment: string)` - Добавляет новое имя в коллекцию данной области.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Добавляет новое имя в коллекцию данной области с помощью пользовательского локаула для формулы.

## <a name="settings-api-in-the-excel-namespace"></a>Параметры API в пространстве имен Excel

Объект [Setting](/javascript/api/excel/excel.setting) представляет пару "ключ-значение" для параметра, хранящегося в документе. Функциональные возможности объекта `Excel.Setting` аналогичны `Office.Settings`, но он использует пакетный синтаксис API, а не модель обратного вызова общего API.

API включают в `getItem()` себя создание записи с помощью ключа `add()` и добавление указанной пары параметров ключа:значения в книгу.

## <a name="others"></a>Другие

* Установите имя столбца таблицы.
* Добавьте столбец таблицы в конец таблицы.
* Добавьте несколько строк в таблицу одновременно.
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* Методы [\*и свойства OrNullObject](/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). Эта функция позволяет получать объект с помощью ключа. Если объекта не существует, `isNullObject` свойство возвращенного объекта будет верным. Это позволяет разработчикам проверять, существует ли объект без обработки исключений. Метод `*OrNullObject` доступен на большинстве объектов коллекции.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.4. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.4 или ранее, см. Excel API в наборе требований [1.4 или ранее](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_4.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
