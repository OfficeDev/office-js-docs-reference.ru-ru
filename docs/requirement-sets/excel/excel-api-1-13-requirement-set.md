---
title: Excel API JavaScript установлено 1.13
description: Сведения о наборе требований ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e6cb0c7e7b47b9f6f0d927a13490ea8f84588519
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502345"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Новые возможности в Excel API JavaScript 1.13

В ExcelApi 1.13 добавлен метод вставки листа в книгу из строки с кодированной базой 64 и событие для обнаружения активации книги. Это также увеличило поддержку формул в диапазонах, добавив API для отслеживания изменений формул и поиска прямых зависимых ячеек формулы. Кроме того, она расширила поддержку PivotTable, добавив API PivotLayout для ALT-текста, стиля и управления пустыми ячейками.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События с измененной формулой](/office/dev/add-ins/excel/excel-add-ins-worksheets#detect-formula-changes) | Отслеживание изменений формул, в том числе источника и типа события, которое вызвало изменение. | [Таблица.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Зависимые формулы](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents#get-the-direct-dependents-of-a-formula) | Найдите прямые зависимые ячейки формулы. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Вставка таблиц](/office/dev/add-ins/excel/excel-add-ins-workbooks#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Вставьте таблицы из другой книги в текущую книгу в качестве строки с кодом Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotTable PivotLayout](/office/dev/add-ins/excel/excel-add-ins-pivottables#other-pivotlayout-functions) | Расширение класса PivotLayout, включая новую поддержку текста alt и управление пустыми ячейками. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, установленный 1.13. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.13 или ранее, см. в Excel API в наборе требований [1.13 или ранее](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_13.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
