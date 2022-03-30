---
title: Excel API JavaScript установлено 1.12
description: Сведения о наборе требований ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 002050f1b1c0b4420229146614182b53976f0f4c
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502359"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Новые возможности в Excel API JavaScript 1.12

ExcelApi 1.12 увеличил поддержку формул в диапазонах, добавив API для отслеживания динамических массивов и поиска прямых прецедентов формулы. Кроме того, добавлен контроль API фильтров PivotTable. Улучшения также были сделаны в областях комментариев, параметров культуры и пользовательских свойств.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События комментариев](/office/dev/add-ins/excel/excel-add-ins-comments#comment-events) | Добавляет события для добавления, изменения и удаления в коллекцию комментариев.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Параметры культуры [даты и времени](/office/dev/add-ins/excel/excel-add-ins-workbooks#access-application-culture-settings) | Предоставляет доступ к дополнительным культурным настройкам даты и времени форматирования. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Прямые прецеденты](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents) | Возвращает диапазоны, используемые для оценки формулы ячейки.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Фильтры поворота | Применяет фильтры, управляемые значением, к полям PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Разлиение диапазона](/office/dev/add-ins/excel/excel-add-ins-ranges-dynamic-arrays) | Позволяет надстройки находить диапазоны, связанные с [динамическими результатами массива](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Настраиваемые свойства на уровне таблицы](/office/dev/add-ins/excel/excel-add-ins-workbooks#worksheet-level-custom-properties) | Позволяет настраивать свойства на уровне таблицы, в дополнение к области на уровне книг. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.12. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.12 или ранее, см. в Excel API в наборе требований [1.12 или ранее](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_12.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
