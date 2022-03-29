---
title: Excel API JavaScript установлено 1.12
description: Сведения о наборе требований ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 01cf9523f60fbf374108a81de841840310a4a875
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487125"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Новые возможности в Excel API JavaScript 1.12

ExcelApi 1.12 увеличил поддержку формул в диапазонах, добавив API для отслеживания динамических массивов и поиска прямых прецедентов формулы. Кроме того, добавлен контроль API фильтров PivotTable. Улучшения также были сделаны в областях комментариев, параметров культуры и пользовательских свойств.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События комментариев](/office/dev/add-ins/excel/excel-add-ins-comments.md#comment-events) | Добавляет события для добавления, изменения и удаления в коллекцию комментариев.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Параметры культуры [даты и времени](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Предоставляет доступ к дополнительным культурным настройкам даты и времени форматирования. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Прямые прецеденты](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents) | Возвращает диапазоны, используемые для оценки формулы ячейки.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Фильтры поворота | Применяет фильтры, управляемые значением, к полям PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Разлиение диапазона](/office/dev/add-ins/excel/excel-add-ins-ranges-dynamic-arrays) | Позволяет надстройки находить диапазоны, связанные с [динамическими результатами массива](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Настраиваемые свойства на уровне таблицы](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Позволяет настраивать свойства на уровне таблицы, в дополнение к области на уровне книг. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.12. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.12 или ранее, см. в Excel API в наборе требований [1.12 или ранее](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_12.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
