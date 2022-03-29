---
title: Excel API JavaScript установлено 1.14
description: Сведения о наборе требований ExcelApi 1.14.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 44f3b56cf0bf9d5366cdfe7c2ae0454b03c2d3d6
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487625"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Новые возможности в Excel API JavaScript 1.14

В ExcelApi 1.14 добавлены объекты для управления функцией таблицы данных диаграммы, методом обнаружения всех ячеек-прецедентов формулы и событиями защиты листа для отслеживания изменений состояния защиты листа. Он также добавил несколько методов [`getItemOrNullObject`](/office/dev/add-ins/develop/application-specific-api-model.md#ornullobject-methods-and-properties) для таких объектов`CommentCollection`, как , и `ShapeCollection``StyleCollection` для улучшения обработки ошибок.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Таблицы данных диаграммы](/office/dev/add-ins/excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Управление внешним видом, форматированием и видимостью таблиц данных на диаграммах. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Прецеденты формул](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Возвращаем все ячейки-прецеденты формулы. | [Range](/javascript/api/excel/excel.range) |
| Запросы | Извлечение атрибутов Power Query, таких как имя, дата обновления и количество запросов. | [Запрос](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [События защиты таблиц](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Отслеживание изменений состояния защиты таблицы и источника этих изменений. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [WorksheetCollection](/javascript/api/excel/excel.worksheet)[](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.14. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.14 или ранее, см. в Excel API в наборе требований [1.14 или ранее](/javascript/api/excel?view=excel-js-1.14&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_14.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
