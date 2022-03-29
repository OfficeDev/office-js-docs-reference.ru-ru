---
title: Excel API JavaScript установлено 1.11
description: Сведения о наборе требований ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 284480174dc2f634576476fc4e2493026819ff20
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487631"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Новые возможности в Excel API JavaScript 1.11

В ExcelApi 1.11 улучшена поддержка комментариев и элементов управления на уровне книг (например, сохранение и закрытие книги). Кроме того, добавлен доступ к настройкам культуры для учета локализации.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Комментарии](/office/dev/add-ins/excel/excel-add-ins-comments.md#mentions) |Теги и извества других пользователей книг с помощью комментариев. | [Комментарий](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Разрешение [комментариев](/office/dev/add-ins/excel/excel-add-ins-comments.md#resolve-comment-threads) | Разрешить потоки комментариев и получить состояние разрешения. | [Comment](/javascript/api/excel/excel.comment) |
| [Параметры культуры](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Получает параметры культурной системы для книги, например форматирование номеров. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Вырезать и вклеить (moveTo)](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | Реплицирует функции cut-and-paste в Excel для диапазона. | [Range](/javascript/api/excel/excel.range) |
| [Сохранение](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#save-the-workbook) и [закрытие](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#close-the-workbook) рабочей книги | Сохранение и закрытие книг. | [Workbook](/javascript/api/excel/excel.workbook) |
| События таблицы | Дополнительные события и сведения о событиях для вычислений таблиц и скрытых строк. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.11. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.11 или ранее, см. в Excel API в наборе требований [1.11 или ранее](/javascript/api/excel?view=excel-js-1.11&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_11.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
