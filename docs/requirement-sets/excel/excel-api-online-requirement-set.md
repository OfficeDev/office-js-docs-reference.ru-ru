---
title: Excel API JavaScript только для интернета
description: Сведения о наборе требований ExcelApiOnline.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ade48cec3e1db452206a1d35291b3f6f3e609566
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486937"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel API JavaScript только для интернета

Набор `ExcelApiOnline` требований — это специальный набор требований, который включает функции, доступные только для Excel в Интернете. API в этом наборе требований считаются производственными API (не подверженными незадокументированные поведенческие или структурные изменения) для Excel в Интернете приложения. `ExcelApiOnline`API считаются "предварительными" API для других платформ (Windows, Mac, iOS) и не могут поддерживаться ни одной из этих платформ.

Когда API в наборе `ExcelApiOnline` требований поддерживаются на всех платформах, они будут добавлены в следующий выпущенный набор требований (`ExcelApi 1.[NEXT]`). После того как это новое требование станет общедоступным, эти API будут удалены из `ExcelApiOnline`. Думайте об этом как об аналогичном процессе продвижения по службе aPI, перемещаемом с предварительного просмотра на выпуск.

> [!IMPORTANT]
> `ExcelApiOnline` — это суперсет последнего набора требований с номерами.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` это единственная версия API только для интернета. Это потому, Excel в Интернете всегда будет иметь одну версию, доступную пользователям, которая является последней версией.

В следующей таблице приводится краткий сводка API, а в следующей таблице списка [API](#api-list) приводится подробный список текущих `ExcelApiOnline` API.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Связанные книги | Управление связями между книгами, включая поддержку обновления и разрыва ссылок на книги. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Именуемые представления листа | Предоставляет программный контроль представлений таблицы для каждого пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| События перемещения таблиц | Обнаружение перемещений таблиц в пределах коллекции, положения таблицы и источника изменения. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>Рекомендуемое использование

Так `ExcelApiOnline` как API поддерживаются только Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований перед вызовом этих API. Это позволяет избежать вызова API только для интернета на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к платформе, необходимо удалить или изменить проверку `isSetSupported` . Это позволит включить функцию надстройки на других платформах. При внесении этого изменения обязательно проверьте эту функцию на этих платформах.

> [!IMPORTANT]
> Манифест не может указываться `ExcelApiOnline 1.1` как требование активации. Это значение не является допустимым для использования в [элементе Set](/javascript/api/manifest/set).

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, включенные в набор `ExcelApiOnline` требований. Полный список всех API Excel JavaScript (`ExcelApiOnline`включая API и ранее выпущенные API), см. Excel [API JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-online.md)]
## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
