---
title: Наборы обязательных элементов cочетаний клавиш
description: Для надстройок Office набор Office клавиш.
ms.date: 02/15/2022
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 62160672ffc8517c5683f54248da9797db2205d0
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487452"
---
# <a name="keyboard-shortcuts-requirement-sets"></a>Наборы обязательных элементов cочетаний клавиш

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований к клавишам shortcuts, Office клиентские приложения, которые поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2013 или более поздней версии для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(обе подписки<br> и разовая покупка Office Mac 2019 и более поздних периодов)   | Office в Интернете  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | Недоступно | Версия: 2111 (сборка 14701.10000) | Недоступно | 16.55 | Сентябрь 2021 г. |

> [!NOTE]
> Набор **требований KeyboardShortcuts 1.1** поддерживается только в Excel.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="keyboardshortcuts-11"></a>KeyboardShortcuts 1.1

Сведения об API в этом наборе требований см. в [Office.actions](/javascript/api/office/office.actions).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
