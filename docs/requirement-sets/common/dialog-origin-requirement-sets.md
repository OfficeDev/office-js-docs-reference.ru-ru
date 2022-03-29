---
title: Наборы обязательных элементов источников диалоговых окон
description: Дополнительные информацию о наборах требований к диалоговом происхождению.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: a9328117ec06aa3357866191dcd8109a6a76deeb
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486953"
---
# <a name="dialog-origin-requirement-sets"></a>Наборы обязательных элементов источников диалоговых окон

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований к диалоговому источнику, Office клиентские приложения, поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 или более поздней версии для Windows<br>(единовременная покупка) | Office для Windows<br>(подписка) |  Office для iPad<br>(подписка)  |  Office для Mac<br>(обе подписки<br> и разовая покупка Office Mac 2019 и более поздних периодов)  | Office в Интернете  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Сборка<br>15.0.5371.1000<br>или более поздней | Сборка<br>16.0.5200.1000<br>или более поздней | Сборка<br>Подлежит уточнению.<br>или более поздней | Подлежит уточнению. | 2.52 или более поздней | 16.52 или более поздней версии | Июль 2021 г. | Версия 2108<br>(Сборка 10377.1000)<br>или более поздней |

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-origin-11"></a>Диалоговое начало 1.1

Диалоговое начало 1.1 — это первая версия API. Он обеспечивает поддержку меж доменного обмена сообщениями между диалогом и родительской страницей. Сведения об этих API см. в справочной [Office.ui](/javascript/api/office/office.ui).

## <a name="see-also"></a>См. также

- [Использование Office Dialog API в надстройках Office](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
