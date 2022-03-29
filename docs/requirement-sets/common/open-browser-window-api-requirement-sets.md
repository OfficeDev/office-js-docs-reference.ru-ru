---
title: Открытие набора обязательных элементов окна браузера
description: Указывает, какие Office платформы и сборки поддерживают API openBrowserWindow.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 6e2e9cfd6eead7db1108a0fe434958e0d4a6513a
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487336"
---
# <a name="open-browser-window-api-requirement-sets"></a>Наборы требований к API окна открытых браузеров

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Набор API OpenBrowserWindow позволяет надстройке открывать браузер для выполнения задач, которые не всегда можно выполнить в песочнице управления веб-просмотром в самой надстройке; например, скачивание PDF-файла, когда управление веб-просмотром предоставляется Microsoft Edge.

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API API OpenBrowserWindow, Office хост-приложений, поддерживающий этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2021 или более поздней Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(обе подписки<br> и разовая покупка Office Mac 2019 и более поздних периодов)   | Office в Интернете  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | Сборка 16.0.14326.20454 или более поздней | Версия 1810 (сборка 16.0.11001.20074) или более поздней версии | 16.0.0.0 или более поздней | 16.0.0.0 или более поздней | Н/Д | Недоступно|

> [!NOTE]
> Набор требований OpenBrowserWindowApi доступен только следующим образом:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Версия и сборка номеров выпусков каналов обновления для Приложения Microsoft 365](/officeupdates/update-history-microsoft365-apps-by-date)
- [Какая у меня версия Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для Office клиентского приложения](/officeupdates/update-history-microsoft365-apps-by-date)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 — это первая версия API. Подробные сведения об API см. в [справочной Office.context.ui](/javascript/api/office/office.context#office-office-context-ui-member).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
