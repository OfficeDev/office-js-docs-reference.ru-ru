---
title: Наборы обязательных элементов для приведения изображений
description: Поддержка наборов требований к принуждению к изображениям с Office надстройки в Excel, PowerPoint и Word.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: c6b04847f1a46e0205cf9699f37eeec803252021
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487129"
---
# <a name="image-coercion-requirement-sets"></a>Наборы обязательных элементов для приведения изображений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 позволяет преобразования в изображение (`Office.CoercionType.Image`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) метода. Поддерживаются следующие приложения.

- Excel 2013 и более поздней Windows
- Excel 2016 и более поздней основе на Mac
- Excel на iPad
- OneNote в Интернете
- PowerPoint 2013 и более поздней Windows
- PowerPoint 2016 и позднее на Mac
- PowerPoint в Интернете
- PowerPoint на iPad
- Word 2013 и более поздней версии для Windows
- Word 2016 и более поздней версии для Mac
- Word в Интернете
- Word для iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 позволяет преобразования в формат SVG (`Office.CoercionType.XmlSvg`) при записи данных с помощью [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) метода. Поддерживаются следующие приложения.

- Excel 2021 г. и более поздней Windows
- Excel 2021 г. и позднее на Mac
- PowerPoint 2021 г. и более поздней Windows
- PowerPoint 2021 г. и более поздней
- PowerPoint в Интернете
- Word 2021 и более поздний Windows
- Word 2021 и более поздний на Mac

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
