---
layout: LandingPage
ms.topic: landing-page
title: Справочник по API JavaScript для Office
description: Интерфейсы JavaScript для Office по ведущему приложению и версии.
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 06/17/2020
ms.openlocfilehash: f3591e0707f20a448f20eb6a444c4c655612f966
ms.sourcegitcommit: 538c15a77b09cf4bf87911e81991d784aeae4ab0
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47824533"
---
# <a name="office-add-ins-javascript-api-reference"></a>Справочник по API JavaScript для надстроек Office

API JavaScript для Office позволяет создавать веб-приложения, взаимодействующие с объектными моделями в ведущих приложениях Office. В этом разделе представлены дополнительные сведения о классах, методах и других типах, доступных для создания надстроек Office.

Ниже приведен список интерфейсов API для [поддерживаемых ведущих приложений Office](/office/dev/add-ins/overview/office-add-in-availability). По общим ссылкам представлены все API, не относящиеся к определенному ведущему приложению (как описывается в [наборе требований для общих API Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)). Остальные элементы ссылаются на версию справочной документации по API, связанную с соответствующим ведущим приложением, по набору требований. Версии справочной документации составляются таким образом, что они включали все API до соответствующего набора требований включительно (например, в ExcelApi 1.3 представлены API из ExcelApi 1.1, 1.2, 1.3, а также общие API).

`ExcelApiOnline 1.1` — это особый набор требований. Он содержит новейшие API для Excel в Интернете, но эти API могут еще не полностью поддерживаться на всех платформах. Дополнительные сведения см. в [наборе требований API JavaScript для Excel (только в Интернете)](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set).

> [!TIP]
> Вы можете в любой момент изменить версию справочной страницы с помощью раскрывающегося меню выбора фильтра над оглавлением. Если такой версии страницы не существует, снова откроется текущая версия.

<h2>Ведущие приложения Office</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Excel</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">ExcelApi Preview</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.12">ExcelApi 1.12</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">Общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Outlook</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Mailbox Preview</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Mailbox 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Mailbox 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Mailbox 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Mailbox 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Mailbox 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Mailbox 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Mailbox 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Mailbox 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">Общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Word</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">WordApi Preview</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">Общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API OneNote</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">Общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Visio</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API PowerPoint</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-1.1">Общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Project</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">Только общие API</a></li>
                </ul>
            </div>
        </a>
    </li>
</ul>

> [!NOTE]
> Если вас интересуют API JavaScript для разработки сценариев Office Scripts, ознакомьтесь со [справочником по API для сценариев Office](/javascript/api/office-scripts/overview).

## <a name="see-also"></a>См. также

- [Сведения о надстройках Office](/office/dev/add-ins/overview)
- [Доступность ведущих приложений и платформ для надстроек Office](/office/dev/add-ins/overview/office-add-in-availability)
- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Изучение API JavaScript для Office с помощью Script Lab](/office/dev/add-ins/overview/explore-with-script-lab)
