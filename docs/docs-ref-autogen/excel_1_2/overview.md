---
title: Справочник по API JavaScript для Office
description: Требования API JavaScript для Office задаются ведущим приложением.
ms.date: 04/17/2020
ms.openlocfilehash: 765b2ee6108f6433ffe17d3ca15ba9c68fbd9617
ms.sourcegitcommit: 6dd770ff4893a67c625e1e4fd06ee197a3992ae0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/21/2020
ms.locfileid: "43595011"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="efbdf-103">Справочник по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="efbdf-103">Office JavaScript API reference</span></span>

<span data-ttu-id="efbdf-104">API JavaScript для Office позволяет создавать веб-приложения, взаимодействующие с объектными моделями в ведущих приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="efbdf-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="efbdf-105">В этом разделе приводятся дополнительные сведения о классах, методах и других типах, доступных для создания надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="efbdf-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="efbdf-106">Ниже приведен список наборов обязательных элементов для конкретных узлов (и общедоступных API с перекрестным узлом).</span><span class="sxs-lookup"><span data-stu-id="efbdf-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="efbdf-107">Каждый элемент связан с версией справочной документации по API, которая поддерживается в этом наборе требований (например, в ExcelApi 1,3 показаны API в ExcelApi 1,1, 1,2, 1,3, а также общий API).</span><span class="sxs-lookup"><span data-stu-id="efbdf-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="efbdf-108">`ExcelApiOnline 1.1`является особым набором требований.</span><span class="sxs-lookup"><span data-stu-id="efbdf-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="efbdf-109">Он содержит последние API для Excel в Интернете, но эти API могут быть не полностью поддерживаются на всех платформах.</span><span class="sxs-lookup"><span data-stu-id="efbdf-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="efbdf-110">Для получения дополнительных сведений обратитесь к разделу [API JavaScript для Excel Online](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) .</span><span class="sxs-lookup"><span data-stu-id="efbdf-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="efbdf-111">Выберите ссылку на этой странице, чтобы просмотреть справочную документацию по API, поддерживаемым указанным набором обязательных элементов, или с помощью раскрывающегося меню отбора фильтра над оглавлением, чтобы изменить набор требований в любое время.</span><span class="sxs-lookup"><span data-stu-id="efbdf-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="efbdf-112">Excel</span><span class="sxs-lookup"><span data-stu-id="efbdf-112">Excel</span></span>

- [<span data-ttu-id="efbdf-113">Предварительный просмотр ExcelApi</span><span class="sxs-lookup"><span data-stu-id="efbdf-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="efbdf-114">Ексцелапионлине 1,1</span><span class="sxs-lookup"><span data-stu-id="efbdf-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="efbdf-115">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="efbdf-115">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="efbdf-116">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="efbdf-116">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="efbdf-117">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="efbdf-117">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="efbdf-118">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="efbdf-118">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="efbdf-119">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="efbdf-119">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="efbdf-120">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="efbdf-120">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="efbdf-121">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="efbdf-121">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="efbdf-122">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="efbdf-122">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="efbdf-123">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="efbdf-123">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="efbdf-124">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="efbdf-124">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="efbdf-125">OneNote</span><span class="sxs-lookup"><span data-stu-id="efbdf-125">OneNote</span></span>

- [<span data-ttu-id="efbdf-126">OneNote 1,1</span><span class="sxs-lookup"><span data-stu-id="efbdf-126">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="efbdf-127">Outlook</span><span class="sxs-lookup"><span data-stu-id="efbdf-127">Outlook</span></span>

- [<span data-ttu-id="efbdf-128">Предварительный просмотр почтового ящика</span><span class="sxs-lookup"><span data-stu-id="efbdf-128">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="efbdf-129">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="efbdf-129">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="efbdf-130">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="efbdf-130">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="efbdf-131">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="efbdf-131">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="efbdf-132">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="efbdf-132">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="efbdf-133">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="efbdf-133">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="efbdf-134">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="efbdf-134">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="efbdf-135">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="efbdf-135">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="efbdf-136">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="efbdf-136">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="efbdf-137">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="efbdf-137">PowerPoint</span></span>

- [<span data-ttu-id="efbdf-138">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="efbdf-138">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="efbdf-139">Visio</span><span class="sxs-lookup"><span data-stu-id="efbdf-139">Visio</span></span>

- [<span data-ttu-id="efbdf-140">Висиоапи 1,1</span><span class="sxs-lookup"><span data-stu-id="efbdf-140">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="efbdf-141">Word</span><span class="sxs-lookup"><span data-stu-id="efbdf-141">Word</span></span>

- [<span data-ttu-id="efbdf-142">Предварительная версия Word</span><span class="sxs-lookup"><span data-stu-id="efbdf-142">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="efbdf-143">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="efbdf-143">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="efbdf-144">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="efbdf-144">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="efbdf-145">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="efbdf-145">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="efbdf-146">Общий API</span><span class="sxs-lookup"><span data-stu-id="efbdf-146">Common API</span></span>

- [<span data-ttu-id="efbdf-147">Общий API</span><span class="sxs-lookup"><span data-stu-id="efbdf-147">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="efbdf-148">См. также</span><span class="sxs-lookup"><span data-stu-id="efbdf-148">See also</span></span>

- [<span data-ttu-id="efbdf-149">Сведения о надстройках Office</span><span class="sxs-lookup"><span data-stu-id="efbdf-149">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="efbdf-150">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="efbdf-150">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="efbdf-151">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="efbdf-151">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="efbdf-152">Изучение API JavaScript для Office с помощью Script Lab</span><span class="sxs-lookup"><span data-stu-id="efbdf-152">Explore Office JavaScript API using Script Lab</span></span>](/office/dev/add-ins/overview/explore-with-script-lab)
