# <a name="requestedheight-element"></a><span data-ttu-id="fa64d-101">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="fa64d-101">RequestedHeight element</span></span>

<span data-ttu-id="fa64d-102">Задает начальное высоту (в точках) контента надстройки или надстройки почты.</span><span class="sxs-lookup"><span data-stu-id="fa64d-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="fa64d-103">**Типа надстройки:** Контент, почты</span><span class="sxs-lookup"><span data-stu-id="fa64d-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fa64d-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="fa64d-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="fa64d-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="fa64d-105">Contained in</span></span>

- <span data-ttu-id="fa64d-106">[DefaultSettings](defaultsettings.md) (Содержимого надстроек) со значением, которое может быть в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="fa64d-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="fa64d-107">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (надстройки почты) со значением, которое может быть в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="fa64d-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="fa64d-108">[ExtensionPoint](extensionpoint.md) (Надстройки контекстной почты) со значением, которое может быть от 140 до 450 пикселей для точки расширения **DetectedEntity** и от 32 до 450 точки расширения **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="fa64d-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>