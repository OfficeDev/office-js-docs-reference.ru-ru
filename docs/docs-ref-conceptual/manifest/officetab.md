# <a name="officetab-element"></a><span data-ttu-id="1138b-101">Элемент OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1138b-101">OfficeTab element</span></span>

<span data-ttu-id="1138b-p101">Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.</span><span class="sxs-lookup"><span data-stu-id="1138b-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1138b-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1138b-105">Child elements</span></span>

|  <span data-ttu-id="1138b-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="1138b-106">Element</span></span> |  <span data-ttu-id="1138b-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1138b-107">Required</span></span>  |  <span data-ttu-id="1138b-108">Описание</span><span class="sxs-lookup"><span data-stu-id="1138b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1138b-109">Group</span><span class="sxs-lookup"><span data-stu-id="1138b-109">Group</span></span>      | <span data-ttu-id="1138b-110">Да</span><span class="sxs-lookup"><span data-stu-id="1138b-110">Yes</span></span> |  <span data-ttu-id="1138b-p102">Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.</span><span class="sxs-lookup"><span data-stu-id="1138b-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="1138b-113">Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="1138b-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="1138b-114">Значения **полужирным** шрифтом, поддерживаются в настольных компьютеров и сети (например, Word 2016 или более поздней версии для Windows и Word Online).</span><span class="sxs-lookup"><span data-stu-id="1138b-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="1138b-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="1138b-115">Outlook</span></span>

- <span data-ttu-id="1138b-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="1138b-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="1138b-117">Word</span><span class="sxs-lookup"><span data-stu-id="1138b-117">Word</span></span>

- <span data-ttu-id="1138b-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1138b-118">**TabHome**</span></span>
- <span data-ttu-id="1138b-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1138b-119">**TabInsert**</span></span>
- <span data-ttu-id="1138b-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="1138b-120">TabWordDesign</span></span>
- <span data-ttu-id="1138b-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="1138b-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="1138b-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="1138b-122">TabReferences</span></span>
- <span data-ttu-id="1138b-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="1138b-123">TabMailings</span></span>
- <span data-ttu-id="1138b-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="1138b-124">TabReviewWord</span></span>
- <span data-ttu-id="1138b-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1138b-125">**TabView**</span></span>
- <span data-ttu-id="1138b-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1138b-126">TabDeveloper</span></span>
- <span data-ttu-id="1138b-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1138b-127">TabAddIns</span></span>
- <span data-ttu-id="1138b-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="1138b-128">TabBlogPost</span></span>
- <span data-ttu-id="1138b-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="1138b-129">TabBlogInsert</span></span>
- <span data-ttu-id="1138b-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1138b-130">TabPrintPreview</span></span>
- <span data-ttu-id="1138b-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="1138b-131">TabOutlining</span></span>
- <span data-ttu-id="1138b-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="1138b-132">TabConflicts</span></span>
- <span data-ttu-id="1138b-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1138b-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1138b-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1138b-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="1138b-135">Excel</span><span class="sxs-lookup"><span data-stu-id="1138b-135">Excel</span></span>

- <span data-ttu-id="1138b-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1138b-136">**TabHome**</span></span>
- <span data-ttu-id="1138b-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1138b-137">**TabInsert**</span></span>
- <span data-ttu-id="1138b-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="1138b-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="1138b-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="1138b-139">TabFormulas</span></span>
- <span data-ttu-id="1138b-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="1138b-140">**TabData**</span></span>
- <span data-ttu-id="1138b-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="1138b-141">**TabReview**</span></span>
- <span data-ttu-id="1138b-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1138b-142">**TabView**</span></span>
- <span data-ttu-id="1138b-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1138b-143">TabDeveloper</span></span>
- <span data-ttu-id="1138b-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1138b-144">TabAddIns</span></span>
- <span data-ttu-id="1138b-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1138b-145">TabPrintPreview</span></span>
- <span data-ttu-id="1138b-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1138b-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="1138b-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1138b-147">PowerPoint</span></span>

- <span data-ttu-id="1138b-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1138b-148">**TabHome**</span></span>
- <span data-ttu-id="1138b-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1138b-149">**TabInsert**</span></span>
- <span data-ttu-id="1138b-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="1138b-150">**TabDesign**</span></span>
- <span data-ttu-id="1138b-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="1138b-151">**TabTransitions**</span></span>
- <span data-ttu-id="1138b-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="1138b-152">**TabAnimations**</span></span>
- <span data-ttu-id="1138b-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="1138b-153">TabSlideShow</span></span>
- <span data-ttu-id="1138b-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="1138b-154">TabReview</span></span>
- <span data-ttu-id="1138b-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1138b-155">**TabView**</span></span>
- <span data-ttu-id="1138b-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1138b-156">TabDeveloper</span></span>
- <span data-ttu-id="1138b-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1138b-157">TabAddIns</span></span>
- <span data-ttu-id="1138b-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1138b-158">TabPrintPreview</span></span>
- <span data-ttu-id="1138b-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="1138b-159">TabMerge</span></span>
- <span data-ttu-id="1138b-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="1138b-160">TabGrayscale</span></span>
- <span data-ttu-id="1138b-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="1138b-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="1138b-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1138b-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="1138b-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="1138b-163">TabSlideMaster</span></span>
- <span data-ttu-id="1138b-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="1138b-164">TabHandoutMaster</span></span>
- <span data-ttu-id="1138b-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="1138b-165">TabNotesMaster</span></span>
- <span data-ttu-id="1138b-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1138b-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1138b-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="1138b-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="1138b-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="1138b-168">OneNote</span></span>

- <span data-ttu-id="1138b-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1138b-169">**TabHome**</span></span>
- <span data-ttu-id="1138b-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1138b-170">**TabInsert**</span></span>
- <span data-ttu-id="1138b-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1138b-171">**TabView**</span></span>
- <span data-ttu-id="1138b-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1138b-172">TabDeveloper</span></span>
- <span data-ttu-id="1138b-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1138b-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="1138b-174">Group</span><span class="sxs-lookup"><span data-stu-id="1138b-174">Group</span></span>

<span data-ttu-id="1138b-p104">Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью об[элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="1138b-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="1138b-179">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1138b-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
