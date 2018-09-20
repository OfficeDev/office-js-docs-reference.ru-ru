# <a name="host-element"></a><span data-ttu-id="f4995-101">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="f4995-101">Host element</span></span>

<span data-ttu-id="f4995-102">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="f4995-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="f4995-103">Синтаксис элемента **узла** , может изменяться в зависимости от того, является ли элемента, определенного в [базовой манифест](#basic-manifest) или в узел [VersionOverrides](#versionoverrides-node) .</span><span class="sxs-lookup"><span data-stu-id="f4995-103">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="f4995-104">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="f4995-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="f4995-105">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="f4995-105">Basic manifest</span></span>

<span data-ttu-id="f4995-106">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="f4995-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="f4995-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f4995-107">Attributes</span></span>

| <span data-ttu-id="f4995-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f4995-108">Attribute</span></span>     | <span data-ttu-id="f4995-109">Тип</span><span class="sxs-lookup"><span data-stu-id="f4995-109">Type</span></span>   | <span data-ttu-id="f4995-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f4995-110">Required</span></span> | <span data-ttu-id="f4995-111">Описание</span><span class="sxs-lookup"><span data-stu-id="f4995-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="f4995-112">Name</span><span class="sxs-lookup"><span data-stu-id="f4995-112">Name</span></span>](#name) | <span data-ttu-id="f4995-113">string</span><span class="sxs-lookup"><span data-stu-id="f4995-113">string</span></span> | <span data-ttu-id="f4995-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f4995-114">required</span></span> | <span data-ttu-id="f4995-115">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="f4995-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="f4995-116">Имя</span><span class="sxs-lookup"><span data-stu-id="f4995-116">Name</span></span>
<span data-ttu-id="f4995-p102">Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="f4995-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="f4995-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="f4995-119">`Document` (Word)</span></span>
- <span data-ttu-id="f4995-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="f4995-120">`Database` (Access)</span></span>
- <span data-ttu-id="f4995-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="f4995-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="f4995-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="f4995-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="f4995-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="f4995-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="f4995-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="f4995-124">`Project` (Project)</span></span>
- <span data-ttu-id="f4995-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="f4995-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="f4995-126">Пример</span><span class="sxs-lookup"><span data-stu-id="f4995-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="f4995-127">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="f4995-127">VersionOverrides node</span></span>
<span data-ttu-id="f4995-128">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="f4995-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="f4995-129">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f4995-129">Attributes</span></span>

|  <span data-ttu-id="f4995-130">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f4995-130">Attribute</span></span>  |  <span data-ttu-id="f4995-131">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f4995-131">Required</span></span>  |  <span data-ttu-id="f4995-132">Описание</span><span class="sxs-lookup"><span data-stu-id="f4995-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f4995-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f4995-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="f4995-134">Да</span><span class="sxs-lookup"><span data-stu-id="f4995-134">Yes</span></span>  | <span data-ttu-id="f4995-135">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="f4995-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="f4995-136">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f4995-136">Child elements</span></span>

|  <span data-ttu-id="f4995-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="f4995-137">Element</span></span> |  <span data-ttu-id="f4995-138">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f4995-138">Required</span></span>  |  <span data-ttu-id="f4995-139">Описание</span><span class="sxs-lookup"><span data-stu-id="f4995-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f4995-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f4995-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="f4995-141">Да</span><span class="sxs-lookup"><span data-stu-id="f4995-141">Yes</span></span>   |  <span data-ttu-id="f4995-142">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="f4995-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="f4995-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="f4995-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="f4995-144">Нет</span><span class="sxs-lookup"><span data-stu-id="f4995-144">No</span></span>   |  <span data-ttu-id="f4995-p103">Определяет параметры форм-фактора мобильного устройства. **Примечание.** Этот элемент поддерживается только в Outlook для iOS.</span><span class="sxs-lookup"><span data-stu-id="f4995-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="f4995-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="f4995-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="f4995-148">Нет</span><span class="sxs-lookup"><span data-stu-id="f4995-148">No</span></span>   |  <span data-ttu-id="f4995-149">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="f4995-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="f4995-150">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="f4995-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="f4995-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f4995-151">xsi:type</span></span>

<span data-ttu-id="f4995-152">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="f4995-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="f4995-153">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="f4995-153">The value must be one of the following:</span></span>

- <span data-ttu-id="f4995-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="f4995-154">`Document` (Word)</span></span>
- <span data-ttu-id="f4995-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="f4995-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="f4995-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="f4995-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="f4995-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="f4995-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="f4995-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="f4995-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="f4995-159">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="f4995-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
