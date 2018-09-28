# <a name="versionoverrides-element"></a><span data-ttu-id="8ce50-101">Элемент VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8ce50-101">VersionOverrides element</span></span>

<span data-ttu-id="8ce50-p101">Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="8ce50-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8ce50-105">Attributes</span></span>

|  <span data-ttu-id="8ce50-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="8ce50-106">Attribute</span></span>  |  <span data-ttu-id="8ce50-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ce50-107">Required</span></span>  |  <span data-ttu-id="8ce50-108">Описание</span><span class="sxs-lookup"><span data-stu-id="8ce50-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8ce50-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="8ce50-109">**xmlns**</span></span>       |  <span data-ttu-id="8ce50-110">Да</span><span class="sxs-lookup"><span data-stu-id="8ce50-110">Yes</span></span>  |  <span data-ttu-id="8ce50-111">Расположение схемы (`http://schemas.microsoft.com/office/mailappversionoverrides`, когда `xsi:type` — `VersionOverridesV1_0`, и `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, когда `xsi:type` — `VersionOverridesV1_1`).</span><span class="sxs-lookup"><span data-stu-id="8ce50-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="8ce50-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="8ce50-112">**xsi:type**</span></span>  |  <span data-ttu-id="8ce50-113">Да</span><span class="sxs-lookup"><span data-stu-id="8ce50-113">Yes</span></span>  | <span data-ttu-id="8ce50-p102">Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="8ce50-116">В настоящее время только Outlook 2016 поддерживает схемы версии 1.1 VersionOverrides и `VersionOverridesV1_1` типа.</span><span class="sxs-lookup"><span data-stu-id="8ce50-116">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="8ce50-117">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="8ce50-117">Child elements</span></span>

|  <span data-ttu-id="8ce50-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="8ce50-118">Element</span></span> |  <span data-ttu-id="8ce50-119">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ce50-119">Required</span></span>  |  <span data-ttu-id="8ce50-120">Описание</span><span class="sxs-lookup"><span data-stu-id="8ce50-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="8ce50-121">**Описание**</span><span class="sxs-lookup"><span data-stu-id="8ce50-121">**Description**</span></span>    |  <span data-ttu-id="8ce50-122">НЕТ</span><span class="sxs-lookup"><span data-stu-id="8ce50-122">No</span></span>   |  <span data-ttu-id="8ce50-p103">Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="8ce50-127">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="8ce50-127">**Requirements**</span></span>  |  <span data-ttu-id="8ce50-128">Нет</span><span class="sxs-lookup"><span data-stu-id="8ce50-128">No</span></span>   |  <span data-ttu-id="8ce50-p104">Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="8ce50-131">Hosts</span><span class="sxs-lookup"><span data-stu-id="8ce50-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="8ce50-132">Да</span><span class="sxs-lookup"><span data-stu-id="8ce50-132">Yes</span></span>  |  <span data-ttu-id="8ce50-p105">Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="8ce50-135">Resources</span><span class="sxs-lookup"><span data-stu-id="8ce50-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="8ce50-136">Да</span><span class="sxs-lookup"><span data-stu-id="8ce50-136">Yes</span></span>  | <span data-ttu-id="8ce50-137">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.</span><span class="sxs-lookup"><span data-stu-id="8ce50-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="8ce50-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="8ce50-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="8ce50-139">Нет</span><span class="sxs-lookup"><span data-stu-id="8ce50-139">No</span></span>  | <span data-ttu-id="8ce50-p106">Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="8ce50-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="8ce50-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="8ce50-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="8ce50-143">Нет</span><span class="sxs-lookup"><span data-stu-id="8ce50-143">No</span></span>  | <span data-ttu-id="8ce50-144">Указывает сведения о связанном с надстройкой веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="8ce50-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="8ce50-145">Пример VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8ce50-145">VersionOverrides example</span></span>
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="8ce50-146">Реализация нескольких версий</span><span class="sxs-lookup"><span data-stu-id="8ce50-146">Implementing multiple versions</span></span>

<span data-ttu-id="8ce50-p107">В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.</span><span class="sxs-lookup"><span data-stu-id="8ce50-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="8ce50-149">Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии.</span><span class="sxs-lookup"><span data-stu-id="8ce50-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="8ce50-150">Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="8ce50-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="8ce50-151">Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="8ce50-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
