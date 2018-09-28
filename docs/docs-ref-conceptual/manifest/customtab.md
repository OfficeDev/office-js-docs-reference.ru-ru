# <a name="customtab-element"></a><span data-ttu-id="06952-101">Элемент CustomTab</span><span class="sxs-lookup"><span data-stu-id="06952-101">CustomTab element</span></span>

<span data-ttu-id="06952-p101">На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка.</span><span class="sxs-lookup"><span data-stu-id="06952-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="06952-p102">На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.</span><span class="sxs-lookup"><span data-stu-id="06952-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="06952-107">Атрибут **id** должен быть уникальным для манифеста.</span><span class="sxs-lookup"><span data-stu-id="06952-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="06952-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="06952-108">Child elements</span></span>

|  <span data-ttu-id="06952-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="06952-109">Element</span></span> |  <span data-ttu-id="06952-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="06952-110">Required</span></span>  |  <span data-ttu-id="06952-111">Описание</span><span class="sxs-lookup"><span data-stu-id="06952-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="06952-112">Group</span><span class="sxs-lookup"><span data-stu-id="06952-112">Group</span></span>](group.md)      | <span data-ttu-id="06952-113">Да</span><span class="sxs-lookup"><span data-stu-id="06952-113">Yes</span></span> |  <span data-ttu-id="06952-114">Определяет группу команд.</span><span class="sxs-lookup"><span data-stu-id="06952-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="06952-115">Подпись</span><span class="sxs-lookup"><span data-stu-id="06952-115">Label</span></span>](#label-tab)      | <span data-ttu-id="06952-116">Да</span><span class="sxs-lookup"><span data-stu-id="06952-116">Yes</span></span> |  <span data-ttu-id="06952-117">Метка элемента CustomTab или Group.</span><span class="sxs-lookup"><span data-stu-id="06952-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="06952-118">Control</span><span class="sxs-lookup"><span data-stu-id="06952-118">Control</span></span>](control.md)    | <span data-ttu-id="06952-119">Да</span><span class="sxs-lookup"><span data-stu-id="06952-119">Yes</span></span> |  <span data-ttu-id="06952-120">Коллекция из одного или нескольких объектов Control.</span><span class="sxs-lookup"><span data-stu-id="06952-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="06952-121">Group</span><span class="sxs-lookup"><span data-stu-id="06952-121">Group</span></span>

<span data-ttu-id="06952-p103">Обязательный. См. статью об [элементе Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="06952-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="06952-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="06952-124">Label (Tab)</span></span>

<span data-ttu-id="06952-p104">Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="06952-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="06952-127">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="06952-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```