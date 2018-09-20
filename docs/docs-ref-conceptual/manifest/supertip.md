# <a name="supertip"></a><span data-ttu-id="1eb04-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="1eb04-101">Supertip</span></span>

<span data-ttu-id="1eb04-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="1eb04-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1eb04-104">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1eb04-104">Child elements</span></span>

|  <span data-ttu-id="1eb04-105">Элемент</span><span class="sxs-lookup"><span data-stu-id="1eb04-105">Element</span></span> |  <span data-ttu-id="1eb04-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1eb04-106">Required</span></span>  |  <span data-ttu-id="1eb04-107">Описание</span><span class="sxs-lookup"><span data-stu-id="1eb04-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1eb04-108">Название</span><span class="sxs-lookup"><span data-stu-id="1eb04-108">Title</span></span>](#title)        | <span data-ttu-id="1eb04-109">Да</span><span class="sxs-lookup"><span data-stu-id="1eb04-109">Yes</span></span> |   <span data-ttu-id="1eb04-110">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="1eb04-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="1eb04-111">Описание</span><span class="sxs-lookup"><span data-stu-id="1eb04-111">Description</span></span>](#description)  | <span data-ttu-id="1eb04-112">Да</span><span class="sxs-lookup"><span data-stu-id="1eb04-112">Yes</span></span> |  <span data-ttu-id="1eb04-113">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="1eb04-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="1eb04-114">Title</span><span class="sxs-lookup"><span data-stu-id="1eb04-114">Title</span></span>

<span data-ttu-id="1eb04-p102">Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="1eb04-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="1eb04-118">Описание</span><span class="sxs-lookup"><span data-stu-id="1eb04-118">Description</span></span>

<span data-ttu-id="1eb04-p103">Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="1eb04-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="1eb04-122">Пример</span><span class="sxs-lookup"><span data-stu-id="1eb04-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
