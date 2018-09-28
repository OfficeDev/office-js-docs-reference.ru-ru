# <a name="requirements-element"></a><span data-ttu-id="4beac-101">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="4beac-101">Requirements element</span></span>

<span data-ttu-id="4beac-102">Указывает минимальный набор элементов API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4beac-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="4beac-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="4beac-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4beac-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4beac-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="4beac-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="4beac-105">Contained in</span></span>

[<span data-ttu-id="4beac-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4beac-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4beac-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="4beac-107">Can contain</span></span>

|<span data-ttu-id="4beac-108">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="4beac-108">**Element**</span></span>|<span data-ttu-id="4beac-109">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="4beac-109">**Content**</span></span>|<span data-ttu-id="4beac-110">**Почта**</span><span class="sxs-lookup"><span data-stu-id="4beac-110">**Mail**</span></span>|<span data-ttu-id="4beac-111">**Область задач**</span><span class="sxs-lookup"><span data-stu-id="4beac-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4beac-112">Sets</span><span class="sxs-lookup"><span data-stu-id="4beac-112">Sets</span></span>](sets.md)|<span data-ttu-id="4beac-113">x</span><span class="sxs-lookup"><span data-stu-id="4beac-113">x</span></span>|<span data-ttu-id="4beac-114">x</span><span class="sxs-lookup"><span data-stu-id="4beac-114">x</span></span>|<span data-ttu-id="4beac-115">x</span><span class="sxs-lookup"><span data-stu-id="4beac-115">x</span></span>|
|[<span data-ttu-id="4beac-116">Методы</span><span class="sxs-lookup"><span data-stu-id="4beac-116">Methods</span></span>](methods.md)|<span data-ttu-id="4beac-117">x</span><span class="sxs-lookup"><span data-stu-id="4beac-117">x</span></span>||<span data-ttu-id="4beac-118">x</span><span class="sxs-lookup"><span data-stu-id="4beac-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="4beac-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="4beac-119">Remarks</span></span>

<span data-ttu-id="4beac-120">Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.</span><span class="sxs-lookup"><span data-stu-id="4beac-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

