# <a name="requirements-element"></a><span data-ttu-id="e2ecf-101">Элемент Requirements</span><span class="sxs-lookup"><span data-stu-id="e2ecf-101">Requirements element</span></span>

<span data-ttu-id="e2ecf-102">Указывает минимальный набор элементов API JavaScript для Office ([набор требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) и/или методов), необходимых для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="e2ecf-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="e2ecf-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="e2ecf-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e2ecf-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e2ecf-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="e2ecf-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="e2ecf-105">Contained in</span></span>

[<span data-ttu-id="e2ecf-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e2ecf-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e2ecf-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="e2ecf-107">Can contain</span></span>

|<span data-ttu-id="e2ecf-108">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="e2ecf-108">**Element**</span></span>|<span data-ttu-id="e2ecf-109">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="e2ecf-109">**Content**</span></span>|<span data-ttu-id="e2ecf-110">**Почта**</span><span class="sxs-lookup"><span data-stu-id="e2ecf-110">**Mail**</span></span>|<span data-ttu-id="e2ecf-111">**Область задач**</span><span class="sxs-lookup"><span data-stu-id="e2ecf-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e2ecf-112">Sets</span><span class="sxs-lookup"><span data-stu-id="e2ecf-112">Sets</span></span>](sets.md)|<span data-ttu-id="e2ecf-113">x</span><span class="sxs-lookup"><span data-stu-id="e2ecf-113">x</span></span>|<span data-ttu-id="e2ecf-114">x</span><span class="sxs-lookup"><span data-stu-id="e2ecf-114">x</span></span>|<span data-ttu-id="e2ecf-115">x</span><span class="sxs-lookup"><span data-stu-id="e2ecf-115">x</span></span>|
|[<span data-ttu-id="e2ecf-116">Методы</span><span class="sxs-lookup"><span data-stu-id="e2ecf-116">Methods</span></span>](methods.md)|<span data-ttu-id="e2ecf-117">x</span><span class="sxs-lookup"><span data-stu-id="e2ecf-117">x</span></span>||<span data-ttu-id="e2ecf-118">x</span><span class="sxs-lookup"><span data-stu-id="e2ecf-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="e2ecf-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="e2ecf-119">Remarks</span></span>

<span data-ttu-id="e2ecf-120">Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.</span><span class="sxs-lookup"><span data-stu-id="e2ecf-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

