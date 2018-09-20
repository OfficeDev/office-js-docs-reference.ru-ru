# <a name="defaultsettings-element"></a><span data-ttu-id="3b392-101">Элемент DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="3b392-101">DefaultSettings element</span></span>

<span data-ttu-id="3b392-102">Задает исходное расположение по умолчанию и другие параметры по умолчанию для содержимого или надстройка области задач.</span><span class="sxs-lookup"><span data-stu-id="3b392-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="3b392-103">**Тип надстройки:** контентные надстройки и надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="3b392-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="3b392-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3b392-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="3b392-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="3b392-105">Contained in</span></span>

[<span data-ttu-id="3b392-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3b392-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3b392-107">Может содержать</span><span class="sxs-lookup"><span data-stu-id="3b392-107">Can contain</span></span>

|<span data-ttu-id="3b392-108">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="3b392-108">**Element**</span></span>|<span data-ttu-id="3b392-109">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="3b392-109">**Content**</span></span>|<span data-ttu-id="3b392-110">**Почта**</span><span class="sxs-lookup"><span data-stu-id="3b392-110">**Mail**</span></span>|<span data-ttu-id="3b392-111">**Область задач**</span><span class="sxs-lookup"><span data-stu-id="3b392-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3b392-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3b392-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="3b392-113">x</span><span class="sxs-lookup"><span data-stu-id="3b392-113">x</span></span>||<span data-ttu-id="3b392-114">x</span><span class="sxs-lookup"><span data-stu-id="3b392-114">x</span></span>|
|[<span data-ttu-id="3b392-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="3b392-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="3b392-116">x</span><span class="sxs-lookup"><span data-stu-id="3b392-116">x</span></span>|||
|[<span data-ttu-id="3b392-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="3b392-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="3b392-118">x</span><span class="sxs-lookup"><span data-stu-id="3b392-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="3b392-119">Замечания</span><span class="sxs-lookup"><span data-stu-id="3b392-119">Remarks</span></span>

<span data-ttu-id="3b392-120">Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие стандартные параметры с помощью элемента [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="3b392-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

