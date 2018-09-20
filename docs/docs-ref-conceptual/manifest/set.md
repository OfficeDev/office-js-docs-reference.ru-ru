# <a name="set-element"></a><span data-ttu-id="d6e0b-101">Элемент Set</span><span class="sxs-lookup"><span data-stu-id="d6e0b-101">Set element</span></span>

<span data-ttu-id="d6e0b-102">Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="d6e0b-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="d6e0b-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="d6e0b-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d6e0b-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d6e0b-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="d6e0b-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="d6e0b-105">Contained in</span></span>

[<span data-ttu-id="d6e0b-106">Sets</span><span class="sxs-lookup"><span data-stu-id="d6e0b-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="d6e0b-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d6e0b-107">Attributes</span></span>

|<span data-ttu-id="d6e0b-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="d6e0b-108">**Attribute**</span></span>|<span data-ttu-id="d6e0b-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="d6e0b-109">**Type**</span></span>|<span data-ttu-id="d6e0b-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="d6e0b-110">**Required**</span></span>|<span data-ttu-id="d6e0b-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="d6e0b-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d6e0b-112">Имя</span><span class="sxs-lookup"><span data-stu-id="d6e0b-112">Name</span></span>|<span data-ttu-id="d6e0b-113">string</span><span class="sxs-lookup"><span data-stu-id="d6e0b-113">string</span></span>|<span data-ttu-id="d6e0b-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d6e0b-114">required</span></span>|<span data-ttu-id="d6e0b-115">Имя [набора требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d6e0b-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="d6e0b-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="d6e0b-116">MinVersion</span></span>|<span data-ttu-id="d6e0b-117">string</span><span class="sxs-lookup"><span data-stu-id="d6e0b-117">string</span></span>|<span data-ttu-id="d6e0b-118">необязательный</span><span class="sxs-lookup"><span data-stu-id="d6e0b-118">optional</span></span>|<span data-ttu-id="d6e0b-p101">Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="d6e0b-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d6e0b-121">Замечания</span><span class="sxs-lookup"><span data-stu-id="d6e0b-121">Remarks</span></span>

<span data-ttu-id="d6e0b-122">Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.</span><span class="sxs-lookup"><span data-stu-id="d6e0b-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="d6e0b-123">Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="d6e0b-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="d6e0b-124">Для надстроек почты существует только один `"Mailbox"` набору требований.</span><span class="sxs-lookup"><span data-stu-id="d6e0b-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="d6e0b-125">Этот набор требований содержит всей подмножество API, поддерживаемые в надстройках почты для Outlook, и необходимо указать `"Mailbox"` требований в почты надстроек в его манифесте (это не необязательно как в случае содержимого и задач надстроек области).</span><span class="sxs-lookup"><span data-stu-id="d6e0b-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="d6e0b-126">Кроме того невозможно объявить поддержку для отдельных методов в надстройках почты.</span><span class="sxs-lookup"><span data-stu-id="d6e0b-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
