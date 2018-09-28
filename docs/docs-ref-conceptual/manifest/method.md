# <a name="method-element"></a><span data-ttu-id="4b289-101">Элемент Method</span><span class="sxs-lookup"><span data-stu-id="4b289-101">Method element</span></span>

<span data-ttu-id="4b289-102">Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4b289-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="4b289-103">**Тип надстройки:** контентные надстройки и надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="4b289-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="4b289-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4b289-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="4b289-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="4b289-105">Contained in</span></span>

[<span data-ttu-id="4b289-106">Методы</span><span class="sxs-lookup"><span data-stu-id="4b289-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="4b289-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b289-107">Attributes</span></span>

|<span data-ttu-id="4b289-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="4b289-108">**Attribute**</span></span>|<span data-ttu-id="4b289-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="4b289-109">**Type**</span></span>|<span data-ttu-id="4b289-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="4b289-110">**Required**</span></span>|<span data-ttu-id="4b289-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="4b289-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4b289-112">Имя</span><span class="sxs-lookup"><span data-stu-id="4b289-112">Name</span></span>|<span data-ttu-id="4b289-113">string</span><span class="sxs-lookup"><span data-stu-id="4b289-113">string</span></span>|<span data-ttu-id="4b289-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4b289-114">required</span></span>|<span data-ttu-id="4b289-p101">Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.</span><span class="sxs-lookup"><span data-stu-id="4b289-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="4b289-117">Замечания</span><span class="sxs-lookup"><span data-stu-id="4b289-117">Remarks</span></span>

<span data-ttu-id="4b289-118">**Методы** и **метод** элементы не поддерживаются надстройки почты. Дополнительные сведения о наборах требований [версии Office и требования наборов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)см.</span><span class="sxs-lookup"><span data-stu-id="4b289-118">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="4b289-119">Так как нет возможности для определения требований к минимальной версии для отдельных методов, чтобы убедиться в том, что метод доступен во время выполнения, следует также использовать оператор **if** при вызове этого метода в скрипте надстройки.</span><span class="sxs-lookup"><span data-stu-id="4b289-119">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="4b289-120">Дополнительные сведения содержатся в разделе [Общие сведения об API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="4b289-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

